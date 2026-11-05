"""Synthetic regressions. No supplier PDFs or corporate templates in Git."""
from contextlib import contextmanager
from datetime import date, datetime
import json
from types import SimpleNamespace

import pymupdf
import pytest
from openpyxl.utils.datetime import CALENDAR_MAC_1904, CALENDAR_WINDOWS_1900, to_excel

from cadastro_pcm.core import date_diagnostics
from cadastro_pcm.core.excel_date import brazilian_date_format, calendar_date, inspect_com_date, require_com_date, validate_date
from cadastro_pcm.core.excel_writer import fingerprint, verify_request_date, write_excel
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.pdf_reader import Page, read_pdf
from cadastro_pcm.core.table_headers import field_for, find_header
from cadastro_pcm.core.template_reader import read_template
from cadastro_pcm.interface import main_window
from test_prompt04 import extended_template, gui, wait_job


@pytest.mark.parametrize('epoch', [CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904])
@pytest.mark.parametrize('day', [date(2026, 11, 5), date(2028, 2, 29)])
def test_date_serial_is_validated_as_calendar_date(epoch, day):
    serial = to_excel(day, epoch)
    assert calendar_date(serial, epoch) == day
    validate_date(serial, 'dd/mm/yyyy;@', day, epoch)
    cell = SimpleNamespace(Value2=serial, NumberFormat='[$-pt-BR]dd/mm/yyyy;@',
                           NumberFormatLocal='dd/mm/aaaa', Text=day.strftime('%d/%m/%Y'))
    report = inspect_com_date(cell, day, epoch)
    assert report['date_ok'] and report['format_ok'] and report['display'] == 'ok'
    require_com_date(report)
    cell.Value2 = serial + 1
    with pytest.raises(ValueError, match='data armazenada'): require_com_date(inspect_com_date(cell, day, epoch))


@pytest.mark.parametrize('fmt', ['dd/mm/yyyy', 'DD/MM/YYYY;@', '[$-416]dd/mm/yyyy',
                                '[$-pt-BR]dd/mm/yyyy;@', 'dd"/"mm"/"yyyy', r'dd\/mm\/yyyy', 'dd/mm/aaaa'])
def test_equivalent_brazilian_date_formats(fmt):
    assert brazilian_date_format(fmt)


@pytest.mark.parametrize('fmt', ['General', 'mm/dd/yyyy', 'dd/mm/yy', 'yyyy-mm-dd', '"dd/mm/yyyy"'])
def test_incorrect_or_text_only_format_still_blocked(fmt):
    assert not brazilian_date_format(fmt)
    with pytest.raises(ValueError, match='formato brasileiro'):
        validate_date(46331, fmt, date(2026, 11, 5), CALENDAR_WINDOWS_1900)


def test_display_width_does_not_replace_value_validation():
    day = date(2026, 11, 5)
    cell = SimpleNamespace(Value2=to_excel(day), NumberFormat='dd/mm/yyyy', NumberFormatLocal='dd/mm/aaaa', Text='########')
    report = inspect_com_date(cell, day, CALENDAR_WINDOWS_1900)
    assert report['display'] == 'largura insuficiente'
    require_com_date(report)
    cell.Value2 += 1
    with pytest.raises(ValueError, match='data armazenada'): require_com_date(inspect_com_date(cell, day, CALENDAR_WINDOWS_1900))
    cell.Value2 -= 1; cell.Text = '46331'
    with pytest.raises(ValueError, match='exibição'): require_com_date(inspect_com_date(cell, day, CALENDAR_WINDOWS_1900))


def test_saved_localized_format_and_serial_survive_reopening(extended_template, tmp_path):
    # A localized format may leave a numeric value in an OOXML reader; both
    # representations must be validated using the workbook's date system.
    wb = read_template(extended_template); ws = wb.worksheets[0]
    ws['E8'] = to_excel(date(2026, 11, 5), wb.epoch)
    ws['E8'].number_format = 'dd/mm/aaaa;@'
    path = tmp_path / 'localized.xlsx'; wb.save(path); wb.close()
    verify_request_date(path, 'Criação de Material Nacional', date(2026, 11, 5))
    with pytest.raises(ValueError, match='data armazenada'):
        verify_request_date(path, 'Criação de Material Nacional', date(2026, 11, 6))


def table(prices=('1.422,00', '11.376,00')):
    return [
        ['4. CONDIÇÕES COMERCIAIS', None, None, None, None, None, None, None],
        ['Item', 'NCM', 'Descrição', 'Quantidade', 'IPI', 'ICMS', 'Valor', 'Valor'],
        ['', '', '', '', '', '', 'Unitário', 'Total'],
        ['', '', '', '', '', '', 'C/ Impostos', 'C/ Impostos'],
        ['2.1.', '7326.90.90', 'MANIPULO COM TRAVA', '8,0', '5%', '12%', prices[0], prices[1]],
        ['2.2.', '', 'Anel de vedação', '2,0', '0%', '12%', '', '20,00'],
        ['Prazo de entrega: 25 dias úteis', '', '', '', '', '', '', ''],
        ['Pagamento via boleto', '', '', '', '', '', '999,00', '999,00'],
    ]


@pytest.mark.parametrize('unit', ['1.422,00', 'R$ 1.422,00', '1422,00'])
def test_split_headers_commercial_section_and_prices_without_currency(unit):
    data = table((unit, '11.376,00'))
    header = find_header(data)
    assert header[1] == 3
    items = extract_items([Page('sintetico', 3, '', [data])])
    assert len(items) == 2
    first, second = items
    assert first.name == 'MANIPULO COM TRAVA' and first.ncm == '7326.90.90'
    assert first.unit_price == 1422 and first.quantity == 8 and first.total_price == 11376
    assert 'C/ Impostos' in first.evidence['unit_price'].excerpt
    assert second.name == 'Anel de vedação' and second.unit_price is None and not second.ncm


@pytest.mark.parametrize('header', ['Valor Unitário C/ Impostos', 'Valor\nUnitário\nC/ Impostos',
                                   'Preço unitário com impostos', 'R$ Unit.'])
def test_monetary_header_variants(header):
    assert field_for(header) == 'unit_price'
    assert field_for('IPI 5%') is None and field_for('ICMS (%)') is None


def test_total_and_percentages_never_fill_missing_unit_price():
    data = table(('', '11.376,00'))
    item = extract_items([Page('sintetico', 3, '', [data])])[0]
    assert item.unit_price is None and item.total_price == 11376
    data[4][6] = '5%'
    assert extract_items([Page('sintetico', 3, '', [data])])[0].unit_price is None


def test_scope_join_requires_unique_explicit_item_reference_in_same_pdf():
    scope = '3. ESCOPO DE FORNECIMENTO\n2.1. Conjunto de manípulo e trava para aperto de tampos.\nCorpo em aço inox AISI 304 polido.\nBucha roscada de bronze.\nAltura de fechamento de 28,5 mm.\n4. CONDIÇÕES COMERCIAIS\n25 dias úteis após aprovação'
    page = Page('proposal-a', 3, scope, [table()])
    item = extract_items([page])[0]
    assert item.description.startswith('Conjunto de manípulo')
    assert 'AISI 304' in item.description and '28,5 mm' in item.description
    assert '25 dias' not in item.description
    assert 'página 3' in item.evidence['description'].source
    other = Page('proposal-b', 1, scope, [])
    unmatched = extract_items([Page('proposal-a', 3, '', [table()]), other])[0]
    assert unmatched.description == unmatched.name
    duplicate = Page('proposal-a', 2, scope, [])
    ambiguous = extract_items([page, duplicate])[0]
    assert ambiguous.description == ambiguous.name
    assert any('referência repetida' in warning for warning in ambiguous.warnings)


def five_page_pdf(path, ruled=True):
    with pymupdf.open() as doc:
        for index in range(5):
            page = doc.new_page(width=842, height=595)
            page.insert_text((30, 25), f'Proposta sintetica - pagina {index + 1}')
            if index != 2: continue
            page.insert_text((30, 55), '3. ESCOPO DE FORNECIMENTO\n2.1. Conjunto de manipulo e trava\nCorpo em aco inox AISI 304 polido.\nBucha roscada de bronze.\nAltura de fechamento de 28,5 mm.')
            # Decorative ruled box must not disable borderless product detection.
            page.draw_rect(pymupdf.Rect(30, 150, 300, 180))
            page.insert_text((35, 167), '4. CONDIÇÕES COMERCIAIS')
            xs = [30, 70, 150, 375, 445, 485, 530, 650, 800]
            rows = table()[1:6]
            for row, cells in enumerate(rows):
                for x, value in zip(xs, cells):
                    if value: page.insert_text((x + 3, 225 + row * 20), value, fontsize=8)
            if ruled:
                for y in range(210, 311, 20): page.draw_line((30, y), (800, y))
                for x in xs: page.draw_line((x, 210), (x, 310))
            page.insert_text((30, 365), 'Prazo de entrega: 25 dias uteis\nPagamento: boleto bancario\nValor total do orcamento: 11.396,00')
        doc.save(path)
    return path


@pytest.mark.parametrize('ruled', [True, False])
def test_real_pymupdf_positions_five_pages_and_decorative_tables(tmp_path, ruled):
    path = five_page_pdf(tmp_path / 'five-pages.pdf', ruled)
    before = path.read_bytes(); pages = read_pdf(path)
    assert len(pages) == 5
    items = extract_items(pages)
    assert len(items) == 2
    first, second = items
    assert first.name == 'MANIPULO COM TRAVA' and first.ncm == '7326.90.90'
    assert first.unit_price == 1422 and first.total_price == 11376 and first.quantity == 8
    assert second.name == 'Anel de vedação' and second.unit_price is None
    assert 'pagina 3' not in first.description and 'AISI 304' in first.description
    assert path.read_bytes() == before


def test_five_page_proposal_gui_approval_to_excel(gui, tmp_path):
    app, window, messages = gui
    source = five_page_pdf(tmp_path / 'proposal.pdf', ruled=False)
    before = fingerprint(window.settings.template)
    window.add_paths([str(source)]); window.process_button.click(); wait_job(app, window)
    assert len(window.items) == 2
    item = window.items[0]
    editor = main_window.Editor(item, window)
    assert editor.inputs['ncm'].text() == '7326.90.90'
    assert editor.inputs['unit_price'].text() == '1422.00'
    assert 'AISI 304' in editor.inputs['description'].toPlainText()
    editor.confirm.setChecked(True); editor.save(item)
    window.items[0] = editor.result_item; window.refresh()
    assert window.excel_button.isEnabled()
    window.excel_button.click(); wait_job(app, window)
    wb = read_template(window.outputs[0] / 'cadastro.xlsx'); ws = wb.worksheets[0]
    assert ws['A12'].value == 'MANIPULO COM TRAVA' and ws['A20'].value == 1422
    assert 'AISI 304' in ws['A14'].value and ws['A14'].value.endswith('NCM: 7326.90.90')
    assert isinstance(ws['E8'].value, datetime) and ws['E8'].number_format == 'dd/mm/yyyy'
    assert ws['C16'].value is None
    assert fingerprint(window.settings.template) == before
    wb.close(); editor.deleteLater()


def test_windows_diagnostic_is_explicit_and_does_not_claim_excel_on_linux(tmp_path, monkeypatch):
    monkeypatch.setattr(date_diagnostics, 'sys', SimpleNamespace(platform='linux'))
    with pytest.raises(RuntimeError, match='Excel real'):
        date_diagnostics.diagnose_date('missing.xlsx', 'sheet', date.today(), tmp_path)
    assert not list(tmp_path.iterdir())


@pytest.mark.parametrize('saved_delta', [0, 1])
def test_windows_diagnostic_reopens_reports_e8_only_and_preserves_original(extended_template, tmp_path, monkeypatch, saved_delta):
    before = extended_template.read_bytes(); sessions = []
    @contextmanager
    def session(path, read_only):
        sessions.append(read_only)
        wb = read_template(path); ws = wb.worksheets[0]
        class Cell:
            HasFormula = False; Locked = False
            @property
            def Value2(self):
                value = ws['E8'].value
                return to_excel(value, wb.epoch) if isinstance(value, datetime) else value
            @Value2.setter
            def Value2(self, value): ws['E8'] = value
            @property
            def NumberFormat(self): return ws['E8'].number_format
            @NumberFormat.setter
            def NumberFormat(self, value): ws['E8'].number_format = value + ';@'
            NumberFormatLocal = 'dd/mm/aaaa'
            @property
            def Text(self): return ws['E8'].value.strftime('%d/%m/%Y')
        sheet = SimpleNamespace(ProtectContents=False, Range=lambda addr: Cell())
        def save():
            ws['E8'] = ws['E8'].value + saved_delta
            wb.save(path)
        yield None, SimpleNamespace(ReadOnly=read_only, Worksheets=lambda name: sheet, Save=save)
        wb.close()
    monkeypatch.setattr(date_diagnostics, 'sys', SimpleNamespace(platform='win32'))
    monkeypatch.setattr(date_diagnostics, 'excel_session', session)
    report, path = date_diagnostics.diagnose_date(extended_template, 'Criação de Material Nacional', date(2026, 11, 5), tmp_path / 'report')
    assert report['ok'] == (saved_delta == 0)
    assert report['source_unchanged'] and sessions == [False, True]
    assert report['after_reopen']['value2'] == 46331 + saved_delta
    assert report['after_reopen']['date_ok'] == (saved_delta == 0)
    assert report['after_reopen']['format_ok']
    if saved_delta: assert 'data armazenada' in report['error']
    else: assert report['after_reopen']['display'] == 'ok'
    assert extended_template.read_bytes() == before and path.is_file()
    assert (path.parent / report['copy']).is_file()
    assert 'Parafuso' not in json.dumps(report)
