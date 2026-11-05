"""Local date diagnostics and regional display checks, with explicit COM fakes."""
from contextlib import contextmanager
from datetime import date, datetime
import json
from pathlib import Path
from types import SimpleNamespace

import pymupdf
import pytest
from openpyxl import Workbook
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, from_excel, to_excel

from cadastro_pcm.core import excel_writer, pdf_exporter
from cadastro_pcm.core.excel_date import (DATE_FORMAT, PDF_DATE_FORMAT, DateValidationError,
    check_com_date, inspect_com_date, require_com_date)
from cadastro_pcm.core.template_reader import read_template
from test_prompt04 import complete_item, extended_template

EXPECTED = date(2026, 11, 5)
EPOCH = CALENDAR_WINDOWS_1900


@pytest.fixture(autouse=True)
def private_reports(monkeypatch, tmp_path):
    monkeypatch.setenv('LOCALAPPDATA', str(tmp_path / 'local'))


class DateCell:
    MergeCells = False

    def __init__(self, text='05/11/2026', local='dd/mm/aaaa', width=100, render=True):
        self.Value2 = to_excel(EXPECTED)
        self._format = DATE_FORMAT; self._local = local; self._text = text
        self._width = width; self._shrink = False; self.render = render
        self.changes = []

    @property
    def Value(self): return from_excel(self.Value2)
    @property
    def NumberFormat(self): return self._format
    @NumberFormat.setter
    def NumberFormat(self, value):
        self.changes.append(('NumberFormat', value)); self._format = value
        if value == PDF_DATE_FORMAT and self.render:
            self._local = 'dd"/"mm"/"aaaa'
            if set(self._text) != {'#'}: self._text = EXPECTED.strftime('%d/%m/%Y')
    @property
    def NumberFormatLocal(self): return self._local
    @property
    def Text(self): return self._text
    @property
    def Width(self): return self._width
    @property
    def ColumnWidth(self): return 15
    @property
    def ShrinkToFit(self): return self._shrink
    @ShrinkToFit.setter
    def ShrinkToFit(self, value):
        self.changes.append(('ShrinkToFit', value)); self._shrink = value
        if value and self._width >= 55: self._text = EXPECTED.strftime('%d/%m/%Y')


@pytest.mark.parametrize('text,local,order', [
    ('05.11.2026', 'dd.mm.yyyy', 'dmy'),
    ('5/11/2026', 'd/m/yyyy', 'dmy'),
    ('11/5/2026', 'm/d/yyyy', 'mdy'),
    ('2026-11-05', 'yyyy-mm-dd', 'ymd'),
    ('\u200e05/11/2026\u200f', 'dd/mm/aaaa', 'dmy'),
    ('٠٥/١١/٢٠٢٦', 'dd/mm/aaaa', 'dmy'),
    ('05.11.2026', 'TT.MM.JJJJ', 'dmy'),
    ('5/11/2026', 'j/m/aaaa', 'dmy'),
])
def test_regional_text_requires_declared_order_and_correct_value(tmp_path, text, local, order):
    cell = DateCell(text, local)
    report = check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)
    assert report['date_ok'] and report['value_ok'] and report['format_ok']
    assert report['display'] == 'regional confirmada' and report['local_date_order'] == order
    assert cell.Value2 == to_excel(EXPECTED) and not cell.changes
    saved = json.loads(next(tmp_path.glob('diagnostico_data_*.json')).read_text())
    assert saved['display_ok'] and saved['attempts'][0]['text'] == text


def test_does_not_guess_us_order_when_local_format_is_brazilian(tmp_path):
    cell = DateCell('11/05/2026', 'dd/mm/aaaa')
    with pytest.raises(DateValidationError, match='exibição') as error:
        check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path, repair=True)
    report = json.loads(error.value.report_path.read_text())
    state = report['attempts'][-1]
    assert report['date_ok'] and report['format_ok'] and not report['display_ok']
    assert state['value2'] == 46331 and state['value'] == '2026-11-05T00:00:00'
    assert state['cell_width_points'] == 100 and state['column_width_chars'] == 15
    assert not cell.changes


def test_invariant_format_provides_order_when_local_mask_is_unknown(tmp_path):
    cell = DateCell('05.11.2026', 'máscara não reconhecida')
    report = check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)
    assert report['display_ok'] and report['local_date_order'] is None
    assert report['display_date_order'] == 'dmy' and report['display_order_source'] == 'NumberFormat'
    cell._text = '11.05.2026'
    with pytest.raises(DateValidationError):
        check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)


@pytest.mark.parametrize('failure', ['date', 'format', 'unexplained'])
def test_local_report_exists_before_date_error(tmp_path, failure):
    cell = DateCell()
    if failure == 'date': cell.Value2 += 1
    elif failure == 'format': cell._format = 'General'
    else: cell._text = '46331'
    with pytest.raises(DateValidationError) as error:
        check_com_date(cell, EXPECTED, EPOCH, phase='antes de salvar', report_dir=tmp_path, repair=True)
    report = json.loads(error.value.report_path.read_text())
    assert report['cell'] == 'E8' and report['phase'] == 'antes de salvar'
    assert {'value', 'value2', 'text', 'number_format', 'number_format_local',
            'cell_width_points', 'expected_date', 'date_ok', 'format_ok', 'display_ok'} <= report['attempts'][0].keys()
    assert not cell.changes


def test_reports_redact_non_date_values_and_personal_literals(tmp_path):
    cell = SimpleNamespace(Value='PESSOA CONFIDENCIAL', Value2='PESSOA CONFIDENCIAL',
                           Text='PESSOA CONFIDENCIAL', NumberFormat='dd/mm/yyyy "PESSOA CONFIDENCIAL"',
                           NumberFormatLocal='dd/mm/aaaa "PESSOA CONFIDENCIAL"', Width=100, ColumnWidth=15)
    with pytest.raises(DateValidationError) as error:
        check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)
    assert 'PESSOA CONFIDENCIAL' not in error.value.report_path.read_text()
    assert 'PESSOA CONFIDENCIAL' not in str(error.value)
    cell.NumberFormat = 'dd/mm/yyyy;"PESSOA CONFIDENCIAL"'
    cell.NumberFormatLocal = '[$-PESSOA CONFIDENCIAL]dd/mm/aaaa'
    with pytest.raises(DateValidationError) as error:
        check_com_date(cell, EXPECTED, EPOCH, phase='após reabrir', report_dir=tmp_path)
    assert 'PESSOA CONFIDENCIAL' not in error.value.report_path.read_text()


def test_hashes_repaired_only_in_cell_without_resizing_columns(tmp_path):
    cell = DateCell('########', width=80)
    result = check_com_date(cell, EXPECTED, EPOCH, phase='cópia', report_dir=tmp_path, repair=True)
    assert result['display'] == 'ok' and cell.ShrinkToFit
    assert cell.Width == 80 and cell.ColumnWidth == 15 and cell.Value2 == 46331
    assert all(name in ('NumberFormat', 'ShrinkToFit') for name, _ in cell.changes)
    report = json.loads(next(tmp_path.glob('diagnostico_data_*.json')).read_text())
    assert report['attempts'][0]['display'] == 'largura insuficiente'
    assert report['attempts'][-1]['display_ok']


def test_too_narrow_cell_is_not_blindly_shrunk(tmp_path):
    cell = DateCell('########', width=20)
    with pytest.raises(DateValidationError, match='largura insuficiente'):
        check_com_date(cell, EXPECTED, EPOCH, phase='cópia', report_dir=tmp_path, repair=True)
    assert not cell.ShrinkToFit and cell.Width == 20 and cell.ColumnWidth == 15


def test_pdf_requires_literal_slashes_even_for_confirmed_regional_display(tmp_path):
    cell = DateCell('11/5/2026', 'm/d/yyyy')
    assert inspect_com_date(cell, EXPECTED, EPOCH)['display'] == 'regional confirmada'
    report = check_com_date(cell, EXPECTED, EPOCH, phase='PDF', report_dir=tmp_path, repair=True, exact=True)
    assert report['display'] == 'ok' and cell.NumberFormat == PDF_DATE_FORMAT
    stuck = DateCell('05.11.2026', 'dd.mm.yyyy', render=False)
    with pytest.raises(DateValidationError, match='barras'):
        check_com_date(stuck, EXPECTED, EPOCH, phase='PDF', report_dir=tmp_path, repair=True, exact=True)


def test_native_writer_saves_and_reopens_confirmed_regional_date(extended_template, tmp_path, monkeypatch):
    calls = []
    @contextmanager
    def session(path, read_only):
        calls.append(read_only)
        wb = read_template(path); ws = wb.worksheets[0]
        class Cell:
            HasFormula = False; Locked = False; MergeCells = False
            Width = 100; ColumnWidth = 15
            def __init__(self, address): self.cell = ws[address]; self.address = address
            @property
            def Value(self): return self.cell.value
            @Value.setter
            def Value(self, value): self.cell.value = value
            @property
            def Value2(self):
                return to_excel(self.cell.value, wb.epoch) if isinstance(self.cell.value, datetime) else self.cell.value
            @Value2.setter
            def Value2(self, value): self.cell.value = value
            @property
            def NumberFormat(self): return self.cell.number_format
            @NumberFormat.setter
            def NumberFormat(self, value): self.cell.number_format = value
            NumberFormatLocal = 'm/d/yyyy'
            @property
            def Text(self): return '11/5/2026'
        sheet = SimpleNamespace(ProtectContents=False, Range=Cell,
                                Cells=SimpleNamespace(SpecialCells=lambda _: object()),
                                Application=SimpleNamespace(Intersect=lambda a, b: None))
        yield None, SimpleNamespace(ReadOnly=read_only, Worksheets=lambda _: sheet, Save=lambda: wb.save(path))
        wb.close()
    monkeypatch.setattr(excel_writer, 'excel_session', session)
    monkeypatch.setattr(excel_writer, 'use_native_excel', lambda: True)
    original = extended_template.read_bytes()
    item = complete_item(); item.request_date = EXPECTED
    result = tmp_path / 'saved.xlsx'
    excel_writer.write_excel(extended_template, result, item, 'Criação de Material Nacional', excel_writer.fingerprint(extended_template))
    assert calls == [False, True] and extended_template.read_bytes() == original
    wb = read_template(result)
    assert wb.worksheets[0]['E8'].value == datetime(2026, 11, 5)
    assert wb.worksheets[0]['E8'].number_format == DATE_FORMAT
    wb.close()


@pytest.mark.parametrize('render', [True, False])
def test_pdf_copy_has_readable_date_and_original_dimensions(tmp_path, monkeypatch, render):
    source = tmp_path / 'source.xlsx'; wb = Workbook(); ws = wb.active
    ws.title = 'Criação de Material Nacional'; ws['A12'] = 'Produto'
    ws['A33'] = 'Controle institucional'; ws['E8'] = datetime(2026, 11, 5)
    ws['E8'].number_format = DATE_FORMAT; wb.save(source); wb.close()
    original = source.read_bytes(); target = tmp_path / 'final.pdf'
    cell = DateCell('05.11.2026', 'dd.mm.yyyy', render=render)
    exports = []
    class Sheet:
        ProtectContents = True
        PageSetup = SimpleNamespace(LeftMargin=36, RightMargin=36, TopMargin=36, BottomMargin=36)
        def Range(self, address):
            if address == 'E8': return cell
            if ':' in address: return SimpleNamespace(Width=500, Height=700)
            return SimpleNamespace(Value=None, Text={'A12': 'Produto', 'A33': 'Controle institucional'}.get(address, ''))
        def ResetAllPageBreaks(self): raise RuntimeError('Protected')
        def ExportAsFixedFormat(self, *args):
            exports.append(True)
            assert self.PageSetup.FitToPagesTall == self.PageSetup.FitToPagesWide == 1
            with pymupdf.open() as doc:
                page = doc.new_page(width=595.28, height=841.89)
                page.insert_text((30, 30), 'Produto\n' + cell.Text + '\nControle institucional', fontsize=10)
                doc.save(target)
    @contextmanager
    def session(path, read_only):
        assert Path(path) != source and not read_only
        yield None, SimpleNamespace(Worksheets=lambda _: Sheet())
    monkeypatch.setattr(pdf_exporter, 'excel_session', session)
    monkeypatch.setattr(pdf_exporter, 'sys', SimpleNamespace(platform='win32'))
    if render:
        pdf_exporter.export_pdf(source, target, 'Criação de Material Nacional', 'Produto')
        assert exports
        with pymupdf.open(target) as doc: assert len(doc) == 1 and '05/11/2026' in doc[0].get_text()
    else:
        with pytest.raises(DateValidationError):
            pdf_exporter.export_pdf(source, target, 'Criação de Material Nacional', 'Produto')
        assert not exports and not target.exists()
    assert source.read_bytes() == original and cell.Width == 100 and cell.ColumnWidth == 15


def test_pdf_date_is_not_accepted_at_illegible_size(tmp_path):
    target = tmp_path / 'tiny-date.pdf'
    with pymupdf.open() as doc:
        page = doc.new_page(); page.insert_text((30, 30), '05/11/2026', fontsize=6)
        doc.save(target)
    with pytest.raises(ValueError, match='8 pontos'):
        pdf_exporter.validate_pdf(target, '', require_a4=True, expected_date='05/11/2026')


def test_independent_readback_logs_without_claiming_com_observations(tmp_path):
    source = tmp_path / 'wrong-date.xlsx'; wb = Workbook(); ws = wb.active
    ws['E8'] = datetime(2026, 11, 6); ws['E8'].number_format = DATE_FORMAT
    wb.save(source); wb.close(); original = source.read_bytes()
    with pytest.raises(DateValidationError, match='data armazenada') as error:
        excel_writer.verify_request_date(source, 'Sheet', EXPECTED)
    report = json.loads(error.value.report_path.read_text())
    assert not report['date_ok'] and report['format_ok'] and report['display_ok'] is None
    assert report['attempts'][0]['value2'] is None and report['attempts'][0]['text'] is None
    assert source.read_bytes() == original
