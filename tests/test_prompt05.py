"""Synthetic tests for conservative item recognition, dates and native PDF flow."""
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pymupdf
import pytest
from openpyxl import Workbook
from openpyxl.styles import Protection
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from PySide6.QtWidgets import QComboBox, QDialog

from cadastro_pcm.core import excel_writer, pdf_exporter
from cadastro_pcm.core.excel_writer import check_template, fingerprint, write_excel
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.models import Evidence, Item
from cadastro_pcm.core.pdf_reader import Page
from cadastro_pcm.core.print_layout import check_print_scale
from cadastro_pcm.core.template_reader import read_template
from cadastro_pcm.interface import main_window
from test_prompt04 import complete_item, extended_template, gui, wait_job


@pytest.mark.parametrize('condition', [
    '25 dias úteis, após aprovação ou recebimento do pedido de compra.',
    '7 semanas após recebimento da ordem', '30/60/90 dias',
    'Prazo de entrega: até 8 dias corridos', 'Prazo: 12 dias',
    'Condição de pagamento: boleto em 15 dias', 'Pagamento via boleto',
    'À vista com desconto', 'Pagamento antecipado de 50%', 'Boleto 28 dias',
    'Transferência bancária após aprovação', 'Frete incluso', 'Transporte: FOB',
    'Banco: conta corrente 1234', 'Chave PIX: 12345678',
    'Valor total do orçamento: R$ 9.999,00', 'Entrega prevista para dezembro',
])
def test_commercial_conditions_never_become_items(condition):
    labels = Page('sintetico', 1, f'Descrição: {condition}\nPreço unitário: 900,00\nNCM: 7318.15.00', [])
    table = Page('sintetico', 1, '', [[
        ['Descrição', 'R$ Unit.', 'NCM'], [condition, '900,00', '7318.15.00']]])
    assert not extract_items([labels])
    assert not extract_items([table])


def test_commercial_sections_cannot_supply_product_prices():
    text = ('Produto: Parafuso M8\nDescrição técnica: Aço inoxidável\n'
            'Prazo de entrega: 3 semanas\nPreço unitário: 999,00\n'
            'Descrição: Recebimento após aprovação\nNCM: 9999.99.99\n'
            'Itens da cotação\nProduto: Anel de vedação\nValor unitário: 4,50\n')
    items = extract_items([Page('sintetico', 1, text, [])])
    assert len(items) == 2
    assert items[0].name == 'Parafuso M8' and items[0].unit_price is None and not items[0].ncm
    assert items[0].description == 'Aço inoxidável'
    assert items[1].name == 'Anel de vedação' and items[1].unit_price == 4.5


def test_table_section_boundary_and_price_association():
    page = Page('sintetico', 1, '', [[
        ['Descrição', 'Quantidade', 'R$ Unit.', 'R$ Total'],
        ['Parafuso M8', '3', '2,00', '6,00'],
        ['Anel', '2', '', '12,00'],
        ['CONDIÇÕES COMERCIAIS', '', '', ''],
        ['15 dias úteis após pedido', '1', '999,00', '999,00'],
        ['Descrição do pagamento', '', '20,00', '20,00'],
    ]])
    items = extract_items([page])
    assert [i.name for i in items] == ['Parafuso M8', 'Anel']
    assert items[0].unit_price == 2 and items[1].unit_price is None
    assert items[1].total_price == 12


def test_real_service_names_and_no_secondary_extraction():
    items = extract_items([Page('sintetico', 1, 'Serviço: Transporte rodoviário de carga\n'
                               'Descrição técnica: Trajeto entre unidades\n'
                               'Fabricante: não extrair\nCódigo do fabricante: 12345678\n'
                               'Preço unitário: 150,00', [])])
    assert len(items) == 1 and items[0].unit_price == 150
    assert items[0].name == 'Transporte rodoviário de carga'
    assert not items[0].manufacturer and not items[0].manufacturer_code and not items[0].ncm


@pytest.mark.parametrize('description', [
    'Parafuso M8\nFabricante: legado\nCódigo do fabricante: ABC\nNCM fiscal informado: 7318.15.00',
    'Parafuso M8\nNCM: 73181500\nPrazo: 3 semanas\nPagamento: boleto',
    'Parafuso M8 - Prazo de entrega: 14 dias',
    'Parafuso M8, entrega em 3 semanas',
])
def test_excel_technical_description_exact_label_and_no_commercial_fields(extended_template, description):
    item = complete_item(); item.description = description
    item.manufacturer = 'Dado antigo'; item.manufacturer_code = 'LEGADO'
    values, _ = check_template(extended_template, item, 'Criação de Material Nacional', fingerprint(extended_template))
    assert values['A14'] == 'Parafuso M8\nNCM: 7318.15.00'


def test_ncm_not_applicable_and_conflicting_description(extended_template):
    item = complete_item(); item.ncm = ''; item.ncm_not_applicable = True
    values, _ = check_template(extended_template, item, 'Criação de Material Nacional', fingerprint(extended_template))
    assert values['A14'] == item.description
    item.description += '\nNCM: 7318.15.00'
    with pytest.raises(ValueError, match='diverge'):
        check_template(extended_template, item, 'Criação de Material Nacional', fingerprint(extended_template))


def test_legacy_commercial_item_cannot_be_approved():
    item = complete_item(); item.name = '15 dias úteis após recebimento do pedido'
    with pytest.raises(ValueError, match='condição comercial'): item.require_ready()


@pytest.mark.parametrize('mac_epoch', [False, True])
def test_date_format_after_reopen_and_institutional_styles_preserved(extended_template, tmp_path, mac_epoch):
    source = extended_template
    if mac_epoch:
        # A separate synthetic template without extensions for the 1904 date system.
        source = tmp_path / 'epoch1904.xlsx'
        wb = Workbook(); ws = wb.active; ws.title = 'Criação de Material Nacional'
        ws['A33'] = 'Controle institucional'; ws['B4'] = datetime(2024, 2, 3)
        wb.epoch = CALENDAR_MAC_1904; wb.save(source)
    before = source.read_bytes()
    item = complete_item(); item.request_date = date(2027, 8, 19)
    target = tmp_path / 'copia.xlsx'
    write_excel(source, target, item, 'Criação de Material Nacional', fingerprint(source))
    assert source.read_bytes() == before
    a, b = read_template(source), read_template(target)
    assert b.worksheets[0]['E8'].value == datetime(2027, 8, 19)
    assert b.worksheets[0]['E8'].number_format == 'dd/mm/yyyy'
    assert b.worksheets[0]['B4'].value == a.worksheets[0]['B4'].value
    assert b.worksheets[0]['B4'].style_id == a.worksheets[0]['B4'].style_id
    a.close(); b.close()
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with ZipFile(source) as src, ZipFile(target) as dst:
        original = ET.fromstring(src.read('xl/styles.xml'))
        updated = ET.fromstring(dst.read('xl/styles.xml'))
        old_styles, new_styles = original.find(ns + 'cellXfs'), updated.find(ns + 'cellXfs')
        assert len(new_styles) >= len(old_styles)
        for old, new in zip(old_styles, new_styles): assert ET.tostring(old) == ET.tostring(new)
        for element in original:
            if element.tag not in (ns + 'numFmts', ns + 'cellXfs'):
                assert ET.tostring(element) == ET.tostring(updated.find(element.tag))


def test_native_writer_sets_date_format_and_reopens(extended_template, tmp_path, monkeypatch):
    output = tmp_path / 'native.xlsx'; formats = []
    @contextmanager
    def session(path, read_only=False):
        wb = read_template(path); sheet = wb.worksheets[0]
        class Cell:
            def __init__(self, address): self.cell = sheet[address]; self.address = address
            HasFormula = False; MergeCells = False; Locked = False
            @property
            def Value(self): return self.cell.value
            @Value.setter
            def Value(self, value): self.cell.value = value
            @property
            def NumberFormat(self): return self.cell.number_format
            @NumberFormat.setter
            def NumberFormat(self, value):
                formats.append((self.address, value)); self.cell.number_format = value
        ws = SimpleNamespace(ProtectContents=False, Range=Cell,
                             Cells=SimpleNamespace(SpecialCells=lambda kind: object()),
                             Application=SimpleNamespace(Intersect=lambda a, b: None))
        yield None, SimpleNamespace(ReadOnly=False, Worksheets=lambda name: ws, Save=lambda: wb.save(path))
        wb.close()
    monkeypatch.setattr(excel_writer, 'excel_session', session)
    monkeypatch.setattr(excel_writer, 'use_native_excel', lambda: True)
    item = complete_item(); item.request_date = date(2028, 2, 29)
    write_excel(extended_template, output, item, 'Criação de Material Nacional', fingerprint(extended_template))
    assert formats == [('E8', 'dd/mm/yyyy')]
    wb = read_template(output)
    assert wb.worksheets[0]['E8'].value == datetime(2028, 2, 29)
    wb.close()


def test_review_hides_secondary_fields_but_reads_legacy_json(gui):
    app, window, messages = gui
    item = Item.model_validate({**complete_item().model_dump(), 'manufacturer': 'LEGACY SUPPLIER',
                               'manufacturer_code': 'LEGACY CODE', 'quantity': 3, 'total_price': 55.5})
    item.evidence['manufacturer'] = Evidence(origin='manual', source='LEGACY SOURCE')
    window.items = [item]; window.refresh()
    editor = main_window.Editor(item, window)
    assert not {'manufacturer', 'manufacturer_code', 'quantity', 'total_price'} & editor.inputs.keys()
    assert {'name', 'description', 'ncm', 'unit_price', 'classification', 'requester'} <= editor.inputs.keys()
    assert 'Fabricante' not in window.details.toPlainText() and 'LEGACY' not in window.details.toPlainText()
    assert [window.table.horizontalHeaderItem(i).text() for i in range(window.table.columnCount())] == ['Material', 'Preço unitário', 'NCM', 'Origem', 'Validação']
    editor.confirm.setChecked(True); editor.save(item)
    assert editor.result_item.manufacturer == 'LEGACY SUPPLIER'  # Read compatibility only.
    editor.deleteLater()


def test_manual_research_keeps_only_main_fields(gui, monkeypatch):
    app, window, messages = gui
    window.items = [complete_item()]; window.refresh()
    captured = []
    def inspect(dialog):
        combo = dialog.findChild(QComboBox)
        captured.extend(combo.itemData(i) for i in range(combo.count()))
        return QDialog.Rejected
    monkeypatch.setattr(QDialog, 'exec', inspect)
    window.research()
    assert captured == ['ncm', 'unit_price', 'description']


def synthetic_pdf(path, texts, width=595.28, height=841.89, fontsize=10):
    with pymupdf.open() as doc:
        page = doc.new_page(width=width, height=height)
        for index, text in enumerate(texts): page.insert_text((40, 40 + index * 25), text, fontsize=fontsize)
        doc.save(path)


def test_pdf_validation_requires_date_all_fields_a4_and_readability(tmp_path):
    path = tmp_path / 'valid.pdf'
    synthetic_pdf(path, ['Parafuso M8', '19/08/2027', 'Controle institucional'])
    pdf_exporter.validate_pdf(path, 'Parafuso M8', expected_texts=[('A33', 'Controle institucional')], require_a4=True, expected_date='19/08/2027')
    with pytest.raises(ValueError, match='Data brasileira'):
        pdf_exporter.validate_pdf(path, 'Parafuso M8', expected_date='20/08/2027')
    with pytest.raises(ValueError, match='ausente ou cortado'):
        pdf_exporter.validate_pdf(path, 'Parafuso M8', expected_texts=[('A14', 'Descrição longa cortada')])
    letter = tmp_path / 'letter.pdf'; synthetic_pdf(letter, ['Parafuso M8'], width=612, height=792)
    with pytest.raises(ValueError, match='A4'): pdf_exporter.validate_pdf(letter, 'Parafuso M8', require_a4=True)
    tiny = tmp_path / 'tiny.pdf'; synthetic_pdf(tiny, ['Parafuso M8'], fontsize=4)
    with pytest.raises(ValueError, match='5 pontos'): pdf_exporter.validate_pdf(tiny, 'Parafuso M8', require_a4=True)
    with pytest.raises(ValueError, match='ausente ou cortado'):
        pdf_exporter.validate_pdf(path, 'Parafuso M8', expected_texts=[('A12', 'Parafuso M8'), ('A14', 'Parafuso M8')])


def test_date_policy_current_and_fixed(gui, monkeypatch):
    app, window, messages = gui
    class Clock(date):
        @classmethod
        def today(cls): return date(2030, 7, 21)
    monkeypatch.setattr(main_window, 'date', Clock)
    window.settings.date_policy = 'atual'
    item = Item(); window.apply_defaults(item)
    assert item.request_date == date(2030, 7, 21)
    window.settings.date_policy = 'fixa'; window.settings.fixed_date = date(2025, 1, 14)
    other = Item(); window.apply_defaults(other)
    assert other.request_date == date(2025, 1, 14)


def test_date_verification_rejects_general_format(tmp_path):
    path = tmp_path / 'serial.xlsx'; wb = Workbook(); ws = wb.active
    ws['E8'] = 46331; wb.save(path)
    with pytest.raises(ValueError, match='formato brasileiro'):
        excel_writer.verify_request_date(path, ws.title, date(2026, 11, 5))


def test_empty_destination_never_consumes_next_cell_or_institutional_label(tmp_path):
    source = tmp_path / 'template.xlsx'; wb = Workbook(); ws = wb.active
    ws.title = 'Criação de Material Nacional'; ws['A33'] = 'Controle institucional'
    for address in excel_writer.MAP.values(): ws[address].protection = Protection(locked=False)
    for address, value in {'D8': 'DATA DA SOLIC.', 'A9': 'ÁREA SOLICITANTE',
                           'A11': 'NOME DO MATERIAL', 'A15': 'Classificação',
                           'A17': 'Grupo de Itens', 'A19': 'Valor Material',
                           'A21': 'Item Importado?', 'A24': 'Requisitante',
                           'A25': 'Gestor', 'D25': 'Data:', 'B4': date(2024, 3, 4)}.items():
        ws[address] = value
    wb.save(source); wb.close()
    target = tmp_path / 'filled.xlsx'
    write_excel(source, target, complete_item(), 'Criação de Material Nacional', fingerprint(source))
    a, b = read_template(source), read_template(target)
    for sheet in a:
        for row in sheet:
            for cell in row:
                if cell.coordinate not in excel_writer.MAP.values():
                    actual = b[sheet.title][cell.coordinate]
                    assert cell.value == actual.value, cell.coordinate
                    assert cell._style == actual._style, cell.coordinate
    assert b.worksheets[0]['A33'].value == 'Controle institucional'
    a.close(); b.close()


def test_scale_guard_rejects_unreadable_fit():
    ws = SimpleNamespace(PageSetup=SimpleNamespace(LeftMargin=36, RightMargin=36, TopMargin=36, BottomMargin=36),
                         Range=lambda area: SimpleNamespace(Width=500, Height=700))
    assert check_print_scale(ws, 'A1:E33') == 1
    ws.Range = lambda area: SimpleNamespace(Width=500, Height=1200)
    with pytest.raises(ValueError, match='85%'): check_print_scale(ws, 'A1:E33')


@pytest.mark.parametrize('protected', [False, True])
def test_pdf_native_flow_uses_disposable_copy_and_fits_one_page(extended_template, tmp_path, monkeypatch, protected):
    workbook = tmp_path / 'filled.xlsx'; item = complete_item(); item.request_date = date(2027, 8, 19)
    write_excel(extended_template, workbook, item, 'Criação de Material Nacional', fingerprint(extended_template))
    before = workbook.read_bytes(); pdf = tmp_path / 'output.pdf'; calls = []
    @contextmanager
    def session(path, read_only):
        assert Path(path) != workbook and not read_only
        assert Path(path).read_bytes() == before
        wb = read_template(path); sheet = wb.worksheets[0]
        setup = SimpleNamespace(LeftMargin=36, RightMargin=36, TopMargin=36, BottomMargin=36)
        class Sheet:
            PageSetup = setup
            ProtectContents = protected
            def Range(self, address):
                if ':' in address: return SimpleNamespace(Width=500, Height=700)
                value = sheet[address].value
                text = value.strftime('%d/%m/%Y') if isinstance(value, datetime) else str(value or '')
                return SimpleNamespace(Value=value, Text=text, NumberFormat=sheet[address].number_format)
            def ResetAllPageBreaks(self):
                calls.append('reset')
                if protected: raise RuntimeError('Quebras não podem ser excluídas nesta folha protegida')
            def ExportAsFixedFormat(self, *args):
                assert setup.FitToPagesWide == setup.FitToPagesTall == 1 and setup.Zoom is False
                assert setup.PaperSize == 9 and setup.Orientation == 1
                assert calls == ['reset']
                texts = [self.Range(c.coordinate).Text for row in sheet for c in row if c.value is not None]
                synthetic_pdf(pdf, texts)
        yield None, SimpleNamespace(Worksheets=lambda name: Sheet())
        wb.close()
    monkeypatch.setattr(pdf_exporter, 'excel_session', session)
    monkeypatch.setattr(pdf_exporter, 'sys', SimpleNamespace(platform='win32'))
    pdf_exporter.export_pdf(workbook, pdf, 'Criação de Material Nacional', item.name)
    assert pdf.is_file() and workbook.read_bytes() == before
    assert not list(tmp_path.glob('pcm_pdf_*'))
    with pymupdf.open(pdf) as doc: assert len(doc) == 1 and '19/08/2027' in doc[0].get_text()


def test_pdf_retry_uses_saved_excel_without_reprocessing(gui, monkeypatch):
    app, window, messages = gui
    window.items = [complete_item()]; window.refresh()
    def fail(*args): raise RuntimeError('Falha sintética de paginação')
    monkeypatch.setattr(main_window, 'export_pdf', fail)
    monkeypatch.setattr(main_window, 'sys', SimpleNamespace(platform='win32'))
    window.generate(True); wait_job(app, window)
    folder = window.outputs[0]; workbook = folder / 'cadastro.xlsx'; before = workbook.read_bytes()
    assert workbook.is_file() and not (folder / 'cadastro.pdf').exists()
    assert window.retry_button.isEnabled() and str(folder) in window.status.text()
    def forbidden(*args, **kwargs): raise AssertionError('Retry must not reprocess or fill Excel')
    monkeypatch.setattr(main_window, 'read_pdf', forbidden)
    monkeypatch.setattr(main_window, 'write_excel', forbidden)
    calls = []
    def succeed(source, target, sheet, name, pages):
        calls.append((source, sheet, name, pages)); synthetic_pdf(target, [name])
    monkeypatch.setattr(main_window, 'export_pdf', succeed)
    window.settings.sheet = 'configuração posterior'
    window.retry_button.click(); wait_job(app, window)
    assert calls == [(workbook, 'Criação de Material Nacional', 'Parafuso M8', 1)]
    assert workbook.read_bytes() == before and (folder / 'cadastro.pdf').is_file()
    assert window.preview_button.isEnabled() and not window.retry_button.isEnabled()
    assert 'PDF exportado' in window.status.text()
