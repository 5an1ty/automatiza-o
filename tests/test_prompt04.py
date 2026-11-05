"""Synthetic regression cases only; official documents stay outside Git."""
from decimal import Decimal
from pathlib import Path
from time import monotonic
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pymupdf
import pytest
from openpyxl import Workbook
from openpyxl.styles import Protection
from PySide6.QtCore import QEventLoop
from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QTextEdit

from cadastro_pcm.core.excel_writer import MAP, check_template, fingerprint, inspect_template, write_excel
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.models import Item
from cadastro_pcm.core.pdf_reader import Page, read_pdf
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core.template_reader import read_template
from cadastro_pcm.interface import main_window


@pytest.mark.parametrize('text', [
    'NCM 4016.93.00- Gaxeta 16', 'Gaxeta 16 - NCM 4016.93.00',
    'Gaxeta 16 NCM: 4016.93.00', '40169300 - Gaxeta 16',
    'Gaxeta 16 (NCM 40169300)', '4016.93.00 - Gaxeta 16',
])
def test_ncm_in_item_description(text):
    item = extract_items([Page('sintetico.pdf', 2, '', [[
        ['Descrição', 'R$ Unit.'], [text, '18,50']]])])[0]
    assert item.name == item.description == 'Gaxeta 16'
    assert item.ncm == '4016.93.00' and item.unit_price == Decimal('18.50')
    assert item.evidence['ncm'].origin == 'pdf'
    assert item.evidence['ncm'].source == 'sintetico.pdf — página 2'
    assert text in item.evidence['ncm'].excerpt


def test_ncm_not_taken_from_contacts_or_manufacturer_code():
    text = 'Telefone: 12345678\nData: 12/03/2026\nProduto: Anel REF 12345678\nCódigo do fabricante: 40169300\nPreço total: 10,00'
    item = extract_items([Page('sintetico', 1, text, [])])[0]
    assert item.name == 'Anel REF 12345678'
    assert not item.ncm and item.manufacturer_code == '40169300'
    assert item.unit_price is None
    assert any('Preço unitário' in w for w in item.warnings)


def test_conflicting_fiscal_codes_require_review():
    item = extract_items([Page('sintetico', 1, '', [[
        ['Descrição', 'NCM', 'R$ Unit.'],
        ['Anel NCM 4016.93.00', '7318.15.00', '12,00']]])])[0]
    assert item.name == 'Anel' and not item.ncm
    assert any('divergente' in w for w in item.warnings)


def quotation(path):
    # A ruled table with a preamble, empty supplier-code cell and separate tax column.
    with pymupdf.open() as doc:
        page = doc.new_page()
        page.insert_text((30, 35), 'Contato: 12345678 - Cotacao sintetica')
        xs = [30, 75, 150, 330, 410, 475, 580]
        rows = [
            ['Qtd.', 'Cod. fornecedor', 'Descrição', 'R$ Unit.', 'IPI', 'R$ Total'],
            ['3', '', 'NCM 7318.15.00- Parafuso M8', 'R$ 18,50', 'R$ 0,00', 'R$ 55,50'],
            ['2', '', 'Anel', '', 'R$ 0,00', 'R$ 8,00'],
        ]
        for y in (80, 110, 140, 170): page.draw_line((30, y), (580, y))
        for x in xs: page.draw_line((x, 80), (x, 170))
        for row, cells in enumerate(rows):
            for x, cell in zip(xs, cells): page.insert_text((x + 3, 100 + row * 30), cell, fontsize=8)
        doc.save(path)
    return path


def test_positional_pdf_prices_empty_cells_and_total(tmp_path):
    pdf = quotation(tmp_path / 'sintetico.pdf')
    before = pdf.read_bytes()
    items = extract_items(read_pdf(pdf))
    assert len(items) == 2
    first, second = items
    assert first.name == 'Parafuso M8' and first.ncm == '7318.15.00'
    assert first.quantity == 3 and first.unit_price == Decimal('18.50')
    assert first.total_price == Decimal('55.50')
    assert second.unit_price is None and second.total_price == 8 and not second.ncm
    assert pdf.read_bytes() == before


def test_discrepant_total_is_preserved():
    item = extract_items([Page('sintetico', 1, '', [[
        ['Descrição', 'Qtd.', 'R$ Unit.', 'R$ Total'], ['Anel', '2', '10,00', '21,00']]])])[0]
    assert item.unit_price == 10 and item.total_price == 21
    assert any('impostos' in w for w in item.warnings)


@pytest.fixture
def extended_template(tmp_path):
    path = tmp_path / 'modelo_sintetico.xlsx'
    wb = Workbook(); ws = wb.active; ws.title = 'Criação de Material Nacional'
    ws.protection.sheet = True; ws['A33'] = 'Controle sintético'
    ws.merge_cells('A14:E14'); ws.merge_cells('A12:E12')
    for address in MAP.values(): ws[address].protection = Protection(locked=False)
    wb.create_sheet('Listas')['A1'] = 'Sim'
    wb['Listas']['A2'] = 'Não'
    wb.save(path)
    # Use Excel x14 validation XML, including namespace prefixes in mc:Ignorable.
    with ZipFile(path) as archive:
        parts = {n: archive.read(n) for n in archive.namelist()}
    xml = parts['xl/worksheets/sheet1.xml'].decode()
    xml = xml.replace('<worksheet ', '<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" mc:Ignorable="x14" ', 1)
    xml = xml.replace('</worksheet>', '''<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">
    <x14:dataValidations count="2" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <x14:dataValidation type="list"><x14:formula1><xm:f>'Listas'!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>E18 E22</xm:sqref></x14:dataValidation>
    <x14:dataValidation type="list"><x14:formula1><xm:f>'Listas'!$A$1:$A$2</xm:f></x14:formula1><xm:sqref>A22</xm:sqref></x14:dataValidation>
    </x14:dataValidations></ext></extLst></worksheet>''')
    parts['xl/worksheets/sheet1.xml'] = xml.encode()
    with ZipFile(path, 'w') as archive:
        for name, content in parts.items(): archive.writestr(name, content)
    return path


def complete_item():
    return Item(name='Parafuso M8', description='Parafuso de teste', unit_price='18.50',
                quantity=3, ncm='73181500', requester='Revisor sintético', approved=True)


def test_extended_lists_and_package_preservation(extended_template, tmp_path):
    template = extended_template; original = template.read_bytes()
    target = tmp_path / 'copia.xlsx'
    report = inspect_template(template)
    assert len(report['sheets']['Criação de Material Nacional']['validations']) == 2
    values, _ = check_template(template, complete_item(), 'Criação de Material Nacional', fingerprint(template))
    assert 'E18' not in values and values['E22'] == 'SIM'
    write_excel(template, target, complete_item(), 'Criação de Material Nacional', fingerprint(template))
    assert template.read_bytes() == original
    with ZipFile(template) as before, ZipFile(target) as after:
        assert before.namelist() == after.namelist()
        for name in before.namelist():
            if name != 'xl/worksheets/sheet1.xml': assert before.read(name) == after.read(name)
        a = before.read('xl/worksheets/sheet1.xml').decode()
        b = after.read('xl/worksheets/sheet1.xml').decode()
        assert a.split('<sheetData>')[0] == b.split('<sheetData>')[0]
        assert a.split('</sheetData>')[1] == b.split('</sheetData>')[1]
        ET.fromstring(b)  # All prefixes must remain bound.
    wb = read_template(target); ws = wb.worksheets[0]
    assert len(ws.data_validations.dataValidation) == 2
    assert ws['A20'].value == 18.5 and ws['A12'].value == 'Parafuso M8'
    assert ws['E18'].value is None and ws['C16'].value is None
    assert ws.protection.sheet and ws['A24'].value is None
    assert '7318.15.00' in ws['A14'].value
    wb.close()


def test_extended_validation_blocks_invalid_values(extended_template, tmp_path):
    item = complete_item(); item.lower_ncm = 'Talvez'
    with pytest.raises(ValueError, match='E22.*não permitido'):
        write_excel(extended_template, tmp_path / 'falha.xlsx', item, 'Criação de Material Nacional', fingerprint(extended_template))
    assert not (tmp_path / 'falha.xlsx').exists()


@pytest.fixture
def gui(monkeypatch, tmp_path, extended_template):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    app = QApplication.instance() or QApplication([])
    config = Settings(requester='Revisor sintético', template=str(extended_template),
                      authorized_template_sha256=fingerprint(extended_template), output=str(tmp_path / 'saida'))
    monkeypatch.setattr(main_window, 'load_settings', lambda: config)
    messages = []
    for method in ('information', 'warning'):
        monkeypatch.setattr(QMessageBox, method, lambda parent, title, text: messages.append((title, text)))
    window = main_window.Window()
    yield app, window, messages
    assert window.worker is None
    window.close(); window.deleteLater(); app.processEvents()


def wait_job(app, window):
    deadline = monotonic() + 10
    while window.worker is not None and monotonic() < deadline:
        app.processEvents(QEventLoop.AllEvents, 20)
    assert window.worker is None, 'Worker did not finish'


def test_gui_selection_source_and_removal(gui, tmp_path, monkeypatch):
    app, window, messages = gui
    pdf = quotation(tmp_path / 'cotacao.pdf')
    assert not window.remove_button.isEnabled() and not window.source_button.isEnabled()
    window.add_paths([str(pdf), str(pdf)])
    assert window.documents.count() == 1 and window.documents.selectedItems()
    assert window.remove_button.isEnabled() and window.source_button.isEnabled()
    window.documents.clearSelection()
    assert not window.remove_button.isEnabled() and not window.source_button.isEnabled()
    window.documents.setCurrentRow(0)
    captured = []
    def inspect_dialog(dialog):
        captured.append(dialog.findChild(QTextEdit).toPlainText())
        return 0
    monkeypatch.setattr(QDialog, 'exec', inspect_dialog)
    window.source_button.click(); wait_job(app, window)
    assert 'Parafuso M8' in captured[0] and 'Página 1' in captured[0]
    item = complete_item(); window.items = [item]; window.refresh()
    window.remove_button.click()
    assert window.documents.count() == 0 and window.items == [item]
    assert not window.process_button.isEnabled()


def test_gui_process_reprocess_preserves_edits_until_confirmed(gui, tmp_path, monkeypatch):
    app, window, messages = gui
    window.add_paths([str(quotation(tmp_path / 'cotacao.pdf'))])
    window.process_button.click()
    assert window.progress.isVisibleTo(window.centralWidget())
    wait_job(app, window)
    assert len(window.items) == 2 and window.table.currentRow() == 0
    assert window.items[0].unit_price == Decimal('18.50')
    assert 'Processamento concluído' in window.status.text()
    assert any(title == 'Processamento concluído' for title, _ in messages)
    assert window.table.columnCount() == 6
    assert all(window.table.horizontalHeaderItem(c).text() != 'Quantidade' for c in range(6))
    window.items[0].name = 'Edição preservada'; window.items[0].approved = True
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.No)
    window.process_button.click()
    assert window.items[0].name == 'Edição preservada' and window.items[0].approved
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.Yes)
    window.process_button.click(); wait_job(app, window)
    assert len(window.items) == 2 and window.items[0].name == 'Parafuso M8'
    assert not window.items[0].approved


def test_gui_review_approval_and_excel_generation(gui):
    app, window, messages = gui
    item = complete_item(); item.approved = False
    window.items = [item]; window.refresh()
    assert not window.excel_button.isEnabled() and 'aprovação' in window.generation_hint.text()
    editor = main_window.Editor(item, window)
    assert 'quantity' not in editor.inputs
    editor.inputs['description'].setPlainText('Descrição técnica revisada')
    editor.inputs['manufacturer'].setText('Fabricante sintético')
    editor.confirm.setChecked(True); editor.save(item)
    assert editor.result_item.approved and editor.result_item.quantity == 3
    assert editor.result_item.name == 'Parafuso M8'
    assert editor.result_item.evidence['description'].origin == 'manual'
    window.items = [editor.result_item]; window.refresh()
    assert window.excel_button.isEnabled()
    window.excel_button.click(); wait_job(app, window)
    folder = window.outputs[0]
    assert (folder / 'cadastro.xlsx').is_file() and (folder / 'registro.json').is_file()
    assert any(title == 'Arquivos gerados' for title, _ in messages)
    assert str(folder) in window.status.text()
    wb = read_template(folder / 'cadastro.xlsx')
    assert wb.worksheets[0]['A20'].value == 18.5
    assert 'Descrição técnica revisada' in wb.worksheets[0]['A14'].value
    wb.close(); editor.deleteLater()


def test_gui_incomplete_approval_and_model_blockers(gui):
    app, window, messages = gui
    item = complete_item(); item.unit_price = None; item.ncm = ''; item.approved = False
    editor = main_window.Editor(item, window); editor.confirm.setChecked(True); editor.save(item)
    assert editor.result() != QDialog.Accepted
    assert 'Preço unitário' in messages[-1][1] and 'NCM' in messages[-1][1]
    window.items = [item]; window.refresh()
    assert not window.excel_button.isEnabled()
    assert 'Preço unitário' in window.generation_hint.text() and 'NCM' in window.generation_hint.text()
    window.items = [complete_item()]; window.refresh()
    assert window.excel_button.isEnabled()
    window.settings.authorized_template_sha256 = 'hash antigo'; window.update_actions()
    assert not window.excel_button.isEnabled() and 'alterado' in window.generation_hint.text()
    window.settings.template = ''; window.update_actions()
    assert 'Modelo não configurado' in window.generation_hint.text()
    editor.deleteLater()


def test_gui_export_failure_reports_saved_excel(gui, monkeypatch):
    app, window, messages = gui
    window.items = [complete_item()]; window.refresh()
    def fail(*args): raise RuntimeError('Falha sintética do Excel PDF')
    monkeypatch.setattr(main_window, 'export_pdf', fail)
    window.generate(True); wait_job(app, window)
    assert 'O PDF não foi gerado' in window.status.text()
    assert 'Falha sintética' in window.status.text()
    assert list(Path(window.settings.output).rglob('cadastro.xlsx'))
    assert not list(Path(window.settings.output).rglob('cadastro.pdf'))
    assert not window.preview_button.isEnabled()


def test_gui_processing_error_and_no_items(gui, tmp_path):
    app, window, messages = gui
    invalid = tmp_path / 'invalido.pdf'; invalid.write_bytes(b'invalido')
    window.add_paths([str(invalid)]); window.process_button.click(); wait_job(app, window)
    assert not window.items and any('invalido.pdf' in text for _, text in messages)
    window.remove_file()
    path = tmp_path / 'sem_itens.pdf'
    with pymupdf.open() as doc:
        page = doc.new_page(); page.insert_text((30, 30), 'Documento sintetico sem produtos')
        doc.save(path)
    window.add_paths([str(path)]); window.process_button.click(); wait_job(app, window)
    assert not window.items and any('layout não reconhecido' in text for _, text in messages)
    assert '0 item(ns)' in window.status.text()
