from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile
import pytest
import pymupdf
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation
from cadastro_pcm.core.models import Item, decimal_br, money_br
from cadastro_pcm.core.pdf_reader import read_pdf, Page
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.excel_writer import MAP, write_excel, fingerprint, inspect_template
from cadastro_pcm.core.file_manager import safe_name, new_output
from cadastro_pcm.core.pdf_exporter import export_pdf
from cadastro_pcm.core.research import apply_suggestion


@pytest.fixture
def template(tmp_path):
    path = tmp_path / 'oficial_sintetico.xlsx'
    wb = Workbook(); ws = wb.active; ws.title = 'Criação de Material Nacional'
    ws.print_area = 'A1:F30'; ws.merge_cells('A14:F14')
    ws['A1'] = 'INSTITUCIONAL'; ws['A26'] = 'Assinatura'; ws['A27'] = None
    ws['F30'] = '=1+1'; ws.protection.sheet = True
    for address in MAP.values(): ws[address].protection = Protection(locked=False)
    ws['A20'].number_format = '"R$" #,##0.00'
    dv = DataValidation(type='list', formula1='"Material,Serviço"'); ws.add_data_validation(dv); dv.add(ws['A16'])
    wb.create_sheet('Instruções Gerais')['A1'] = 'Não modificar'
    wb.save(path); return path


def approved():
    return Item(name='Gaxeta 16', description='Gaxeta\n16 mm', requester='Usuário real', unit_price='239', quantity=9, ncm='4016.93.00', approved=True)


def test_pdf_labels_multiple_and_unit_price(tmp_path):
    path = tmp_path / 'cotacao_sintetica.pdf'
    doc = pymupdf.open(); page = doc.new_page()
    page.insert_text((40,40), 'Material: Gaxeta 16\nQuantidade: 9\nPreço unitário: R$ 239,00\nPreço total: R$ 2.151,00\nNCM: 4016.93.00\nMaterial: Parafuso\nQuantidade: 2\nPreço unitário: R$ 5,00')
    doc.save(path); doc.close()
    pages = read_pdf(path); items = extract_items(pages)
    assert len(items) == 2
    assert items[0].unit_price == Decimal('239')
    assert items[0].total_price == Decimal('2151')
    assert items[0].quantity == 9 and items[0].ncm == '40169300'
    assert items[1].name == 'Parafuso'
    assert 'página 1' in items[0].evidence['unit_price'].source


def test_table_strategy():
    page = Page('teste', 1, 'Tabela', [[['Descrição', 'Qtd', 'Valor unitário', 'Total'], ['Gaxeta', '9', '239,00', '2.151,00'], ['Anel', '2', '3,00', '6,00']]])
    items = extract_items([page]); assert len(items) == 2
    assert items[0].unit_price == 239 and items[0].total_price == 2151


def test_total_never_used_as_unit():
    items = extract_items([Page('teste', 1, 'Material: X\nPreço total: 2.151,00', [])])
    assert items[0].unit_price is None


def test_scanned_pdf_requires_ocr(tmp_path):
    path = tmp_path / 'vazio.pdf'; doc = pymupdf.open(); doc.new_page(); doc.save(path); doc.close()
    with pytest.raises(ValueError, match='OCR'): read_pdf(path)


@pytest.mark.parametrize('value', ['123', 'abcdefgh', '4016-93-00'])
def test_invalid_ncm(value):
    with pytest.raises(ValueError): Item(ncm=value)


def test_currency():
    assert decimal_br('R$ 2.151,00') == Decimal('2151.00')
    assert money_br(Decimal('239')) == 'R$ 239,00'


def test_excel_copy_integrity(template, tmp_path):
    before = template.read_bytes(); target = tmp_path / 'copia.xlsx'
    write_excel(template, target, approved(), 'Criação de Material Nacional', fingerprint(template))
    assert template.read_bytes() == before
    wb = load_workbook(target); ws = wb.worksheets[0]
    assert ws['A12'].value == 'Gaxeta 16' and ws['A20'].value == 239
    assert ws['C16'].value is None and ws['A27'].value is None
    assert ws['F30'].value == '=1+1' and ws.protection.sheet
    assert '40169300' in ws['A14'].value and ws['E22'].value == 'SIM'
    assert wb.worksheets[1]['A1'].value == 'Não modificar'
    assert list(ws.merged_cells.ranges) and ws.print_area
    wb.close()
    with ZipFile(template) as original, ZipFile(target) as copy:
        assert original.namelist() == copy.namelist()
        for name in original.namelist():
            if name != 'xl/worksheets/sheet1.xml': assert original.read(name) == copy.read(name)


def test_locked_cells_refused(template, tmp_path):
    wb = load_workbook(template); wb.worksheets[0]['A12'].protection = Protection(locked=True); wb.save(template)
    with pytest.raises(ValueError, match='bloqueado'):
        write_excel(template, tmp_path/'x.xlsx', approved(), wb.sheetnames[0], fingerprint(template))
    assert not (tmp_path/'x.xlsx').exists()


def test_invalid_validation_refused(template, tmp_path):
    item = approved(); item.classification = 'Inexistente'
    with pytest.raises(ValueError, match='não permitido'):
        write_excel(template, tmp_path/'x.xlsx', item, 'Criação de Material Nacional', fingerprint(template))


def test_authorization_and_approval(template, tmp_path):
    with pytest.raises(ValueError, match='autorize'):
        write_excel(template, tmp_path/'x.xlsx', approved(), 'Criação de Material Nacional', '')
    item = approved(); item.approved = False
    with pytest.raises(ValueError, match='aprove'): item.require_ready()
    report = inspect_template(template)
    assert report['sha256'] == fingerprint(template)
    assert len(report['sheets']) == 2


def test_files(tmp_path):
    assert safe_name('CON') == '_CON'
    assert '/' not in safe_name('a/b:<>')
    assert new_output(tmp_path, 'Gaxeta') != new_output(tmp_path, 'Gaxeta')


def test_research_resets_approval():
    item = approved()
    updated = apply_suggestion(item, 'unit_price', '240', 'https://fabricante.example/produto', True)
    assert item.unit_price == 239
    assert updated.unit_price == 240 and not updated.approved
    assert updated.evidence['unit_price'].estimated


def test_no_fake_export_on_linux(tmp_path):
    import sys
    if sys.platform == 'win32': pytest.skip('Teste específico Linux')
    with pytest.raises(RuntimeError, match='Windows'):
        export_pdf('ausente.xlsx', tmp_path/'x.pdf', 'aba', 'nome')
    assert not (tmp_path/'x.pdf').exists()


def test_synthetic_pdf_to_excel(template, tmp_path):
    pdf = tmp_path / 'entrada.pdf'
    doc = pymupdf.open(); page = doc.new_page()
    page.insert_text((40, 40), 'Material: Gaxeta 16\nQuantidade: 9\nValor unitário: 239,00\nValor total: 2.151,00\nNCM: 4016.93.00')
    doc.save(pdf); doc.close()
    item = extract_items(read_pdf(pdf))[0]
    item.requester = 'Revisor do teste'; item.approved = True
    target = new_output(tmp_path / 'saida', item.name) / 'cadastro.xlsx'
    write_excel(template, target, item, 'Criação de Material Nacional', fingerprint(template))
    wb = load_workbook(target)
    assert wb.worksheets[0]['A20'].value == 239
    assert wb.worksheets[0]['A12'].value == 'Gaxeta 16'
    wb.close()


def test_gui_review(monkeypatch):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication
    from cadastro_pcm.core.settings import Settings
    from cadastro_pcm.interface import main_window
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor'))
    app = QApplication.instance() or QApplication([])
    window = main_window.Window()
    item = approved(); window.items = [item]; window.refresh()
    assert window.table.item(0, 2).text() == 'R$ 239,00'
    editor = main_window.Editor(item, window)
    editor.inputs['name'].setText('Nome corrigido')
    editor.confirm.setChecked(True); editor.save(item)
    assert editor.result_item.name == 'Nome corrigido'
    assert editor.result_item.evidence['name'].origin == 'manual'
    editor.deleteLater(); window.deleteLater(); app.processEvents()


def test_interface_selection_and_actions(monkeypatch, tmp_path):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication
    from cadastro_pcm.core.settings import Settings
    from cadastro_pcm.interface import main_window
    app = QApplication.instance() or QApplication([])
    model = tmp_path / 'modelo.xlsx'; model.touch()
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor', template=str(model), authorized_template_sha256='teste'))
    window = main_window.Window()
    assert not window.edit_button.isEnabled()
    assert not window.process_button.isEnabled()
    pdf = tmp_path / 'cotação.pdf'; pdf.write_bytes(b'%PDF')
    window.add_paths([str(pdf), str(pdf)])
    assert window.documents.count() == 1
    assert window.document_path(0) == str(pdf)
    assert 'cotação.pdf' in window.documents.item(0).text()
    assert window.process_button.isEnabled()
    item = approved(); window.items = [item]; window.refresh()
    assert window.excel_button.isEnabled()
    assert 'Gaxeta 16' in window.details.toPlainText()
    assert window.detail_badge.text() == 'Pronto para gerar'
    assert not window.preview_button.isEnabled()
    folder = tmp_path / 'saida'; folder.mkdir(); (folder / 'cadastro.pdf').touch()
    window.outputs[0] = folder; window.update_actions()
    assert window.preview_button.isEnabled()
    item.unit_price = None; window.refresh()
    assert not window.pdf_button.isEnabled()
    assert window.detail_badge.text() == 'Incompleto'
    window.documents.setCurrentRow(0); window.remove_file()
    assert not window.process_button.isEnabled()
    window.deleteLater(); app.processEvents()


def test_settings_dialog_structure(monkeypatch):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication, QDialog, QTabWidget, QComboBox
    from cadastro_pcm.core.settings import Settings
    from cadastro_pcm.interface import main_window
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor'))
    window = main_window.Window()
    captured = []
    def inspect_dialog(dialog):
        captured.append(dialog.findChild(QTabWidget).count())
        assert dialog.findChild(QComboBox).currentText() == 'atual'
        return 0
    monkeypatch.setattr(QDialog, 'exec', inspect_dialog)
    window.configure()
    assert captured == [3]
    window.deleteLater(); app.processEvents()
