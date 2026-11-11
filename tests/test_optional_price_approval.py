"""Explicit deferred prices and human bulk approval; native Excel is simulated."""
from contextlib import contextmanager
from copy import copy
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
import json

import pytest
import pymupdf
from openpyxl import load_workbook
from openpyxl.utils.datetime import to_excel
from PySide6.QtCore import QItemSelectionModel
from PySide6.QtWidgets import QApplication, QMessageBox

from cadastro_pcm.core.models import Item
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core import generation, excel_writer, pdf_exporter
from cadastro_pcm.interface import main_window
from cadastro_pcm.interface.components import item_status
from test_workflow import template, approved


@pytest.fixture
def gui(monkeypatch, tmp_path, template):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    settings = Settings(requester='Teste', template=str(template),
                        authorized_template_sha256=excel_writer.fingerprint(template),
                        output=str(tmp_path / 'saida'))
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    monkeypatch.setattr(QMessageBox, 'information', lambda *args: None)
    monkeypatch.setattr(QMessageBox, 'warning', lambda *args: None)
    app = QApplication.instance() or QApplication([])
    window = main_window.Window()
    yield app, window
    window.close()


@pytest.mark.parametrize('price,later,ready', [('10', False, True), (None, False, False), (None, True, True), ('0', False, True)])
def test_price_validation_and_legacy(price, later, ready):
    item = Item.model_validate({**approved().model_dump(), 'unit_price': price, 'price_fill_later': later})
    assert not item.pending_fields() if ready else item.pending_fields()
    if ready:
        item.require_ready()
    else:
        with pytest.raises(ValueError, match='Preço'): item.require_ready()
    legacy = approved().model_dump(); legacy.pop('price_fill_later')
    assert Item.model_validate(legacy).price_fill_later is False
    assert Item.model_validate_json(item.model_dump_json()).price_fill_later == later


@pytest.mark.parametrize('existing', [None, 123.45])
def test_blank_excel_clears_old_price_preserves_style_and_protection(template, tmp_path, existing):
    wb = load_workbook(template); wb.active['A20'] = existing; wb.save(template); wb.close()
    original = template.read_bytes()
    item = approved().model_copy(update={'unit_price': None, 'price_fill_later': True})
    target = tmp_path / 'blank.xlsx'
    excel_writer.write_excel(template, target, item, 'Criação de Material Nacional', excel_writer.fingerprint(template))
    wb = load_workbook(target); source = load_workbook(template)
    ws = wb.active
    assert ws['A20'].value is None and ws['A20'].number_format == source.active['A20'].number_format
    assert copy(ws['A20'].protection) == copy(source.active['A20'].protection)
    assert ws['C16'].value is None and ws['A27'].value is None and ws.protection.sheet
    assert ws['A12'].value == item.name and ws['A1'].value == 'INSTITUCIONAL'
    wb.close(); source.close()
    assert template.read_bytes() == original


@pytest.mark.parametrize('formula', [True, False])
def test_cannot_clear_protected_or_formula_price(template, tmp_path, formula):
    from openpyxl.styles import Protection
    wb = load_workbook(template)
    wb.active['A20'] = '=1+1' if formula else 77
    if not formula:
        wb.active['A20'].protection = Protection(locked=True)
    wb.save(template); wb.close()
    item = approved().model_copy(update={'unit_price': None, 'price_fill_later': True})
    with pytest.raises(ValueError, match='fórmula' if formula else 'bloqueado'):
        excel_writer.write_excel(template, tmp_path / 'blocked.xlsx', item,
                                 'Criação de Material Nacional', excel_writer.fingerprint(template))


def test_native_com_writes_empty_not_zero(template, tmp_path, monkeypatch):
    wb = load_workbook(template); wb.active['A20'] = 77; wb.save(template); wb.close()
    calls = []
    @contextmanager
    def session(path, read_only):
        book = load_workbook(path); sheet = book.active
        class Cell:
            HasFormula = False
            Locked = False
            MergeCells = False
            def __init__(self, address): self.address = address
            @property
            def Value(self): return sheet[self.address].value
            @Value.setter
            def Value(self, value):
                calls.append((self.address, value)); sheet[self.address] = value
        ws = SimpleNamespace(Range=Cell, ProtectContents=True,
                             Cells=SimpleNamespace(SpecialCells=lambda *args: object()),
                             Application=SimpleNamespace(Intersect=lambda *args: None))
        yield None, SimpleNamespace(ReadOnly=False, Worksheets=lambda _: ws, Save=lambda: book.save(path))
        book.close()
    monkeypatch.setattr(excel_writer, 'excel_session', session)
    output = tmp_path / 'native.xlsx'
    excel_writer.write_excel_com(template, output, {'A20': None}, 'Criação de Material Nacional', None)
    assert calls == [('A20', None)]
    wb = load_workbook(output); assert wb.active['A20'].value is None; wb.close()


def test_editor_toggle_persistence_revoke_and_individual_approval(gui):
    _, window = gui
    item = approved().model_copy(update={'unit_price': None})
    editor = main_window.Editor(item, window)
    assert not editor.price_later.isChecked()
    editor.price_later.setChecked(True)
    assert not editor.inputs['unit_price'].isEnabled()
    assert not editor.price_warning.isHidden()
    editor.confirm.setChecked(True)
    editor.save(item)
    deferred = editor.result_item
    assert deferred.approved and deferred.price_fill_later and deferred.unit_price is None
    assert 'Preço a preencher' in item_status(deferred)[0]
    reopened = main_window.Editor(deferred, window)
    assert reopened.price_later.isChecked() and not reopened.confirm.isChecked()
    reopened.price_later.setChecked(False)
    assert reopened.inputs['unit_price'].isEnabled()
    reopened.inputs['unit_price'].setText('25,50')
    reopened.confirm.setChecked(True)
    reopened.inputs['name'].setText('Nome revisado')
    assert not reopened.confirm.isChecked()
    reopened.save(deferred)
    assert not reopened.result_item.approved and not reopened.result_item.price_fill_later
    assert str(reopened.result_item.unit_price) == '25.50'
    reopened.confirm.setChecked(True); reopened.save(deferred)
    assert reopened.result_item.approved


def test_select_all_none_partial_and_visible_only(gui):
    _, window = gui
    window.items = [approved() for _ in range(4)]; window.refresh(); window.table.clearSelection()
    original = [i.model_dump_json() for i in window.items]
    window.toggle_all_items(); assert window.selected_rows() == [0, 1, 2, 3]
    window.toggle_all_items(); assert window.selected_rows() == []
    window.table.selectRow(1); assert window.selected_rows() == [1]
    window.table.setRowHidden(3, True)
    window.toggle_all_items(); assert window.selected_rows() == [0, 1, 2]
    window.toggle_all_items(); assert window.selected_rows() == []
    assert [i.model_dump_json() for i in window.items] == original


def test_bulk_approval_confirmation_pending_and_deferred(gui, monkeypatch):
    _, window = gui
    items = [approved().model_copy(update={'approved': False}),
             approved().model_copy(update={'approved': False, 'unit_price': None, 'price_fill_later': True}),
             approved().model_copy(update={'approved': False, 'ncm': ''}),
             approved().model_copy(update={'approved': False, 'unit_price': None})]
    window.items = items; window.refresh(); window.toggle_all_items()
    assert window.selected_rows() == [0, 1, 2, 3]
    captured = []
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: captured.append(args[-1]) or QMessageBox.No)
    window.approve_selected(); assert not any(i.approved for i in items)
    assert '2 aptos' in captured[0] and '1 com preço a preencher' in captured[0] and 'NCM' in captured[0]
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.Yes)
    window.approve_selected()
    assert [i.approved for i in items] == [True, True, False, False]
    assert window.selected_rows() == [0, 1, 2, 3]
    assert window.batch_button.text().endswith('(2)')
    assert not items[3].price_fill_later


def test_bulk_approval_only_selected(gui, monkeypatch):
    _, window = gui
    window.items = [approved().model_copy(update={'approved': False}) for _ in range(3)]
    window.refresh(); window.table.selectRow(1)
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.Yes)
    window.approve_selected()
    assert [i.approved for i in window.items] == [False, True, False]


def test_batch_optional_price_and_record(gui, monkeypatch):
    _, window = gui
    item = approved().model_copy(update={'unit_price': None, 'price_fill_later': True})
    seen = []
    def export(xlsx, target, *args):
        wb = load_workbook(xlsx); seen.append(wb.active['A20'].value); wb.close()
        doc = pymupdf.open(); page = doc.new_page(); page.insert_text((40, 40), item.name)
        doc.save(target); doc.close()
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch([(0, item), (1, approved())], window.settings)
    assert all(r.status == 'sucesso' for r in results)
    assert seen == [None, 239]
    record = json.loads((results[0].folder / 'registro.json').read_text())
    assert record['price_fill_later'] is True and record['unit_price'] is None


def test_export_pdf_optional_price_native_simulation(template, tmp_path, monkeypatch):
    item = approved().model_copy(update={'unit_price': None, 'price_fill_later': True})
    workbook = tmp_path / 'sem_preco.xlsx'
    excel_writer.write_excel(template, workbook, item, 'Criação de Material Nacional', excel_writer.fingerprint(template))
    original = workbook.read_bytes()
    target = tmp_path / 'sem_preco.pdf'
    observed = []
    @contextmanager
    def session(path, read_only):
        book = load_workbook(path); sheet = book.active
        class Sheet:
            ProtectContents = True
            def Range(self, address):
                cell = sheet[address]; value = cell.value
                text = value.strftime('%d/%m/%Y') if isinstance(value, datetime) else str(value or '')
                return SimpleNamespace(Value=value, Value2=to_excel(value, book.epoch) if isinstance(value, datetime) else value,
                                       Text=text, NumberFormat=cell.number_format)
            def ResetAllPageBreaks(self): pass
            def ExportAsFixedFormat(self, *args):
                assert self.Range('A20').Value is None
                observed.append('A20 vazio antes de exportar')
                doc = pymupdf.open(); page = doc.new_page(width=595.28, height=841.89)
                y = 40
                for row in sheet:
                    for cell in row:
                        if cell.value is not None:
                            page.insert_text((40, y), self.Range(cell.coordinate).Text, fontsize=10)
                            y += 15
                doc.save(args[1]); doc.close()
        yield None, SimpleNamespace(Worksheets=lambda _: Sheet())
        book.close()
    monkeypatch.setattr(pdf_exporter, 'excel_session', session)
    monkeypatch.setattr(pdf_exporter, 'sys', SimpleNamespace(platform='win32'))
    monkeypatch.setattr(pdf_exporter, 'configure_page', lambda *args: None)
    monkeypatch.setattr(pdf_exporter, 'check_print_scale', lambda *args: None)
    pdf_exporter.export_pdf(workbook, target, 'Criação de Material Nacional', item.name)
    assert observed == ['A20 vazio antes de exportar'] and target.is_file()
    assert workbook.read_bytes() == original
    with pymupdf.open(target) as doc:
        text = doc[0].get_text()
        assert item.name in text and '239' not in text and '0,00' not in text
