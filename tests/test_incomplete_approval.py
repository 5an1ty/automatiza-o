"""Approval with explicit per-field authorization and bounded long-list dialogs."""
import json
from pathlib import Path

import pytest
from openpyxl import load_workbook
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QPushButton

from cadastro_pcm.core import generation, excel_writer
from cadastro_pcm.core.models import Item
from cadastro_pcm.interface import main_window
from cadastro_pcm.interface.approval_dialog import ApprovalDialog
from cadastro_pcm.interface.components import item_status
from test_workflow import approved, template
from test_optional_price_approval import gui


@pytest.mark.parametrize('fields', [{}, {'unit_price': None}, {'ncm': ''}, {'unit_price': None, 'ncm': ''}])
def test_explicit_authorization_and_record(fields):
    item = approved().model_copy(update={**fields, 'approved': False})
    if fields:
        with pytest.raises(ValueError): item.approve()
    item.approve(True)
    item.require_ready()
    assert set(item.authorized_pending_fields) == set(fields)
    assert item.pending_fields() if fields else not item.pending_fields()
    assert item_status(item)[0] == ('Aprovado com pendências' if fields else 'Pronto para gerar')
    restored = Item.model_validate_json(item.model_dump_json())
    assert restored.authorized_pending_fields == item.authorized_pending_fields
    restored.require_ready()


def test_authorization_is_not_permission_for_other_missing_fields():
    item = approved().model_copy(update={'unit_price': None, 'approved': False})
    item.approve(True)
    item.ncm = ''
    with pytest.raises(ValueError, match='NCM'): item.require_ready()
    item.name = ''
    with pytest.raises(ValueError, match='Nome'): item.approve(True)
    legacy = approved().model_dump(); legacy.pop('authorized_pending_fields')
    assert Item.model_validate(legacy).authorized_pending_fields == []


@pytest.mark.parametrize('choice', [QDialog.Rejected, ApprovalDialog.COMPLETE_ONLY, ApprovalDialog.ALLOW_INCOMPLETE])
def test_gui_three_choices_preserve_data(gui, monkeypatch, choice):
    _, window = gui
    items = [approved().model_copy(update={'approved': False}),
             approved().model_copy(update={'approved': False, 'unit_price': None, 'ncm': ''}),
             approved().model_copy(update={'approved': False, 'name': ''})]
    window.items = items; window.refresh(); window.toggle_all_items()
    before = [item.model_dump_json() for item in items]
    monkeypatch.setattr(ApprovalDialog, 'exec', lambda dialog: choice)
    window.approve_selected()
    expected = [False, False, False] if choice == 0 else ([True, False, False] if choice == 1 else [True, True, False])
    assert [item.approved for item in items] == expected
    if choice == 0:
        assert [item.model_dump_json() for item in items] == before
    if choice == 2:
        assert items[1].authorized_pending_fields == ['unit_price', 'ncm']
        assert items[1].ncm == '' and items[1].unit_price is None
        assert items[2].authorized_pending_fields == []
        assert window.batch_button.text().endswith('(2)')


def test_individual_incomplete_then_revision_revokes(gui, monkeypatch):
    _, window = gui
    item = approved().model_copy(update={'approved': False, 'unit_price': None, 'ncm': ''})
    monkeypatch.setattr(ApprovalDialog, 'exec', lambda dialog: ApprovalDialog.ALLOW_INCOMPLETE)
    editor = main_window.Editor(item, window)
    editor.confirm.setChecked(True); editor.save(item)
    assert editor.result() == QDialog.Accepted
    result = editor.result_item
    assert result.approved and result.authorized_pending_fields == ['unit_price', 'ncm']
    assert not result.ncm_not_applicable and not result.price_fill_later
    revised = main_window.Editor(result, window)
    revised.inputs['name'].setText('Nome revisado'); revised.save(result)
    assert not revised.result_item.approved and not revised.result_item.authorized_pending_fields


def test_approve_all_acts_on_visible_and_preserves_selection(gui, monkeypatch):
    _, window = gui
    window.items = [approved().model_copy(update={'approved': False, 'ncm': ''}) for _ in range(4)]
    window.refresh(); window.table.setRowHidden(3, True)
    selection = window.selected_rows()
    monkeypatch.setattr(ApprovalDialog, 'exec', lambda dialog: ApprovalDialog.ALLOW_INCOMPLETE)
    window.approve_all_button.click()
    assert [i.approved for i in window.items] == [True, True, True, False]
    assert window.selected_rows() == selection


@pytest.mark.parametrize('count', [1, 10, 39, 100])
@pytest.mark.parametrize('result', [False, True])
def test_dialog_bounded_scroll_and_footer(gui, count, result):
    app, window = gui
    items = [approved().model_copy(update={'name': f'Material {n} com uma descrição longa', 'unit_price': None, 'ncm': ''}) for n in range(count)]
    dialog = ApprovalDialog(items, window, result=result)
    dialog.show(); app.processEvents()
    area = dialog.screen().availableGeometry()
    assert dialog.width() <= area.width() and dialog.height() <= area.height()
    assert dialog.details.count() == count
    assert dialog.details.horizontalScrollBarPolicy() == Qt.ScrollBarAlwaysOff
    for button in dialog.buttons.findChildren(QPushButton):
        assert button.isVisible()
        assert dialog.rect().contains(button.mapTo(dialog, button.rect().center()))
        assert button.focusPolicy() != Qt.NoFocus
    if count >= 39:
        bar = dialog.details.verticalScrollBar()
        assert bar.maximum() > 0
        bar.setValue(bar.maximum()); app.processEvents()
        assert bar.value() == bar.maximum()
        assert dialog.details.visualItemRect(dialog.details.item(count - 1)).intersects(dialog.details.viewport().rect())
    dialog.close()


def test_generate_incomplete_batch_preserves_corporate_fields(gui, monkeypatch):
    _, window = gui
    before = Path(window.settings.template).read_bytes()
    items = [approved().model_copy(update={'unit_price': None, 'ncm': '', 'approved': False}) for _ in range(3)]
    for item in items: item.approve(True)
    def export(xlsx, target, *args):
        book = load_workbook(xlsx); sheet = book.active
        assert sheet['A20'].value is None
        assert sheet['A14'].value == items[0].description
        assert 'NCM:' not in sheet['A14'].value
        assert sheet['E22'].value == 'SIM'
        assert sheet['C16'].value is None and sheet['A27'].value is None
        assert sheet.protection.sheet and sheet['A1'].value == 'INSTITUCIONAL'
        book.close()
        Path(target).write_bytes(b'%PDF synthetic export stub')
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch(list(enumerate(items)), window.settings)
    assert all(result.status == 'sucesso' for result in results)
    for result in results:
        record = json.loads((result.folder / 'registro.json').read_text())
        assert record['authorized_pending_fields'] == ['unit_price', 'ncm']
        assert not record['ncm_not_applicable'] and record['ncm'] == '' and record['unit_price'] is None
    assert Path(window.settings.template).read_bytes() == before


@pytest.mark.parametrize('choice', [QDialog.Rejected, ApprovalDialog.COMPLETE_ONLY])
def test_individual_decline_or_cancel_does_not_approve(gui, monkeypatch, choice):
    _, window = gui
    item = approved().model_copy(update={'approved': False, 'ncm': ''})
    before = item.model_dump_json()
    monkeypatch.setattr(ApprovalDialog, 'exec', lambda dialog: choice)
    editor = main_window.Editor(item, window)
    editor.confirm.setChecked(True); editor.save(item)
    assert item.model_dump_json() == before
    assert not editor.result_item.approved and not editor.result_item.authorized_pending_fields
    assert editor.result() == (QDialog.Accepted if choice == ApprovalDialog.COMPLETE_ONLY else QDialog.Rejected)


def test_small_monitor_keeps_footer_in_useful_area(gui, monkeypatch):
    from types import SimpleNamespace
    from PySide6.QtCore import QRect
    app, window = gui
    screen = SimpleNamespace(availableGeometry=lambda: QRect(0, 0, 640, 480))
    monkeypatch.setattr(window, 'screen', lambda: screen)
    dialog = ApprovalDialog([approved().model_copy(update={'ncm': ''}) for _ in range(100)], window)
    monkeypatch.setattr(dialog, 'screen', lambda: screen)
    dialog.show(); app.processEvents()
    assert dialog.width() <= 640 and dialog.height() <= 480
    assert dialog.buttons.orientation() == Qt.Vertical
    for button in dialog.buttons.findChildren(QPushButton):
        assert dialog.rect().contains(button.mapTo(dialog, button.rect().topLeft()))
        assert dialog.rect().contains(button.mapTo(dialog, button.rect().bottomRight()))
    assert dialog.details.viewport().height() > 0
    dialog.close()
