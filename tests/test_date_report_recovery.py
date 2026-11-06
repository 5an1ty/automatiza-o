"""Recover sanitized E8 diagnostics even when automatic report files are absent."""
import json
from pathlib import Path

import pytest
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
from openpyxl import Workbook

from cadastro_pcm.core import excel_date
from cadastro_pcm.core.excel_date import DateValidationError, check_com_date, check_stored_date
from cadastro_pcm.interface import main_window
from test_e8_presentation import DateCell, EXPECTED, EPOCH
from test_prompt04 import gui, extended_template, complete_item, wait_job


def fail_writes(*args, **kwargs):
    raise PermissionError('Synthetic report storage failure')


def raise_date_error():
    check_com_date(DateCell('46331'), EXPECTED, EPOCH, phase='após reabrir')


def test_report_write_denied_uses_verified_temporary_file(tmp_path, monkeypatch):
    primary, fallback = tmp_path / 'primary', tmp_path / 'fallback'
    original = Path.open
    def open_file(path, *args, **kwargs):
        if path.parent == primary: raise PermissionError('Synthetic permission denial')
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'open', open_file)
    monkeypatch.setattr(excel_date.tempfile, 'mkdtemp', lambda **kwargs: str(fallback))
    with pytest.raises(DateValidationError) as error:
        check_com_date(DateCell('46331'), EXPECTED, EPOCH, phase='teste', report_dir=primary)
    path = error.value.report_path
    assert path.parent == fallback and path.is_file()
    assert json.loads(path.read_text()) == error.value.diagnostic


def test_file_removed_before_confirmation_is_not_announced_as_saved(tmp_path, monkeypatch):
    primary, fallback = tmp_path / 'primary', tmp_path / 'fallback'
    original = Path.read_text
    def read_file(path, *args, **kwargs):
        if path.parent == primary:
            path.unlink()
            raise FileNotFoundError('Synthetic disappeared report')
        return original(path, *args, **kwargs)
    monkeypatch.setattr(Path, 'read_text', read_file)
    monkeypatch.setattr(excel_date.tempfile, 'mkdtemp', lambda **kwargs: str(fallback))
    with pytest.raises(DateValidationError) as error:
        check_com_date(DateCell('46331'), EXPECTED, EPOCH, phase='teste', report_dir=primary)
    assert error.value.report_path.parent == fallback
    assert error.value.report_path.is_file() and not list(primary.iterdir())


def test_no_storage_still_retains_report_and_blocks_unconfirmed_date(monkeypatch):
    monkeypatch.setattr(excel_date, 'save_date_report', fail_writes)
    with pytest.raises(DateValidationError, match='Não foi possível gravar') as error:
        raise_date_error()
    assert error.value.report_path is None
    assert 'Relatório técnico local: None' not in str(error.value)
    report = error.value.diagnostic
    assert report['cell'] == 'E8' and report['date_ok'] and not report['display_ok']
    assert report['attempts'][0]['text'] == '46331'


def test_stored_readback_also_preserves_report_without_storage(monkeypatch):
    monkeypatch.setattr(excel_date, 'save_date_report', fail_writes)
    wb = Workbook(); ws = wb.active; ws['E8'] = 46332
    ws['E8'].number_format = 'dd/mm/yyyy'
    with pytest.raises(DateValidationError) as error:
        check_stored_date(ws['E8'], EXPECTED, EPOCH)
    assert error.value.report_path is None and not error.value.diagnostic['date_ok']
    wb.close()


@pytest.mark.parametrize('save_allowed', [True, False])
def test_worker_wrapped_error_can_save_or_copy_without_automatic_json(gui, tmp_path, monkeypatch, save_allowed):
    app, window, messages = gui
    monkeypatch.setattr(excel_date, 'save_date_report', fail_writes)
    target = tmp_path / 'PRIVATE_USER' / 'manual.json'
    if save_allowed: target.parent.mkdir()
    monkeypatch.setattr(QFileDialog, 'getSaveFileName', lambda *args: (str(target), 'JSON (*.json)'))
    dialogs = []
    def inspect_dialog(box):
        report = json.loads(box.detailedText()); dialogs.append(report)
        assert report['cell'] == 'E8' and 'não foi localizado' in box.informativeText()
        next(b for b in box.buttons() if b.text().startswith('Salvar diagnóstico')).click()
        next(b for b in box.buttons() if b.text() == 'Copiar diagnóstico').click()
        return QMessageBox.Close
    monkeypatch.setattr(QMessageBox, 'exec', inspect_dialog)
    def wrapped():
        try: raise_date_error()
        except DateValidationError as exc: raise RuntimeError('Excel não gerado: ' + str(exc)) from exc
    window.run_job(wrapped, lambda result: None); wait_job(app, window)
    assert len(dialogs) == 1 and window.date_diagnostic['content'] == dialogs[0]
    copied = QApplication.clipboard().text()
    assert json.loads(copied) == dialogs[0] and 'PRIVATE_USER' not in copied
    assert target.is_file() == save_allowed
    if save_allowed:
        assert json.loads(target.read_text()) == dialogs[0]
        assert any(title == 'Diagnóstico salvo' for title, _ in messages)
    else:
        assert any('Não foi possível salvar nesta pasta' in text for _, text in messages)
    # A subsequent unrelated failure must not show the previous E8 report.
    def unrelated(): raise RuntimeError('Outra falha sintética')
    window.run_job(unrelated, lambda result: None); wait_job(app, window)
    assert len(dialogs) == 1 and window.date_diagnostic is None


def test_pdf_failure_preserves_excel_and_diagnostic_in_done_callback(gui, monkeypatch):
    app, window, messages = gui
    monkeypatch.setattr(excel_date, 'save_date_report', fail_writes)
    window.items = [complete_item()]; window.refresh()
    def export_failure(*args): raise_date_error()
    monkeypatch.setattr(main_window, 'export_pdf', export_failure)
    captured = []
    def inspect_dialog(box):
        captured.append(json.loads(box.detailedText()))
        return QMessageBox.Close
    monkeypatch.setattr(QMessageBox, 'exec', inspect_dialog)
    window.generate(True); wait_job(app, window)
    assert captured and captured[0]['cell'] == 'E8'
    assert (window.outputs[0] / 'cadastro.xlsx').is_file()
    assert not (window.outputs[0] / 'cadastro.pdf').exists()
    assert 'O PDF não foi gerado' in window.status.text()


def test_copy_also_works_if_automatic_report_disappeared(gui, tmp_path, monkeypatch):
    app, window, messages = gui
    def broken():
        try:
            check_com_date(DateCell('46331'), EXPECTED, EPOCH, phase='teste', report_dir=tmp_path)
        except DateValidationError as exc:
            exc.report_path.unlink()
            raise
    captured = []
    def inspect_dialog(box):
        assert 'não foi localizado' in box.informativeText()
        next(b for b in box.buttons() if b.text() == 'Copiar diagnóstico').click()
        captured.append(json.loads(QApplication.clipboard().text()))
        return QMessageBox.Close
    monkeypatch.setattr(QMessageBox, 'exec', inspect_dialog)
    window.run_job(broken, lambda result: None); wait_job(app, window)
    assert captured[0]['cell'] == 'E8' and not captured[0]['display_ok']


def test_save_and_copy_actions_keep_dialog_open_until_close(gui):
    app, window, messages = gui
    box = main_window.DiagnosticMessageBox(QMessageBox.Warning, 'Diagnóstico', 'E8', QMessageBox.Close, window)
    save = box.addButton('Salvar diagnóstico', QMessageBox.ActionRole)
    copy = box.addButton('Copiar diagnóstico', QMessageBox.ActionRole)
    visible = []
    def click(button):
        button.click(); visible.append(box.isVisible())
    QTimer.singleShot(10, lambda: click(save))
    QTimer.singleShot(20, lambda: click(copy))
    QTimer.singleShot(30, lambda: box.button(QMessageBox.Close).click())
    QTimer.singleShot(1000, box.reject)
    box.exec()
    assert visible == [True, True] and not box.isVisible()
    box.deleteLater()
