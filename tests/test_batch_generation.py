"""Batch regressions use synthetic workbooks; PDF/COM are explicitly simulated."""
from pathlib import Path
from time import monotonic
from types import SimpleNamespace
import sys
import threading

import pytest
from openpyxl import load_workbook
from PySide6.QtCore import QItemSelectionModel, QEventLoop
from PySide6.QtWidgets import QApplication, QMessageBox

from cadastro_pcm.core import generation, excel_runtime
from cadastro_pcm.core.excel_writer import fingerprint
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.interface import main_window
from test_workflow import template, approved


@pytest.fixture
def settings(template, tmp_path):
    return Settings(template=str(template), authorized_template_sha256=fingerprint(template),
                    requester='Teste', output=str(tmp_path / 'saida'))


def fake_pdf(xlsx, target, *args):
    Path(target).write_bytes(b'%PDF-1.4 synthetic export stub')


@pytest.mark.parametrize('count', [1, 3])
def test_batch_independent_files_names_and_original(count, settings, monkeypatch):
    original = Path(settings.template).read_bytes()
    monkeypatch.setattr(generation, 'export_pdf', fake_pdf)
    entries = [(n, approved().model_copy(update={'name': 'CON' if n == 0 else 'A/B'})) for n in range(count)]
    progress = []
    root, results = generation.generate_batch(entries, settings, lambda *args: progress.append(args))
    assert len({r.folder for r in results}) == count
    assert all(r.status == 'sucesso' and r.folder.parent == root for r in results)
    assert [p[0] for p in progress] == list(range(1, count + 1))
    for result in results:
        assert all((result.folder / f).is_file() for f in ('cadastro.xlsx', 'cadastro.pdf', 'registro.json'))
        book = load_workbook(result.folder / 'cadastro.xlsx')
        assert book.active['C16'].value is None
        assert book.active['A27'].value is None
        assert book.active.protection.sheet
        assert book.active['A1'].value == 'INSTITUCIONAL'
        book.close()
    assert Path(settings.template).read_bytes() == original
    other, _ = generation.generate_batch(entries, settings)
    assert other != root


def test_pending_failure_continuation_and_retry(settings, monkeypatch):
    calls = []
    def export(xlsx, target, *args):
        calls.append(args[1])
        if args[1] == 'Falha':
            raise RuntimeError('Excel indisponível para este item')
        fake_pdf(xlsx, target)
    monkeypatch.setattr(generation, 'export_pdf', export)
    items = [approved(), approved().model_copy(update={'approved': False}),
             approved().model_copy(update={'name': 'Falha'}), approved()]
    root, results = generation.generate_batch(list(enumerate(items)), settings)
    assert [r.status for r in results] == ['sucesso', 'ignorado', 'erro', 'sucesso']
    assert not items[1].approved and len(calls) == 3
    assert 'aprove' in results[1].reason.lower()
    failed = results[2]
    assert (failed.folder / 'cadastro.xlsx').is_file() and not (failed.folder / 'cadastro.pdf').exists()
    preserved = {p: p.read_bytes() for r in results if r.folder for p in r.folder.iterdir()}
    monkeypatch.setattr(generation, 'export_pdf', fake_pdf)
    _, retried = generation.generate_batch([(r.row, r.item) for r in results if r.status == 'erro'], settings, root=root)
    assert len(retried) == 1 and retried[0].row == 2 and retried[0].status == 'sucesso'
    assert all(p.read_bytes() == data for p, data in preserved.items())


def test_missing_pdf_is_failure(settings, monkeypatch):
    monkeypatch.setattr(generation, 'export_pdf', lambda *args: None)
    _, results = generation.generate_batch([(0, approved())], settings)
    assert results[0].status == 'erro' and 'não foi criado' in results[0].reason


def test_invalid_template_is_individual_failure(settings, monkeypatch):
    monkeypatch.setattr(generation, 'export_pdf', fake_pdf)
    original = generation.check_template
    def check(path, item, *args):
        if item.name == 'Falha': raise ValueError('Lista incompatível')
        return original(path, item, *args)
    monkeypatch.setattr(generation, 'check_template', check)
    _, results = generation.generate_batch([(0, approved().model_copy(update={'name': 'Falha'})), (1, approved())], settings)
    assert [r.status for r in results] == ['erro', 'sucesso']
    assert results[0].reason == 'Lista incompatível'


def test_batch_com_sessions_cleanup_on_error_and_continue(settings, monkeypatch):
    calls = []
    active = []
    pythoncom = SimpleNamespace(CoInitialize=lambda: calls.append(('init', threading.get_ident())),
                                CoUninitialize=lambda: calls.append(('uninit', threading.get_ident())))
    def dispatch(name):
        assert not active
        active.append(1)
        def quit_excel():
            active.clear(); calls.append(('quit', threading.get_ident()))
        return SimpleNamespace(Quit=quit_excel)
    client = SimpleNamespace(DispatchEx=dispatch)
    monkeypatch.setitem(sys.modules, 'pythoncom', pythoncom)
    monkeypatch.setitem(sys.modules, 'win32com', SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, 'win32com.client', client)
    monkeypatch.setattr(sys, 'platform', 'win32')
    def export(xlsx, target, sheet, name, *args):
        with excel_runtime.excel_session():
            if name == 'Falha': raise RuntimeError('COM synthetic failure')
            fake_pdf(xlsx, target)
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch([(0, approved().model_copy(update={'name': 'Falha'})), (1, approved())], settings)
    assert [r.status for r in results] == ['erro', 'sucesso']
    assert [c[0] for c in calls] == ['init', 'quit', 'uninit'] * 2 and not active


def test_gui_multiselection_confirmation_thread_progress(settings, monkeypatch):
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    app = QApplication.instance() or QApplication([])
    window = main_window.Window()
    window.items = [approved(), approved().model_copy(update={'approved': False}), approved()]
    window.refresh()
    assert window.batch_button.text().endswith('(2)')
    selection = window.table.selectionModel()
    selection.clearSelection()
    for row in (0, 2):
        selection.select(window.table.model().index(row, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
    assert window.selected_rows() == [0, 2]
    captured = []
    threads = []
    gate = threading.Event()
    def export(*args):
        threads.append(threading.get_ident())
        assert gate.wait(5)
        fake_pdf(*args)
    monkeypatch.setattr(generation, 'export_pdf', export)
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.No)
    window.generate_selected()
    assert window.worker is None and not Path(settings.output).exists()
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.Yes)
    monkeypatch.setattr(window, 'show_batch_results', lambda *args: captured.append(args))
    window.generate_selected()
    deadline = monotonic() + 8
    observed = False
    while window.worker is not None and monotonic() < deadline:
        app.processEvents(QEventLoop.AllEvents, 20)
        if 'Gerando cadastro' in window.status.text():
            observed = True; gate.set()
    assert window.worker is None and observed
    assert len(captured[0][1]) == 2
    assert all(t != threading.get_ident() for t in threads)
    assert set(window.outputs) == {0, 2}
    assert window.centralWidget().isEnabled()
    window.close()


def test_results_dialog_retry_only_failures(settings, monkeypatch, tmp_path):
    from PySide6.QtWidgets import QDialog, QPushButton
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    app = QApplication.instance() or QApplication([])
    window = main_window.Window()
    results = [generation.BatchResult(0, approved(), 'sucesso', folder=tmp_path),
               generation.BatchResult(1, approved(), 'erro', 'Falha de teste', tmp_path),
               generation.BatchResult(2, approved(), 'ignorado', 'Sem aprovação')]
    captured = []
    monkeypatch.setattr(window, 'start_batch', lambda *args: captured.append(args))
    def accept_retry(dialog):
        retry = next(b for b in dialog.findChildren(QPushButton) if b.text() == 'Tentar novamente os itens com falha')
        assert retry.isEnabled()
        return QDialog.Accepted
    monkeypatch.setattr(QDialog, 'exec', accept_retry)
    window.show_batch_results(tmp_path, results, settings)
    assert captured == [([1], tmp_path, settings)]
    window.close()


def test_pending_fields_not_eligible_even_when_approved(settings, monkeypatch):
    monkeypatch.setattr(generation, 'export_pdf', fake_pdf)
    item = approved().model_copy(update={'unit_price': None})
    _, results = generation.generate_batch([(0, item)], settings)
    assert results[0].status == 'ignorado' and 'Preço' in results[0].reason
    assert item.approved and results[0].folder is None


def test_excel_failure_continues_without_exporting_failed_item(settings, monkeypatch):
    writer = generation.write_excel
    exports = []
    def write(source, target, item, *args):
        if item.name == 'Falha': raise RuntimeError('Gravação recusada')
        return writer(source, target, item, *args)
    def export(*args):
        exports.append(args[3]); fake_pdf(*args)
    monkeypatch.setattr(generation, 'write_excel', write)
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch([(0, approved().model_copy(update={'name': 'Falha'})), (1, approved())], settings)
    assert [r.status for r in results] == ['erro', 'sucesso']
    assert 'Gravação recusada' in results[0].reason
    assert exports == ['Gaxeta 16']
