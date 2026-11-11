"""Native pipeline simulations: no real Microsoft Excel is installed in Linux."""
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
import json
import sys

import pymupdf
import pytest
from openpyxl import Workbook, load_workbook

from cadastro_pcm.core import pdf_exporter, pdf_recovery, excel_runtime, generation
from cadastro_pcm.core.pdf_diagnostics import PdfExportError, pending_marker, usable_pdf
from cadastro_pcm.core.excel_writer import fingerprint, write_excel
from cadastro_pcm.core.settings import Settings
from test_workflow import template, approved


@pytest.fixture
def native(tmp_path, monkeypatch):
    source = tmp_path / 'planilha.xlsx'
    book = Workbook(); sheet = book.active; sheet.title = 'Criação de Material Nacional'
    sheet.print_area = 'A1:E33'; sheet['A12'] = 'Material teste'; sheet['A33'] = 'Controle'
    book.create_sheet('Outra aba')['A1'] = 'Nao exportar'
    book.save(source); book.close()
    state = SimpleNamespace(source=source, pages=1, error=None, empty=False, missing=False,
                            clipped=False, closed=False, close_error=False, selected=[], called=False)
    class Sheet:
        PageSetup = SimpleNamespace(LeftMargin=36, RightMargin=36, TopMargin=36, BottomMargin=36)
        ProtectContents = False
        def Range(self, address):
            value = {'A12': 'Material teste', 'A33': 'Controle'}.get(address)
            return SimpleNamespace(Value=value, Text=value or '', Width=500, Height=700)
        def Select(self, replace):
            assert replace is True; state.single_sheet = True
        def Calculate(self): state.calculated = True
        def ResetAllPageBreaks(self): pass
        def ExportAsFixedFormat(self, *args):
            state.called = True
            if state.error: raise state.error
            if state.missing: return
            target = Path(args[1]); assert target.is_absolute()
            if state.empty:
                target.write_bytes(b''); return
            document = pymupdf.open()
            for _ in range(state.pages):
                page = document.new_page(width=595.28, height=841.89)
                page.insert_text((40, 40), 'Material teste')
                if not state.clipped: page.insert_text((40, 60), 'Controle')
            document.save(target); document.close()
    @contextmanager
    def session(path, read_only):
        assert Path(path) != source
        def select(name):
            state.selected.append(name); return Sheet()
        try:
            yield None, SimpleNamespace(Worksheets=select)
        finally:
            state.closed = True
            if state.close_error: raise RuntimeError('close failed')
    monkeypatch.setattr(pdf_exporter, 'excel_session', session)
    monkeypatch.setattr(pdf_exporter, 'sys', SimpleNamespace(platform='win32'))
    return state


def run(native, target):
    return pdf_exporter.export_pdf(native.source, target, 'Criação de Material Nacional', 'Material teste')


def test_success_absolute_space_path_and_safe_report(native, tmp_path):
    original = native.source.read_bytes()
    target = tmp_path / 'pasta com espaços' / 'material.pdf'
    assert run(native, target) == target.resolve()
    assert native.closed and native.selected == ['Criação de Material Nacional']
    assert native.single_sheet and native.calculated
    assert usable_pdf(target) and native.source.read_bytes() == original
    report = json.loads(next((target.parent / 'diagnostico_pdf').glob('*.json')).read_text())
    assert report['pages'] == 1 and report['stage'] == 'concluido'
    text = json.dumps(report)
    assert 'Material teste' not in text and str(tmp_path) not in text
    assert any(e['stage'] == 'configurar_paginacao' and e['page_setup']['PaperSize'] == 9 for e in report['events'])


@pytest.mark.parametrize('mode,stage,code', [('pages', 'validar_pdf', 'page_count'),
    ('clipped', 'validar_pdf', 'missing_content'), ('empty', 'conferir_arquivo', 'conferir_arquivo'),
    ('missing', 'conferir_arquivo', 'conferir_arquivo'), ('error', 'exportar_com', 'exportar_com'),
    ('close_error', 'encerrar_excel', 'encerrar_excel')])
def test_failures_different_stages_preserve_evidence(native, tmp_path, mode, stage, code):
    if mode == 'pages': native.pages = 2
    elif mode == 'error':
        class ComError(Exception): hresult = -2147352567; excepinfo = (None, None, 'private corporate text', None, None, -2146827284)
        native.error = ComError('secret client path')
    else: setattr(native, mode, True)
    target = tmp_path / 'failed.pdf'
    with pytest.raises(PdfExportError) as failure: run(native, target)
    report = failure.value.report
    assert report['failure_stage'] == stage and report['failure_code'] == code
    assert failure.value.report_path.is_file() and native.closed
    assert not usable_pdf(target)
    assert pending_marker(target).is_file()
    if mode == 'pages':
        assert 'PDF gerado, mas não aprovado na validação.' in str(failure.value)
        assert report['pages'] == 2
    if mode not in ('missing', 'error'):
        assert list((tmp_path / 'diagnostico_pdf').glob('reprovado_*.pdf'))
    safe = json.dumps(report)
    assert 'secret client' not in safe and 'private corporate' not in safe
    if mode == 'error': assert report['exceptions'][0]['hresult'] == '0x80020009'


def test_retry_does_not_delete_or_modify_excel(native, tmp_path):
    original = native.source.read_bytes(); target = tmp_path / 'retry.pdf'; native.pages = 2
    with pytest.raises(PdfExportError): run(native, target)
    rejected = next((tmp_path / 'diagnostico_pdf').glob('reprovado_*.pdf'))
    evidence = rejected.read_bytes(); native.pages = 1
    run(native, target)
    assert usable_pdf(target) and rejected.read_bytes() == evidence
    assert native.source.read_bytes() == original


def test_locked_existing_destination_is_never_overwritten(native, tmp_path):
    target = tmp_path / 'locked.pdf'; target.write_bytes(b'existing')
    with pytest.raises(PdfExportError, match='já existe'): run(native, target)
    assert target.read_bytes() == b'existing' and not native.called
    assert not pending_marker(target).exists()


def test_unwritable_destination_reports_io_error(native, tmp_path, monkeypatch):
    target = tmp_path / 'denied' / 'test.pdf'
    mkdir = Path.mkdir
    def deny(self, *args, **kwargs):
        if self == target.parent or target.parent in self.parents: raise PermissionError(13, 'permission denied')
        return mkdir(self, *args, **kwargs)
    monkeypatch.setattr(Path, 'mkdir', deny)
    with pytest.raises(PdfExportError) as error: run(native, target)
    assert error.value.report['failure_stage'] == 'preparar_arquivos'
    assert error.value.report['exceptions'][0]['errno'] == 13
    assert not native.called


def test_print_area_blocker_before_com(native, tmp_path):
    book = load_workbook(native.source); book.active['G40'] = 'must not be cut'; book.save(native.source); book.close()
    with pytest.raises(PdfExportError) as error: run(native, tmp_path / 'invalid.pdf')
    assert error.value.report['failure_code'] == 'print_area'
    assert not native.called


def test_generated_workbook_and_new_manual_version(template, tmp_path, monkeypatch):
    folder = tmp_path / 'material'; folder.mkdir()
    source = folder / 'cadastro.xlsx'; item = approved()
    write_excel(template, source, item, 'Criação de Material Nacional', fingerprint(template))
    (folder / 'registro.json').write_text(item.model_dump_json())
    info = pdf_recovery.existing_workbook(source)
    book = load_workbook(source); book.active['A20'] = 77; book.save(source); book.close()
    with pytest.raises(ValueError, match='mudou'): pdf_recovery.export_existing(source, folder / 'pdf.pdf', 'Criação de Material Nacional', info['sha256'])
    latest = pdf_recovery.existing_workbook(source); before = source.read_bytes(); calls = []
    monkeypatch.setattr(pdf_recovery, 'export_pdf', lambda *args: calls.append(args) or args[1])
    pdf_recovery.export_existing(source, folder / 'pdf.pdf', 'Criação de Material Nacional', latest['sha256'])
    assert len(calls) == 1 and source.read_bytes() == before
    with pytest.raises(ValueError, match='Aba'): pdf_recovery.export_existing(source, folder / 'pdf.pdf', 'Outra', latest['sha256'])


def test_synthetic_test_cleans_all_files(monkeypatch):
    visited = []
    def export(source, target, sheet, name, report):
        assert Path(source).is_file(); visited.append(Path(source).parent)
        report['stage'] = 'concluido'; Path(target).write_bytes(b'synthetic')
    monkeypatch.setattr(pdf_recovery, 'export_pdf', export)
    result = pdf_recovery.synthetic_export_test()
    assert result['ok'] and result['report']['stage'] == 'concluido'
    assert all(not path.exists() for path in visited)
    def fail(source, *args, **kwargs):
        visited.append(Path(source).parent); raise RuntimeError('failure')
    monkeypatch.setattr(pdf_recovery, 'export_pdf', fail)
    assert not pdf_recovery.synthetic_export_test()['ok']
    assert all(not path.exists() for path in visited)


def test_cleanup_errors_keep_primary_failure_and_attempt_all(monkeypatch):
    calls = []
    def fail_close(*args): calls.append('close'); raise RuntimeError('private')
    book = SimpleNamespace(Close=fail_close)
    excel = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda *args, **kwargs: book), Quit=lambda: calls.append('quit'))
    client = SimpleNamespace(DispatchEx=lambda *args: excel)
    monkeypatch.setitem(sys.modules, 'pythoncom', SimpleNamespace(CoInitialize=lambda: calls.append('init'), CoUninitialize=lambda: calls.append('uninit')))
    monkeypatch.setitem(sys.modules, 'win32com', SimpleNamespace(client=client))
    monkeypatch.setitem(sys.modules, 'win32com.client', client)
    monkeypatch.setattr(excel_runtime, 'sys', SimpleNamespace(platform='win32'))
    with pytest.raises(ValueError, match='primary') as failure:
        with excel_runtime.excel_session('file.xlsx'): raise ValueError('primary')
    assert calls == ['init', 'close', 'quit', 'uninit']
    assert failure.value.excel_cleanup_errors[0]['operation'] == 'Close'


def test_retry_only_failed_pdfs_preserves_workbooks(template, tmp_path, monkeypatch):
    settings = Settings(output=str(tmp_path / 'out'), template=str(template), authorized_template_sha256=fingerprint(template))
    def export(source, target, *args): raise RuntimeError('PDF synthetic failure')
    monkeypatch.setattr(generation, 'export_pdf', export)
    root, results = generation.generate_batch([(0, approved()), (1, approved())], settings)
    before = {r.folder / 'cadastro.xlsx': (r.folder / 'cadastro.xlsx').read_bytes() for r in results}
    def retry(source, target, *args):
        Path(target).write_bytes(b'validated synthetic PDF'); return target
    monkeypatch.setattr(pdf_recovery, 'export_pdf', retry)
    retried = generation.retry_batch_pdfs(results, settings)
    assert all(r.status == 'sucesso' for r in retried)
    assert all(path.read_bytes() == value for path, value in before.items())
    assert {r.folder for r in retried} == {r.folder for r in results}


def test_unsafe_excel_quit_stops_batch(template, tmp_path, monkeypatch):
    settings = Settings(output=str(tmp_path / 'out'), template=str(template), authorized_template_sha256=fingerprint(template))
    calls = []
    def export(*args):
        calls.append(1)
        raise PdfExportError('Excel não encerrado', {'exceptions': [{'cleanup': [{'operation': 'Quit'}]}]})
    monkeypatch.setattr(generation, 'export_pdf', export)
    _, results = generation.generate_batch([(0, approved()), (1, approved())], settings)
    assert len(calls) == 1
    assert [r.status for r in results] == ['erro', 'ignorado']
    assert 'interrompido' in results[1].reason


def test_pdf_test_menu_and_report_ui(monkeypatch, tmp_path, template):
    from PySide6.QtCore import QEventLoop
    from PySide6.QtWidgets import QApplication, QDialog, QMenu
    from time import monotonic
    from cadastro_pcm.interface import main_window
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    app = QApplication.instance() or QApplication([])
    monkeypatch.setattr(main_window, 'load_settings', lambda: Settings(requester='Revisor'))
    window = main_window.Window()
    assert any(action.text() == 'Testar exportação PDF' for menu in window.findChildren(QMenu) for action in menu.actions())
    monkeypatch.setattr(pdf_recovery, 'synthetic_export_test', lambda: {'ok': True, 'report': {'stage': 'concluido'}})
    captured = []
    monkeypatch.setattr(QDialog, 'exec', lambda dialog: captured.append(dialog.windowTitle()) or 1)
    window.test_pdf_export()
    deadline = monotonic() + 5
    while window.worker is not None and monotonic() < deadline:
        app.processEvents(QEventLoop.AllEvents, 20)
    assert window.worker is None and 'Teste de exportação PDF' in captured
    window.close()


def test_unapproved_pdf_never_printable_after_quarantine_failure(native, tmp_path, monkeypatch):
    native.pages = 2
    replace = Path.replace
    def locked(self, target):
        if self.suffix == '.pdf': raise PermissionError(13, 'file locked')
        return replace(self, target)
    monkeypatch.setattr(Path, 'replace', locked)
    target = tmp_path / 'pending.pdf'
    with pytest.raises(PdfExportError): run(native, target)
    assert target.is_file() and not usable_pdf(target)
