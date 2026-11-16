from cadastro_pcm.core.records import material_record
"""Shared output organization; real Excel export is simulated on Linux."""
import json
from pathlib import Path
from types import SimpleNamespace
import pymupdf
import pytest
from cadastro_pcm.core import generation, file_manager, pdf_recovery
from cadastro_pcm.core.file_manager import export_central_pdf, material_pdf
from cadastro_pcm.core.pdf_diagnostics import PdfExportError, pending_marker
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core.excel_writer import fingerprint
from test_workflow import template, approved
from test_pdf_pipeline import native


def synthetic_export(source, target, sheet, name, pages):
    with pymupdf.open() as doc:
        page = doc.new_page(); page.insert_text((40, 40), name); doc.save(target)


def config(template, tmp_path):
    return Settings(requester='Revisor', output=str(tmp_path / 'documentos'), template=str(template),
                    authorized_template_sha256=fingerprint(template))


def test_individual_batch_multiple_runs_and_record_links(template, tmp_path, monkeypatch):
    settings = config(template, tmp_path)
    original = template.read_bytes()
    monkeypatch.setattr(generation, 'export_pdf', synthetic_export)
    first, error, _ = generation.generate_documents(approved(), settings)
    assert not error
    excel = (first / 'cadastro.xlsx').read_bytes()
    record = json.loads(material_record(first).read_text())
    path = material_pdf(first); data = path.read_bytes()
    assert path.parent == first and record['pdf_path'] == str(path)
    shared = material_pdf(first, shared=True)
    assert shared.parent == Path(settings.output) / 'lote pdf' and shared.read_bytes() == data
    for _ in range(2):
        root, results = generation.generate_batch([(i, approved()) for i in range(3)], settings)
        assert all(r.status == 'sucesso' for r in results)
        assert all(r.folder.parent == root for r in results)
        assert all(material_pdf(r.folder, shared=True).parent == shared.parent for r in results)
    pdfs = list(shared.parent.iterdir())
    assert len(pdfs) == 7 and all(p.suffix == '.pdf' for p in pdfs)
    assert len({p.name for p in pdfs}) == 7
    assert path.read_bytes() == data and (first / 'cadastro.xlsx').read_bytes() == excel
    assert template.read_bytes() == original
    assert not list(Path(settings.output).rglob('*.pendente'))


def test_reexport_preserves_previous_pdf_excel_and_record_fields(template, tmp_path):
    settings = config(template, tmp_path)
    folder, error, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    assert not error
    source = folder / 'cadastro.xlsx'; excel = source.read_bytes()
    previous = material_pdf(folder); data = previous.read_bytes()
    record = json.loads(material_record(folder).read_text())
    record['custom_metadata'] = {'preservar': True}
    material_record(folder).write_text(json.dumps(record))
    current = export_central_pdf(source, folder, settings.output, settings.sheet, approved().name,
                                 exporter=synthetic_export)
    assert current != previous and material_pdf(folder, shared=True) == current
    assert current.read_bytes() and previous.read_bytes() == data and source.read_bytes() == excel
    changed = json.loads(material_record(folder).read_text())
    assert {k: v for k, v in changed.items() if k not in ('pdf_path', 'pdf_shared_path', 'pdf_exports')} == {k: v for k, v in record.items() if k not in ('pdf_path', 'pdf_shared_path', 'pdf_exports')}


def test_collision_does_not_overwrite(template, tmp_path, monkeypatch):
    settings = config(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, pdf=False)
    shared = Path(settings.output) / 'lote pdf'; shared.mkdir()
    existing = shared / ('cadastro_' + 'b' * 32 + '.pdf'); existing.write_bytes(b'preservar')
    values = iter(['b' * 32, 'c' * 32, 'd' * 32])
    monkeypatch.setattr(file_manager, 'uuid4', lambda: SimpleNamespace(hex=next(values)))
    path = export_central_pdf(folder / 'cadastro.xlsx', folder, settings.output, settings.sheet,
                             approved().name, exporter=synthetic_export)
    assert path.name == 'cadastro_' + 'c' * 32 + '.pdf'
    assert existing.read_bytes() == b'preservar'


@pytest.mark.parametrize('mode', ['clipped', 'blank', 'close_error'])
def test_native_rejected_pdf_never_enters_shared_folder(native, tmp_path, mode):
    folder = native.source.parent
    (folder / 'registro.json').write_text(approved().model_dump_json())
    setattr(native, mode, True)
    with pytest.raises(PdfExportError):
        export_central_pdf(native.source, folder, tmp_path / 'saida', 'Criação de Material Nacional', 'Material teste')
    assert not list((tmp_path / 'saida' / 'lote pdf').iterdir())
    assert list(folder.rglob('reprovado_*.pdf')) and list(folder.rglob('diagnostico_pdf_*.json'))
    assert not json.loads(material_record(folder).read_text()).get('pdf_path')


def test_failed_reexport_keeps_previous_link_and_diagnostics(template, tmp_path):
    settings = config(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    path = material_pdf(folder); before = path.read_bytes()
    record = material_record(folder).read_bytes()
    def fail(source, target, *args):
        target.write_bytes(b'rejected')
        pending_marker(target).write_text('pending')
        raise RuntimeError('Exportação rejeitada')
    with pytest.raises(RuntimeError):
        export_central_pdf(folder / 'cadastro.xlsx', folder, settings.output, settings.sheet, approved().name, exporter=fail)
    assert material_pdf(folder) == path and path.read_bytes() == before
    assert json.loads(material_record(folder).read_text())['generation_state']['pdf'] == 'falhou'
    assert json.loads(material_record(folder).read_text())['pdf_path'] == str(path)
    assert list(folder.rglob('*.pendente'))
    assert list(material_pdf(folder, shared=True).parent.iterdir()) == [material_pdf(folder, shared=True)]


def test_preview_print_and_open_output_use_correct_locations(template, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QApplication, QMessageBox, QDialog
    from PySide6.QtPdf import QPdfDocument
    from cadastro_pcm.interface import main_window
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    app = QApplication.instance() or QApplication([])
    settings = config(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    warnings = []
    monkeypatch.setattr(QMessageBox, 'warning', lambda *args: warnings.append(args[2]) or QMessageBox.Ok)
    window = main_window.Window(); window.items = [approved()]; window.refresh(); window.outputs[0] = folder
    window.update_actions(); expected = material_pdf(folder)
    assert window.pdf_path() == expected and window.preview_button.isEnabled()
    captured = []
    original = QPdfDocument.load
    def load(self, filename):
        captured.append(Path(filename)); return original(self, filename)
    monkeypatch.setattr(QPdfDocument, 'load', load)
    monkeypatch.setattr(QDialog, 'exec', lambda self: 0)
    window.preview(); assert not warnings
    assert captured == [expected]
    opened = []
    monkeypatch.setattr(main_window.QDesktopServices, 'openUrl', lambda url: opened.append(Path(url.toLocalFile())) or True)
    monkeypatch.setattr(main_window, 'os', SimpleNamespace(name='nt'))
    monkeypatch.setattr(QMessageBox, 'information', lambda *args: QMessageBox.Ok)
    window.print_pdf(); assert opened[-1] == expected
    window.table.clearSelection(); window.open_output()
    assert opened[-1] == Path(settings.output).resolve()
    window.close()


def test_existing_excel_reexport_uses_shared_folder_without_destination_picker(template, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QApplication, QMessageBox, QDialog, QFileDialog
    from cadastro_pcm.interface import main_window
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    app = QApplication.instance() or QApplication([])
    settings = config(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, pdf=False)
    source = folder / 'cadastro.xlsx'; original = source.read_bytes()
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    monkeypatch.setattr(pdf_recovery, 'export_pdf', synthetic_export)
    monkeypatch.setattr(QFileDialog, 'getOpenFileName', lambda *args: (str(source), 'Excel'))
    def forbidden(*args):
        raise AssertionError('O destino PDF deve ser automático.')
    monkeypatch.setattr(QFileDialog, 'getSaveFileName', forbidden)
    monkeypatch.setattr(QDialog, 'exec', lambda self: QDialog.Accepted)
    monkeypatch.setattr(QMessageBox, 'question', lambda *args: QMessageBox.No)
    errors = []
    monkeypatch.setattr(QMessageBox, 'warning', lambda *args: errors.append(args[2]))
    window = main_window.Window()
    monkeypatch.setattr(window, 'run_job', lambda task, done: done(task()))
    window.export_existing_excel()
    assert not errors
    assert material_pdf(folder, shared=True).parent == Path(settings.output) / 'lote pdf'
    assert material_pdf(folder).is_file() and source.read_bytes() == original
    window.close()


def test_copy_failure_removes_only_owned_partial_pdf(template, tmp_path, monkeypatch):
    import shutil
    settings = config(template, tmp_path)
    folder, _, _ = generation.generate_documents(approved(), settings, exporter=synthetic_export)
    path = material_pdf(folder); pdf = path.read_bytes(); record = material_record(folder).read_bytes()
    def fail_copy(incoming, target):
        target.write(b'partial'); raise OSError('Disco cheio')
    monkeypatch.setattr(shutil, 'copyfileobj', fail_copy)
    with pytest.raises(OSError, match='Disco cheio'):
        export_central_pdf(folder / 'cadastro.xlsx', folder, settings.output, settings.sheet,
                           approved().name, exporter=synthetic_export)
    assert list(material_pdf(folder, shared=True).parent.iterdir()) == [material_pdf(folder, shared=True)] and path.read_bytes() == pdf
    assert json.loads(material_record(folder).read_text())['generation_state']['pdf'] == 'falhou'
    assert json.loads(material_record(folder).read_text())['pdf_path'] == str(path)
