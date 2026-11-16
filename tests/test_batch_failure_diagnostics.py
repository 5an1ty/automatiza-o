from cadastro_pcm.core.records import material_record
"""Synthetic regressions; no corporate inputs or actual Windows Excel available."""
import json
from contextlib import contextmanager
from pathlib import Path
import pytest
from cadastro_pcm.core import generation, pdf_exporter
from cadastro_pcm.core.file_manager import export_central_pdf, material_pdf
from cadastro_pcm.core.pdf_diagnostics import PdfExportError
from cadastro_pcm.core.settings import Settings
from cadastro_pcm.core.excel_writer import fingerprint
from test_workflow import template, approved
from test_pdf_pipeline import native
from test_central_pdfs import synthetic_export


@pytest.mark.parametrize('exists', [False, True])
def test_fifteen_special_and_duplicate_names_short_absolute_destinations(template, tmp_path, monkeypatch, exists):
    output = tmp_path / ('caminho_de_saida_' * 5)
    shared = output / 'lote pdf'
    if exists:
        shared.mkdir(parents=True)
        (shared / 'anterior.pdf').write_bytes(b'nao apagar')
    settings = Settings(output=str(output), template=str(template), authorized_template_sha256=fingerprint(template))
    names = ['Válvula 1½" / 90° <inox> | ? : * ' + 'AISI304 ' * 12] * 15
    destinations = []
    def export(source, target, sheet, name, pages):
        assert shared.is_dir() and source.is_absolute() and target.is_absolute()
        assert target.parent != source.parent and len(str(target)) < 218
        assert target.parent != shared  # Unapproved documents never enter the shared directory.
        destinations.append(target)
        synthetic_export(source, target, sheet, name, pages)
    monkeypatch.setattr(generation, 'export_pdf', export)
    original = template.read_bytes()
    _, results = generation.generate_batch([(n, approved().model_copy(update={'name': name})) for n, name in enumerate(names)], settings)
    assert len(results) == 15 and all(r.status == 'sucesso' for r in results)
    assert all(r.details['excel'] == 'gerado' and r.details['pdf'] == 'aprovado' for r in results)
    assert len({material_pdf(r.folder) for r in results}) == 15
    assert all(material_pdf(r.folder, shared=True).parent == shared for r in results)
    assert all((r.folder / 'cadastro.xlsx').is_file() and material_record(r.folder).is_file() for r in results)
    assert len(destinations) == 15 and template.read_bytes() == original
    if exists: assert (shared / 'anterior.pdf').read_bytes() == b'nao apagar'


def test_nested_path_regression_reproduced_then_short_staging_succeeds(native, tmp_path, monkeypatch):
    original_session = pdf_exporter.excel_session
    @contextmanager
    def excel_with_legacy_path_limit(path, read_only):
        with original_session(path, read_only) as (excel, book):
            worksheets = book.Worksheets
            def sheet(name):
                ws = worksheets(name)
                export = ws.ExportAsFixedFormat
                def limited(*args):
                    if len(str(path)) > 218 or len(args[1]) > 218:
                        raise OSError('Excel: caminho longo para ExportAsFixedFormat')
                    return export(*args)
                ws.ExportAsFixedFormat = limited
                return ws
            book.Worksheets = sheet
            yield excel, book
    monkeypatch.setattr(pdf_exporter, 'excel_session', excel_with_legacy_path_limit)
    material = tmp_path / ('lote_' + 'y' * 35) / ('material_' + 'x' * 90)
    material.mkdir(parents=True)
    old = material / 'exportacao_pdf_1234567890' / 'cadastro.pdf'
    assert len(str(old)) > 218
    with pytest.raises(PdfExportError, match='caminho longo'):
        pdf_exporter.export_pdf(native.source, old, 'Criação de Material Nacional', 'Material teste')
    (material / 'registro.json').write_text(approved().model_dump_json())
    source = native.source.read_bytes()
    target = export_central_pdf(native.source, material, tmp_path, 'Criação de Material Nacional', 'Material teste')
    assert target.parent == tmp_path / 'lote pdf' and target.is_file()
    assert native.source.read_bytes() == source
    assert list(material.rglob('diagnostico_pdf_*.json'))  # Previous evidence untouched.


def test_single_com_failure_continues_with_exact_local_details_and_safe_logs(template, tmp_path, monkeypatch, caplog):
    settings = Settings(output=str(tmp_path / 'saida'), template=str(template), authorized_template_sha256=fingerprint(template))
    class ComError(Exception):
        hresult = -2147352567
        excepinfo = (None, None, 'mensagem COM completa confidencial', None, None, -2146827284)
    def export(source, destination, sheet, name, pages):
        if name == 'Item privado falho':
            cause = ComError('mensagem privada de ExportAsFixedFormat')
            raise PdfExportError('Falha na exportação PDF — etapa exportar_com.',
                {'failure_stage': 'exportar_com', 'exceptions': [{'hresult': '0x80020009'}]}) from cause
        synthetic_export(source, destination, sheet, name, pages)
    monkeypatch.setattr(generation, 'export_pdf', export)
    entries = [(n, approved().model_copy(update={'name': 'Item privado falho' if n == 7 else f'Item {n}'})) for n in range(15)]
    _, results = generation.generate_batch(entries, settings)
    assert len(results) == 15 and sum(r.status == 'sucesso' for r in results) == 14
    failed = results[7]; details = failed.details
    assert failed.status == 'erro' and details['excel'] == 'gerado' and details['pdf'] == 'não aprovado'
    assert details['item_id'] == 'item_008' and details['stage'] == 'exportar_com'
    assert details['exception_type'] == 'PdfExportError' and 'mensagem privada' in details['message']
    assert any(c.get('hresult') == '0x80020009' for c in details['codes'])
    assert Path(details['destination']).is_absolute()
    assert 'privada' not in caplog.text and 'Item privado falho' not in caplog.text
    assert 'privada' not in json.dumps(failed.diagnostic['content'])
    assert len(list((Path(settings.output) / 'lote pdf').glob('*.pdf'))) == 14


def test_central_directory_failure_is_distinguished_from_excel_failure(template, tmp_path, monkeypatch):
    settings = Settings(output=str(tmp_path), template=str(template), authorized_template_sha256=fingerprint(template))
    (tmp_path / 'lote pdf').write_text('arquivo existente, não substituir')
    monkeypatch.setattr(generation, 'export_pdf', lambda *args: pytest.fail('COM não deve iniciar'))
    _, results = generation.generate_batch([(0, approved())], settings)
    result = results[0]
    assert result.status == 'erro' and result.details['stage'] == 'criar_diretorio_pdf'
    assert result.details['excel'] == 'gerado' and result.details['pdf'] == 'não aprovado'
    assert result.details['exception_type'] == 'FileExistsError'
    assert (tmp_path / 'lote pdf').read_text() == 'arquivo existente, não substituir'


def test_results_table_has_full_tooltip_and_local_detail_panel(template, tmp_path, monkeypatch):
    from PySide6.QtWidgets import QApplication, QDialog, QTableWidget, QTextEdit
    from cadastro_pcm.interface import main_window
    monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
    app = QApplication.instance() or QApplication([])
    settings = Settings(requester='Revisor', output=str(tmp_path))
    monkeypatch.setattr(main_window, 'load_settings', lambda: settings)
    window = main_window.Window()
    message = 'Mensagem integral confidencial ' * 50
    payload = generation.local_failure(OSError(13, message), 'publicar_pdf')
    result = generation.BatchResult(2, approved(), 'erro', message, None, payload)
    def inspect(dialog):
        table = dialog.findChild(QTableWidget)
        assert table.columnCount() == 7
        assert message in table.item(0, 0).toolTip()
        table.setCurrentCell(0, 0)
        panel = dialog.findChild(QTextEdit)
        assert message in panel.toPlainText() and 'item_003' in panel.toPlainText()
        return QDialog.Rejected
    monkeypatch.setattr(QDialog, 'exec', inspect)
    window.show_batch_results(tmp_path, [result], settings)
    window.close()
