"""Shared document generation and sequential, independently recorded batch results."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import logging

from .excel_writer import write_excel, check_template
from .pdf_exporter import export_pdf
from .file_manager import new_output
from .pdf_diagnostics import diagnostic_payload as date_error_payload
from .pdf_diagnostics import PdfExportError, unsafe_cleanup


def generate_documents(item, settings, pdf=True, *, writer=None, exporter=None, folder=None):
    writer = writer or write_excel
    exporter = exporter or export_pdf
    item.require_ready()
    folder = Path(folder) if folder is not None else new_output(settings.output, item.name)
    try:
        xlsx = writer(settings.template, folder / 'cadastro.xlsx', item, settings.sheet,
                           settings.authorized_template_sha256)
    except Exception as exc:
        raise RuntimeError(f'Excel não gerado. Confira modelo, acesso ao Microsoft Excel e permissões de gravação: {exc}') from exc
    if not Path(xlsx).is_file():
        raise RuntimeError('O arquivo Excel não foi criado.')
    (folder / 'registro.json').write_text(item.model_dump_json(indent=2), encoding='utf-8')
    error, diagnostic = '', None
    if pdf:
        try:
            exporter(xlsx, folder / 'cadastro.pdf', settings.sheet, item.name, settings.expected_pages)
            if not (folder / 'cadastro.pdf').is_file():
                raise RuntimeError('O arquivo PDF não foi criado.')
        except Exception as exc:
            if isinstance(exc, PdfExportError):
                logging.error('Falha PDF na etapa %s', exc.report.get('failure_stage'))
            else:
                logging.error('Falha na exportação PDF: %s', type(exc).__name__)
            outcome = str(exc) if isinstance(exc, PdfExportError) else f'O PDF não foi gerado: {exc}'
            error = f'O Excel e o registro foram salvos em {folder}. {outcome}. Use Exportar PDF do Excel salvo para tentar novamente.'
            diagnostic = date_error_payload(exc)
    return folder, error, diagnostic


@dataclass
class BatchResult:
    row: int
    item: object
    status: str
    reason: str = ''
    folder: Path | None = None
    diagnostic: object = None


def eligibility(item):
    try:
        item.require_ready()
    except ValueError as exc:
        return str(exc)
    return ''


def generate_batch(entries, settings, progress=lambda *args: None, root=None):
    """Entries are immutable snapshots; COM is owned by the existing per-call sessions."""
    root = Path(root) if root else new_output(settings.output, 'lote_' + datetime.now().strftime('%Y-%m-%d_%H%M'))
    local = settings.model_copy(deep=True)
    local.output = str(root)
    results = []
    total = sum(not eligibility(item) for _, item in entries)
    position = 0
    for row, item in entries:
        reason = eligibility(item)
        if reason:
            results.append(BatchResult(row, item, 'ignorado', reason))
            continue
        position += 1
        progress(position, total, item.name)
        folder = None
        try:
            check_template(local.template, item, local.sheet, local.authorized_template_sha256)
            folder = new_output(root, item.name)
            folder, error, diagnostic = generate_documents(item, local, True, folder=folder)
            result = BatchResult(row, item, 'erro' if error else 'sucesso', error, folder, diagnostic)
        except Exception as exc:
            logging.exception('Falha no item %s do lote', row + 1)
            result = BatchResult(row, item, 'erro', str(exc), folder, date_error_payload(exc))
        results.append(result)
        if unsafe_cleanup(result.diagnostic):
            processed = {r.row for r in results}
            results.extend(BatchResult(r, i, 'ignorado', 'Lote interrompido: não foi possível encerrar a instância Excel.')
                           for r, i in entries if r not in processed)
            break
    return root, results


def retry_batch_pdfs(results, settings, progress=lambda *args: None):
    """Retry saved workbooks only; never refill or replace successful documents."""
    from .pdf_recovery import existing_workbook, export_existing
    retried = []
    failures = [r for r in results if r.status == 'erro' and r.folder and (r.folder / 'cadastro.xlsx').is_file()]
    for position, result in enumerate(failures, 1):
        progress(position, len(failures), result.item.name)
        try:
            source = result.folder / 'cadastro.xlsx'
            info = existing_workbook(source)
            export_existing(source, result.folder / 'cadastro.pdf', settings.sheet, info['sha256'], settings.expected_pages)
            retried.append(BatchResult(result.row, result.item, 'sucesso', folder=result.folder))
        except Exception as exc:
            payload = date_error_payload(exc)
            retried.append(BatchResult(result.row, result.item, 'erro', str(exc), result.folder, payload))
            if unsafe_cleanup(payload):
                remaining = failures[position:]
                retried.extend(BatchResult(r.row, r.item, 'erro', 'Nova tentativa interrompida: Excel não encerrado.', r.folder)
                               for r in remaining)
                break
    return retried
