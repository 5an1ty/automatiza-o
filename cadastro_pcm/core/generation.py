"""Shared document generation and sequential, independently recorded batch results."""
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import logging

from .excel_writer import write_excel, check_template
from .pdf_exporter import export_pdf
from .file_manager import new_output
from .excel_date import date_error_payload


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
            logging.exception('Falha na exportação PDF')
            error = f'O Excel e o registro foram salvos em {folder}. O PDF não foi gerado: {exc}. Use Exportar PDF do Excel salvo para tentar novamente.'
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
    return root, results
