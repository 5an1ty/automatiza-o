"""Shared document generation and sequential, independently recorded batch results."""
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import logging

from .excel_writer import write_excel, check_template
from .pdf_exporter import export_pdf
from .file_manager import new_output, material_pdf, export_central_pdf
from .pdf_diagnostics import diagnostic_payload as date_error_payload
from .pdf_diagnostics import PdfExportError, unsafe_cleanup, exception_codes


def generate_documents(item, settings, pdf=True, *, writer=None, exporter=None, folder=None, pdf_output=None):
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
    try:
        (folder / 'registro.json').write_text(item.model_dump_json(indent=2), encoding='utf-8')
    except Exception as exc:
        exc.pcm_context = {'stage': 'salvar_registro', 'destination': str(folder / 'registro.json')}
        raise
    error, diagnostic = '', None
    if pdf:
        try:
            export_central_pdf(xlsx, folder, pdf_output or settings.output, settings.sheet,
                               item.name, settings.expected_pages, exporter=exporter)
        except Exception as exc:
            if isinstance(exc, PdfExportError):
                logging.error('Falha PDF na etapa %s', exc.report.get('failure_stage'))
            else:
                logging.error('Falha na exportação PDF: %s', type(exc).__name__)
            outcome = str(exc) if isinstance(exc, PdfExportError) else f'O PDF não foi gerado: {exc}'
            error = f'O Excel e o registro foram salvos em {folder}. {outcome}. Use Exportar PDF do Excel salvo para tentar novamente.'
            diagnostic = local_failure(exc)
    return folder, error, diagnostic


@dataclass
class BatchResult:
    row: int
    item: object
    status: str
    reason: str = ''
    folder: Path | None = None
    diagnostic: object = None
    details: dict = field(default_factory=dict)

    def __post_init__(self):
        folder = Path(self.folder) if self.folder else None
        payload = self.diagnostic or {}
        local = payload.get('local', {})
        report = payload.get('content', {})
        excel = bool(folder and (folder / 'cadastro.xlsx').is_file())
        try:
            pdf = material_pdf(folder) if folder and (folder / 'registro.json').is_file() else None
        except (OSError, ValueError):
            pdf = None
        from .pdf_diagnostics import usable_pdf
        if self.diagnostic:
            report.setdefault('item_id', f'item_{self.row + 1:03d}')
        pdf_state = 'aprovado' if pdf and usable_pdf(pdf) else 'não aprovado'
        if self.status == 'erro' and pdf and usable_pdf(pdf):
            pdf_state = 'falhou; PDF anterior disponível'
        if local.get('pdf_published'):
            pdf_state = 'aprovado; associação falhou'
        self.details = dict(item_id=f'item_{self.row + 1:03d}',
                            excel='gerado' if excel else 'não gerado',
                            pdf=pdf_state,
                            excel_path=str(folder / 'cadastro.xlsx') if folder else '',
                            stage=local.get('stage') or report.get('failure_stage') or
                                  ('concluido' if self.status == 'sucesso' else 'validar_item'),
                            exception_type=local.get('exception_type', ''),
                            message=local.get('message', self.reason),
                            codes=local.get('codes', report.get('exceptions', [])),
                            destination=local.get('destination', str(pdf or '')),
                            temporary_directory=local.get('temporary_directory', ''),
                            retained_diagnostic_directory=local.get('retained_diagnostic_directory', ''))


def local_failure(exc, stage='gerar_excel'):
    """Raw messages/paths stay in the local UI; the exportable report stays sanitized."""
    payload = date_error_payload(exc)
    if payload is None:
        payload = {'content': {'exceptions': exception_codes(exc)}, 'path': None, 'local_only': True}
    context = getattr(exc, 'pcm_context', {})
    if payload.get('local_only'):
        payload['content'].update(stage='falhou', failure_stage=context.get('stage', stage))
    chain, seen = [], set()
    current = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        chain.append(f'{type(current).__name__}: {current}')
        current = current.__cause__ or current.__context__
    payload['local'] = dict(context, stage=payload['content'].get('failure_stage') or context.get('stage', stage),
                            exception_type=type(exc).__name__, message='\nCausado por: '.join(chain),
                            codes=exception_codes(exc))
    return payload


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
        stage = 'validar_modelo'
        try:
            check_template(local.template, item, local.sheet, local.authorized_template_sha256)
            stage = 'criar_pasta_material'
            folder = new_output(root, item.name)
            stage = 'gerar_excel'
            folder, error, diagnostic = generate_documents(item, local, True, folder=folder, pdf_output=settings.output)
            result = BatchResult(row, item, 'erro' if error else 'sucesso', error, folder, diagnostic)
        except Exception as exc:
            logging.error('Falha no item %s do lote: %s', row + 1, type(exc).__name__)
            result = BatchResult(row, item, 'erro', str(exc), folder, local_failure(exc, stage))
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
            def exporter(workbook, destination, sheet, name, pages):
                return export_existing(workbook, destination, sheet, info['sha256'], pages)
            export_central_pdf(source, result.folder, settings.output, settings.sheet,
                               result.item.name, settings.expected_pages, exporter=exporter)
            retried.append(BatchResult(result.row, result.item, 'sucesso', folder=result.folder))
        except Exception as exc:
            payload = local_failure(exc, 'conferir_excel_salvo')
            retried.append(BatchResult(result.row, result.item, 'erro', str(exc), result.folder, payload))
            if unsafe_cleanup(payload):
                remaining = failures[position:]
                retried.extend(BatchResult(r.row, r.item, 'erro', 'Nova tentativa interrompida: Excel não encerrado.', r.folder)
                               for r in remaining)
                break
    return retried
