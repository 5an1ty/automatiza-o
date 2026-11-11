import sys
import shutil
import tempfile
import logging
import re
from datetime import date, datetime
from pathlib import Path
import pymupdf
from openpyxl.utils.datetime import from_excel
from .template_reader import read_template
from .excel_runtime import excel_session
from .print_layout import print_area, configure_page, check_print_scale
from .excel_date import check_com_date
from .pdf_diagnostics import event, exception_codes, save_diagnostic, PdfExportError, pending_marker
from .print_layout import PROPOSED
from uuid import uuid4


def validate_pdf(path, name, expected_pages=1, expected_texts=(), require_a4=False, expected_date=None):
    if not Path(path).is_file():
        raise ValueError('O Excel não criou o PDF.')
    with pymupdf.open(path) as pdf:
        if len(pdf) != expected_pages:
            raise ValueError(f'PDF com {len(pdf)} páginas; esperado {expected_pages}. Revise limites, escala e quebras no Excel.')
        texts = [page.get_text() for page in pdf]
        if any(not text.strip() for text in texts):
            raise ValueError('PDF contém página sem texto; confira páginas vazias e impressão.')
        if ' '.join(name.split()).casefold() not in ' '.join(' '.join(texts).split()).casefold():
            raise ValueError('Nome do material não encontrado no PDF exportado.')
        combined = ' '.join(' '.join(texts).split()).casefold()
        normalized = [(address, ' '.join(text.split()).casefold()) for address, text in expected_texts]
        counts = {text: sum(value.count(text) for _, value in normalized)
                  for _, text in normalized if text}
        for address, text in normalized:
            # Repeated labels/values must remain present at least as often as in
            # the source, so an intact duplicate cannot conceal a clipped cell.
            if text and combined.count(text) < counts[text]:
                raise ValueError(f'Conteúdo de {address} ausente ou cortado no PDF. Confira mesclagens e alturas de linha no Excel.')
        if expected_date and expected_date not in combined:
            raise ValueError('Data brasileira de E8 não encontrada no PDF exportado.')
        for page in pdf:
            if require_a4 and (abs(page.rect.width - 595.28) > 2 or abs(page.rect.height - 841.89) > 2):
                raise ValueError('PDF não está em A4 vertical.')
            if require_a4:
                for block in page.get_text('dict')['blocks']:
                    for line in block.get('lines', []):
                        if expected_date and expected_date in ''.join(span['text'] for span in line['spans']):
                            if any(span['size'] < 8 for span in line['spans'] if any(c.isdecimal() for c in span['text'])):
                                raise ValueError('Data do PDF menor que 8 pontos; ajuste localizado insuficiente para legibilidade.')
                        for span in line['spans']:
                            x0, y0, x1, y1 = span['bbox']
                            if x0 < -0.5 or y0 < -0.5 or x1 > page.rect.width + 0.5 or y1 > page.rect.height + 0.5:
                                raise ValueError('Texto ultrapassa os limites da página PDF.')
                            if span['text'].strip() and span['size'] < 5:
                                raise ValueError('PDF contém texto menor que 5 pontos; confira a legibilidade no Excel.')


def export_pdf(workbook, destination, sheet, name, expected_pages=1, *, report=None):
    if sys.platform != 'win32':
        raise RuntimeError('Exportação requer Windows e Microsoft Excel instalado.')
    report = report if report is not None else {}
    report.update(version=1, sheet=sheet if sheet in PROPOSED else 'outra aba', expected_pages=expected_pages)
    destination = Path(destination).resolve()
    directory = destination.parent / 'diagnostico_pdf'
    marker = pending_marker(destination)
    started = False
    event(report, 'preparar_arquivos')
    try:
        if destination.exists() or destination == Path(workbook).resolve():
            raise ValueError('PDF de destino já existe ou coincide com a entrada.')
        destination.parent.mkdir(parents=True, exist_ok=True)
        # A failed or interrupted export can never become printable by existence alone.
        marker.write_text('PDF ainda não aprovado para impressão.', encoding='utf-8')
        started = True
        wb = read_template(workbook)
        try:
            area = print_area(wb[sheet])
            addresses = [cell.coordinate for row in wb[sheet] for cell in row
                         if cell.value is not None and str(cell.value).strip()]
            requested_date = wb[sheet]['E8'].value
            if isinstance(requested_date, (int, float)):
                requested_date = from_excel(requested_date, wb.epoch)
            if isinstance(requested_date, datetime): requested_date = requested_date.date()
            if requested_date is not None and not isinstance(requested_date, date):
                raise ValueError('E8 não contém uma data válida. Revise a cópia Excel antes de exportar.')
            date_text = requested_date.strftime('%d/%m/%Y') if requested_date else None
            date_epoch = wb.epoch
        finally:
            wb.close()
        event(report, 'abrir_excel', print_area=area)
        # Only a disposable copy receives PageSetup, date formatting or break changes.
        with tempfile.TemporaryDirectory(prefix='pcm_pdf_', dir=destination.parent) as folder:
            working = Path(folder) / 'impressao.xlsx'
            shutil.copyfile(workbook, working)
            with excel_session(working, read_only=False) as (_, book):
                ws = book.Worksheets(sheet)
                event(report, 'selecionar_aba')
                select = getattr(ws, 'Select', None)
                if callable(select):
                    select(True)  # Replace grouped-sheet selections on the disposable copy.
                calculate = getattr(ws, 'Calculate', None)
                if callable(calculate):
                    calculate()
                event(report, 'preparar_impressao')
                if ws.Range('C16').Value not in (None, ''):
                    raise ValueError('Campo Tipo preenchido; exportação bloqueada.')
                if date_text:
                    check_com_date(ws.Range('E8'), requested_date, date_epoch,
                                   phase='preparação do PDF', repair=True, exact=True)
                event(report, 'aplicar_page_setup')
                configure_page(ws, area, expected_pages)
                event(report, 'medir_escala')
                scale = check_print_scale(ws, area, expected_pages)
                setup = {}
                for key in ('PaperSize', 'Orientation', 'Zoom', 'FitToPagesWide', 'FitToPagesTall',
                            'LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin'):
                    try:
                        value = getattr(ws.PageSetup, key)
                        if isinstance(value, (int, float, bool)): setup[key] = value
                    except Exception:
                        pass
                try:
                    setup['PrintArea'] = re.findall(r'\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+', str(ws.PageSetup.PrintArea))
                except Exception:
                    setup['PrintArea'] = []
                event(report, 'configurar_paginacao', page_setup=setup, scale=scale)
                protected_breaks = False
                try:
                    ws.ResetAllPageBreaks()
                except Exception:
                    if not ws.ProtectContents:
                        raise ValueError('Não foi possível limpar quebras na cópia de impressão. Confira a configuração do Excel; o modelo foi preservado.')
                    protected_breaks = True
                event(report, 'conferir_textos', protected_breaks=protected_breaks)
                expected_texts = [(address, str(ws.Range(address).Text or '')) for address in addresses]
                for address, text in expected_texts:
                    if not text.strip() or set(text.strip()) == {'#'}:
                        raise ValueError(f'{address}: conteúdo não exibido corretamente pelo Excel; confira largura e formato.')
                event(report, 'exportar_com')
                # Worksheet-scoped export; positional COM arguments retained for compatibility.
                ws.ExportAsFixedFormat(0, str(destination), 0, True, False)
                event(report, 'exportacao_retornou', exists=destination.is_file(),
                      size=destination.stat().st_size if destination.is_file() else 0)
                event(report, 'conferir_arquivo')
                if not destination.is_file() or destination.stat().st_size == 0:
                    raise ValueError('O Excel não criou um PDF com tamanho válido.')
                event(report, 'validar_pdf')
                with pymupdf.open(destination) as produced:
                    report['pages'] = len(produced)
                try:
                    validate_pdf(destination, name, expected_pages, expected_texts, require_a4=True, expected_date=date_text)
                except ValueError as exc:
                    if protected_breaks:
                        raise ValueError(f'O Excel não permitiu limpar quebras na cópia protegida e o ajuste de impressão não produziu um PDF válido: {exc}') from exc
                    raise
                event(report, 'encerrar_excel', validation='aprovado')
        event(report, 'concluido', validation='aprovado')
        marker.unlink(missing_ok=True)
        try:
            save_diagnostic(report, directory)
        except (OSError, ValueError):
            logging.warning('Relatório PDF não gravado; exportação e validação concluídas.')
        return destination
    except Exception as exc:
        stage = report['stage']
        report['failure_stage'] = stage
        report['exceptions'] = exception_codes(exc)
        from .excel_date import date_error_payload
        if payload := date_error_payload(exc):
            report['date_diagnostic'] = payload['content']
        known = {'páginas; esperado': 'page_count', 'página sem texto': 'empty_page',
                 'Nome do material não encontrado': 'missing_item_name', 'ausente ou cortado': 'missing_content',
                 'Data brasileira': 'missing_date', 'A4 vertical': 'paper_size',
                 'limites da página': 'text_bounds', 'pontos': 'font_size',
                 'escala de': 'scale', 'fora da área': 'print_area', 'rejeitado': 'cell_validation'}
        report['failure_code'] = next((code for text, code in known.items() if text in str(exc)), stage)
        created = started and destination.is_file()
        report['pdf_created'] = created
        try:
            report['pdf_size'] = destination.stat().st_size if created else 0
        except OSError:
            report['pdf_size'] = None
        if started and created:
            try:
                directory.mkdir(parents=True, exist_ok=True)
                rejected = directory / ('reprovado_' + uuid4().hex + '.pdf')
                destination.replace(rejected)
                report['rejected_pdf_retained'] = True
            except OSError:
                # Marker remains, so preview/print cannot consume an unapproved PDF.
                report['rejected_pdf_retained'] = False
        event(report, 'falhou', failure_stage=stage)
        path = None
        try:
            path = save_diagnostic(report, directory)
        except (OSError, ValueError):
            pass  # The UI retains the sanitized report in memory for Save/Copy.
        if stage == 'validar_pdf' and created:
            message = 'PDF gerado, mas não aprovado na validação.'
        elif stage == 'encerrar_excel' or getattr(exc, 'excel_cleanup_errors', None):
            message = 'Falha ao encerrar a instância Excel; PDF não liberado para impressão.'
        else:
            message = 'Falha na exportação PDF — etapa ' + stage + '.'
        raise PdfExportError(message + ' ' + str(exc), report, path) from exc
