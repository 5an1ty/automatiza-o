import sys
import shutil
import tempfile
import logging
from datetime import date, datetime
from pathlib import Path
import pymupdf
from openpyxl.utils.datetime import from_excel
from .template_reader import read_template
from .excel_runtime import excel_session
from .print_layout import print_area, configure_page, check_print_scale
from .excel_date import check_com_date


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


def export_pdf(workbook, destination, sheet, name, expected_pages=1):
    if sys.platform != 'win32':
        raise RuntimeError('Exportação requer Windows e Microsoft Excel instalado.')
    destination = Path(destination).resolve()
    if destination.exists() or destination == Path(workbook).resolve():
        raise ValueError('PDF de destino já existe ou coincide com a entrada.')
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
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        # Layout changes and page break removal are limited to a disposable copy.
        # Neither the official template nor the saved filled workbook is changed.
        with tempfile.TemporaryDirectory(prefix='pcm_pdf_', dir=destination.parent) as folder:
            working = Path(folder) / 'impressao.xlsx'
            shutil.copyfile(workbook, working)
            with excel_session(working, read_only=False) as (_, book):
                ws = book.Worksheets(sheet)
                if ws.Range('C16').Value not in (None, ''):
                    raise ValueError('Campo Tipo preenchido; exportação bloqueada.')
                if date_text:
                    check_com_date(ws.Range('E8'), requested_date, date_epoch,
                                   phase='preparação do PDF', repair=True, exact=True)
                configure_page(ws, area, expected_pages)
                check_print_scale(ws, area, expected_pages)
                protected_breaks = False
                try:
                    ws.ResetAllPageBreaks()
                except Exception as exc:
                    if not ws.ProtectContents:
                        raise ValueError('Não foi possível limpar quebras na cópia de impressão. Confira a configuração do Excel; o modelo foi preservado.') from exc
                    # Excel's Fit To pagination ignores manual breaks. A protected
                    # sheet can deny deleting them; never unprotect it to print.
                    # Continue with Fit To and require the complete PDF checks.
                    protected_breaks = True
                    logging.info('Quebras mantidas na cópia protegida; paginação será conferida no PDF com FitToPages.')
                expected_texts = [(address, str(ws.Range(address).Text or '')) for address in addresses]
                for address, text in expected_texts:
                    if not text.strip() or set(text.strip()) == {'#'}:
                        raise ValueError(f'{address}: conteúdo não exibido corretamente pelo Excel; confira largura e formato.')
                ws.ExportAsFixedFormat(0, str(destination), 0, True, False)
                try:
                    validate_pdf(destination, name, expected_pages, expected_texts, require_a4=True, expected_date=date_text)
                except ValueError as exc:
                    if protected_breaks:
                        raise ValueError(f'O Excel não permitiu limpar quebras na cópia protegida e o ajuste de impressão não produziu um PDF válido: {exc}') from exc
                    raise
        return destination
    except Exception:
        destination.unlink(missing_ok=True)
        raise
