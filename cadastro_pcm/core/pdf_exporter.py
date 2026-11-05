import sys
from pathlib import Path
import pymupdf
from .template_reader import read_template
from .excel_runtime import excel_session
from .print_layout import print_area, configure_page


def validate_pdf(path, name, expected_pages=1):
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


def export_pdf(workbook, destination, sheet, name, expected_pages=1):
    if sys.platform != 'win32':
        raise RuntimeError('Exportação requer Windows e Microsoft Excel instalado.')
    destination = Path(destination).resolve()
    if destination.exists() or destination == Path(workbook).resolve():
        raise ValueError('PDF de destino já existe ou coincide com a entrada.')
    wb = read_template(workbook)
    try:
        area = print_area(wb[sheet])
    finally:
        wb.close()
    try:
        with excel_session(workbook, read_only=True) as (_, book):
            ws = book.Worksheets(sheet)
            if ws.Range('C16').Value not in (None, ''):
                raise ValueError('Campo Tipo preenchido; exportação bloqueada.')
            configure_page(ws, area)
            ws.ExportAsFixedFormat(0, str(destination), 0, True, False)
            validate_pdf(destination, name, expected_pages)
        return destination
    except Exception:
        destination.unlink(missing_ok=True)
        raise
