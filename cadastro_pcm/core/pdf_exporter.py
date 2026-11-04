import sys
from pathlib import Path
import pymupdf


def export_pdf(workbook, destination, sheet, name, expected_pages=1):
    if sys.platform != 'win32':
        raise RuntimeError('Exportação requer Windows e Microsoft Excel instalado.')
    destination = Path(destination).resolve()
    if destination.exists():
        raise ValueError('PDF de destino já existe.')
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = book = None
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
        book = excel.Workbooks.Open(str(Path(workbook).resolve()), UpdateLinks=0, ReadOnly=True)
        ws = book.Worksheets(sheet)
        if not ws.PageSetup.PrintArea:
            raise ValueError('Aba sem área de impressão.')
        if ws.Range('C16').Value not in (None, ''):
            raise ValueError('Campo Tipo preenchido; exportação bloqueada.')
        ws.ExportAsFixedFormat(0, str(destination), 0, True, False)
        with pymupdf.open(destination) as pdf:
            text = '\n'.join(page.get_text() for page in pdf)
            if len(pdf) != expected_pages or ' '.join(name.split()).casefold() not in ' '.join(text.split()).casefold():
                raise ValueError('PDF não passou na verificação de páginas e nome. Revise a impressão no Excel.')
        return destination
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    finally:
        try:
            if book is not None:
                book.Close(False)
        finally:
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                pythoncom.CoUninitialize()
