"""Isolated Excel lifetime. No shared instance, unprotect, macros or external links."""
from contextlib import contextmanager
import sys


@contextmanager
def excel_session(path=None, read_only=True):
    if sys.platform != 'win32':
        raise RuntimeError('Esta operação requer Windows e Microsoft Excel desktop instalado.')
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('pywin32 indisponível. Execute iniciar.bat para instalar as dependências.') from exc
    pythoncom.CoInitialize()
    excel = book = None
    try:
        try:
            excel = win32com.client.DispatchEx('Excel.Application')
        except Exception as exc:
            raise RuntimeError('Não foi possível iniciar o Excel. Verifique a instalação e ativação do Microsoft Excel.') from exc
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        excel.AskToUpdateLinks = False
        if path is not None:
            from pathlib import Path
            book = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0, ReadOnly=read_only)
        yield excel, book
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
