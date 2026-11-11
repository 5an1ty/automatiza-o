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
    primary_error = None
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
    except BaseException as exc:
        primary_error = exc
        raise
    finally:
        errors = []
        for operation, cleanup in (
                ('Close', lambda: book.Close(False) if book is not None else None),
                ('Quit', lambda: excel.Quit() if excel is not None else None),
                ('CoUninitialize', pythoncom.CoUninitialize)):
            try:
                cleanup()
            except Exception as exc:
                code = getattr(exc, 'hresult', None)
                errors.append({'operation': operation, 'type': type(exc).__name__,
                               'hresult': f'0x{code & 0xffffffff:08X}' if isinstance(code, int) else None})
        if errors:
            if primary_error is not None:
                primary_error.excel_cleanup_errors = errors
            else:
                error = RuntimeError('Falha na liberação da instância Excel criada pelo aplicativo.')
                error.excel_cleanup_errors = errors
                raise error
