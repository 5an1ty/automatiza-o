"""Content-free PDF pipeline evidence; corporate documents remain local."""
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4


def event(report, stage, **values):
    report['stage'] = stage
    report.setdefault('events', []).append({'time_utc': datetime.now(timezone.utc).isoformat(),
                                            'stage': stage, **values})
    logging.info('PDF %s', json.dumps(report['events'][-1], ensure_ascii=False))


def exception_codes(exc):
    result, visited = [], set()
    while exc is not None and id(exc) not in visited:
        visited.add(id(exc))
        entry = {'type': type(exc).__name__}
        hresult = getattr(exc, 'hresult', None)
        if isinstance(hresult, int): entry['hresult'] = f'0x{hresult & 0xffffffff:08X}'
        info = getattr(exc, 'excepinfo', None)
        if isinstance(info, tuple) and len(info) > 5 and isinstance(info[5], int):
            entry['scode'] = f'0x{info[5] & 0xffffffff:08X}'
        if isinstance(exc, OSError) and exc.errno is not None: entry['errno'] = exc.errno
        if getattr(exc, 'excel_cleanup_errors', None): entry['cleanup'] = exc.excel_cleanup_errors
        result.append(entry)
        exc = exc.__cause__ or exc.__context__
    return result


def save_diagnostic(report, directory):
    directory = Path(directory); directory.mkdir(parents=True, exist_ok=True)
    path = directory / ('diagnostico_pdf_' + uuid4().hex + '.json')
    text = json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False)
    path.write_text(text, encoding='utf-8')
    if json.loads(path.read_text(encoding='utf-8')) != report:
        raise OSError('Não foi possível confirmar o relatório PDF.')
    return path


class PdfExportError(ValueError):
    def __init__(self, message, report, path=None):
        super().__init__(message)
        self.report = report
        self.report_path = path


def diagnostic_payload(exc):
    current, seen = exc, set()
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        if isinstance(current, PdfExportError):
            return {'content': current.report, 'path': str(current.report_path) if current.report_path else None}
        current = current.__cause__ or current.__context__
    from .excel_date import date_error_payload
    return date_error_payload(exc)


def pending_marker(path):
    return Path(str(path) + '.pendente')


def usable_pdf(path):
    return Path(path).is_file() and not pending_marker(path).exists()


def unsafe_cleanup(payload):
    report = payload.get('content', {}) if payload else {}
    return any(failure.get('operation') == 'Quit'
               for exc in report.get('exceptions', []) for failure in exc.get('cleanup', []))
