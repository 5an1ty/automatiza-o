"""Excel date semantics independent of COM variants and localized display strings."""
from datetime import date, datetime
from math import isfinite
from pathlib import Path
import json
import re
import tempfile
import unicodedata
from uuid import uuid4
from openpyxl.utils.datetime import from_excel, CALENDAR_MAC_1904

DATE_FORMAT = 'dd/mm/yyyy'
PDF_DATE_FORMAT = '[$-416]dd"/"mm"/"yyyy'


def format_body(number_format):
    value = str(number_format or '').split(';')[0].lower().strip()
    value = re.sub(r'\[\$-[^\]]+\]', '', value)
    return value.replace('"/"', '/').replace('\\/', '/')


def brazilian_date_format(number_format):
    return format_body(number_format) in ('dd/mm/yyyy', 'dd/mm/aaaa')


def date_order(number_format):
    value = format_body(number_format)
    value = re.sub(r'"([./-])"', r'\1', value).replace('\\.', '.').replace('\\-', '-')
    token = r'(d{1,2}|m{1,2}|y{4}|a{4}|j{4}|j{1,2}|t{1,2})'
    match = re.fullmatch(token + r'\s*([./-])\s*' + token + r'\s*\2\s*' + token, value)
    if not match: return None
    def field(part):
        if part[0] in 'ay' or part == 'jjjj': return 'y'
        return 'd' if part[0] in 'jt' else part[0]
    order = ''.join(field(part) for part in (match[1], match[3], match[4]))
    return order if set(order) == {'d', 'm', 'y'} else None


def calendar_date(value, epoch):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, (int, float)) and not isinstance(value, bool) and isfinite(value):
        converted = from_excel(value, epoch)
        if isinstance(converted, datetime): return converted.date()
    raise ValueError('E8: valor não é uma data real nem um número de série Excel válido.')


def validate_date(value, number_format, expected, epoch):
    if calendar_date(value, epoch) != expected:
        raise ValueError('E8: data armazenada diverge da data solicitada.')
    if not brazilian_date_format(number_format):
        raise ValueError('E8: formato brasileiro DD/MM/AAAA não foi preservado.')


def normalized_text(text):
    # Direction marks and Unicode digits are presentation, not different dates.
    return ''.join(str(unicodedata.decimal(c)) if c.isdecimal() else c
                   for c in str(text) if unicodedata.category(c) != 'Cf').strip()


def presentation_date(text, order):
    match = re.fullmatch(r'(\d{1,4})\s*([./-])\s*(\d{1,4})\s*\2\s*(\d{1,4})', normalized_text(text))
    if not match or order is None: return None
    parts = dict(zip(order, (match[1], match[3], match[4])))
    if len(parts['y']) != 4 or len(parts['d']) > 2 or len(parts['m']) > 2: return None
    try: return date(int(parts['y']), int(parts['m']), int(parts['d']))
    except ValueError: return None


def safe_value(value):
    if isinstance(value, (date, datetime)): return value.isoformat()
    if isinstance(value, (int, float)) and not isinstance(value, bool) and isfinite(value): return value
    if value is None: return None
    return '<valor não datável omitido>'


def safe_format(value):
    def locale(match):
        code = match[1]
        return match[0] if re.fullmatch(r'(?:[0-9a-fA-F]{3,8}|[a-zA-Z]{2}(?:-[a-zA-Z]{2})?)', code) else '[$-omitido]'
    sections = []
    for section in value.split(';'):
        if date_order(section) or section.lower().strip() in ('general', 'geral', '0', '0.00', '@', ''):
            sections.append(re.sub(r'\[\$-([^\]]+)\]', locale, section))
        else:
            sections.append('<formato não reconhecido omitido>')
    return ';'.join(sections)


def inspect_com_date(cell, expected, epoch):
    unavailable = []
    def read(name):
        try: return getattr(cell, name)
        except Exception:
            unavailable.append(name)
            return None
    raw_value, serial = read('Value'), read('Value2')
    number_format = str(read('NumberFormat') or '')
    local_format = str(read('NumberFormatLocal') or '')
    text = str(read('Text') or '').strip()
    width, column_width = read('Width'), read('ColumnWidth')
    def matches(value):
        try: return calendar_date(value, epoch) == expected
        except (ValueError, OverflowError): return False
    local_order = date_order(local_format)
    order = local_order or date_order(number_format)
    order_source = 'NumberFormatLocal' if local_order else 'NumberFormat'
    parsed = presentation_date(text, order)
    if text == expected.strftime('%d/%m/%Y'):
        display, cause = 'ok', 'DD/MM/AAAA exato'
    elif parsed == expected:
        display, cause = 'regional confirmada', 'Data conferida pela ordem declarada em ' + order_source + ': ' + order.upper()
    elif text and set(text) == {'#'}:
        display, cause = 'largura insuficiente', 'Excel exibiu apenas #; valor e formato são avaliados separadamente'
    else:
        display, cause = 'divergente', 'Texto não corresponde à data esperada na ordem declarada; não foi inferida outra ordem'
    safe_text = text if (parsed is not None or re.fullmatch(r'[\d.,]+|#+', text) or not text) else '<texto não datável omitido>'
    return {
        'value': safe_value(raw_value), 'value_ok': matches(raw_value) if 'Value' not in unavailable else None,
        'value2': safe_value(serial), 'expected_date': expected.isoformat(),
        'date_system': '1904' if epoch == CALENDAR_MAC_1904 else '1900',
        'date_ok': matches(serial), 'number_format': safe_format(number_format),
        'number_format_local': safe_format(local_format),
        'local_date_order': local_order, 'display_date_order': order, 'display_order_source': order_source,
        'format_ok': brazilian_date_format(number_format), 'text': safe_text,
        'cell_width_points': safe_value(width), 'column_width_chars': safe_value(column_width),
        'width_ok': False if display == 'largura insuficiente' else
                    (width > 0 if display in ('ok', 'regional confirmada') and isinstance(width, (int, float)) else None),
        'display': display, 'display_cause': cause,
        'display_ok': display in ('ok', 'regional confirmada'),
        'properties_unavailable': unavailable,
    }


def require_com_date(report, require_display=False, exact=False):
    if not report['date_ok']:
        raise ValueError('E8: data armazenada via Value2 diverge da data solicitada.')
    if not report['format_ok']:
        raise ValueError('E8: formato brasileiro não preservado pelo Excel.')
    if report['display'] == 'divergente':
        raise ValueError('E8: exibição da data não foi confirmada pelo valor e formato local declarados.')
    if require_display and report['display'] == 'largura insuficiente':
        raise ValueError('E8: largura insuficiente para uma data legível; a estrutura do formulário foi preservada.')
    if exact and report['display'] != 'ok':
        raise ValueError('E8: apresentação DD/MM/AAAA com barras ainda não confirmada para o PDF.')


def date_report_content(attempts, phase):
    return {'phase': phase, 'cell': 'E8', 'attempts': attempts,
            'date_ok': attempts[-1]['date_ok'], 'format_ok': attempts[-1]['format_ok'],
            'display_ok': attempts[-1]['display_ok']}


def save_date_report(attempts, phase, report_dir=None):
    if report_dir is None:
        from .settings import settings_path
        report_dir = settings_path().parent / 'diagnostico-data'
    content = date_report_content(attempts, phase)
    def write(folder):
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / f'diagnostico_data_{uuid4().hex}.json'
        try:
            with path.open('x', encoding='utf-8') as output:
                json.dump(content, output, ensure_ascii=False, indent=2, allow_nan=False)
            if json.loads(path.read_text(encoding='utf-8')) != content:
                raise OSError('Gravação do relatório não confirmada.')
        except (OSError, ValueError):
            try: path.unlink(missing_ok=True)
            except OSError: pass
            raise
        return path
    try: return write(Path(report_dir))
    except (OSError, ValueError): return write(Path(tempfile.mkdtemp(prefix='pcm_data_')))


class DateValidationError(ValueError):
    def __init__(self, reason, path, diagnostic=None):
        self.reason, self.report_path = reason, path
        self.diagnostic = diagnostic
        location = (f' Relatório técnico local: {path}' if path else
                    ' Não foi possível gravar o relatório automático. Use Salvar diagnóstico ou Copiar diagnóstico na janela de erro.')
        super().__init__(reason + location)


def date_error_payload(error):
    """Keep sanitized E8 observations across worker exception wrappers."""
    seen = set()
    while error is not None and id(error) not in seen:
        seen.add(id(error))
        if isinstance(error, DateValidationError) and error.diagnostic is not None:
            return {'content': error.diagnostic,
                    'path': str(error.report_path) if error.report_path else None}
        error = error.__cause__ or error.__context__
    return None


def check_com_date(cell, expected, epoch, *, phase, report_dir=None, repair=False, exact=False):
    """Check and optionally repair presentation only on the caller's work copy."""
    attempts = [inspect_com_date(cell, expected, epoch)]
    adjustment = ''
    if repair and attempts[0]['date_ok'] and attempts[0]['format_ok']:
        try:
            if exact or attempts[-1]['display'] == 'largura insuficiente':
                cell.NumberFormat = PDF_DATE_FORMAT
                adjustment = 'Localidade pt-BR e barras literais, somente em E8'
                attempts.append(inspect_com_date(cell, expected, epoch))
            if attempts[-1]['display'] == 'largura insuficiente':
                # No AutoFit/column resizing: preserve the form and pagination.
                # Bound shrinkage conservatively; the PDF also checks date font.
                import pymupdf
                width = attempts[-1]['cell_width_points']
                minimum = pymupdf.get_text_length(expected.strftime('%d/%m/%Y'), fontsize=8) * 1.15 + 4
                if isinstance(width, (int, float)) and width >= minimum and not cell.MergeCells:
                    cell.ShrinkToFit = True
                    adjustment += '; ShrinkToFit limitado à célula, sem alterar dimensões'
                    attempts.append(inspect_com_date(cell, expected, epoch))
        except Exception:
            adjustment += '; ajuste recusado pelo Excel; proteção preservada'
    attempts[-1]['adjustment'] = adjustment
    # Persist before raising. Unexpected text/values are redacted, never copied
    # from corporate cells into diagnostic files. Reports contain only E8.
    needs_report = len(attempts) > 1 or any(not r['date_ok'] or not r['format_ok'] or r['display'] != 'ok' for r in attempts)
    path = None
    if needs_report:
        try: path = save_date_report(attempts, phase, report_dir)
        except (OSError, ValueError): pass  # The error dialog retains the report in memory.
    try: require_com_date(attempts[-1], require_display=True, exact=exact)
    except ValueError as exc:
        raise DateValidationError(str(exc), path, date_report_content(attempts, phase)) from exc
    return attempts[-1]


def check_stored_date(cell, expected, epoch, report_dir=None):
    """Independent OOXML readback; never invent COM/display observations."""
    try:
        validate_date(cell.value, cell.number_format, expected, epoch)
    except ValueError as exc:
        try: date_ok = calendar_date(cell.value, epoch) == expected
        except (ValueError, OverflowError): date_ok = False
        report = {'value': safe_value(cell.value), 'value2': None, 'text': None,
                  'number_format': safe_format(cell.number_format), 'number_format_local': None,
                  'cell_width_points': None, 'expected_date': expected.isoformat(),
                  'date_ok': date_ok, 'format_ok': brazilian_date_format(cell.number_format),
                  'display_ok': None, 'observation': 'Arquivo OOXML reaberto; propriedades COM não observadas'}
        phase = 'verificação independente do arquivo salvo'
        try: path = save_date_report([report], phase, report_dir)
        except (OSError, ValueError): path = None
        raise DateValidationError(str(exc), path, date_report_content([report], phase)) from exc
