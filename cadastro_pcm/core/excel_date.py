"""Excel date semantics independent of COM variants and localized display strings."""
from datetime import date, datetime
from math import isfinite
import re
from openpyxl.utils.datetime import from_excel, CALENDAR_MAC_1904

DATE_FORMAT = 'dd/mm/yyyy'


def brazilian_date_format(number_format):
    # Excel may persist locale prefixes, quoted/escaped separators or a text
    # section. Do not accept different field ordering or a two-digit year.
    value = str(number_format or '').split(';')[0].lower().strip()
    value = re.sub(r'\[\$-[^\]]+\]', '', value)
    value = value.replace('"/"', '/').replace('\\/', '/')
    return value in ('dd/mm/yyyy', 'dd/mm/aaaa')


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


def inspect_com_date(cell, expected, epoch):
    value = cell.Value2
    number_format = str(cell.NumberFormat or '')
    local_format = str(cell.NumberFormatLocal or '')
    text = str(cell.Text or '').strip()
    try:
        correct_value = calendar_date(value, epoch) == expected
    except (ValueError, OverflowError):
        correct_value = False
    return {
        'value2': value.isoformat() if isinstance(value, (date, datetime)) else value,
        'expected_date': expected.isoformat(),
        'date_system': '1904' if epoch == CALENDAR_MAC_1904 else '1900',
        'date_ok': correct_value,
        'number_format': number_format,
        'number_format_local': local_format,
        'format_ok': brazilian_date_format(number_format),
        'text': text,
        'display': ('ok' if text == expected.strftime('%d/%m/%Y') else
                    'largura insuficiente' if text and set(text) == {'#'} else 'divergente'),
    }


def require_com_date(report):
    if not report['date_ok']:
        raise ValueError('E8: data armazenada via Value2 diverge da data solicitada.')
    if not report['format_ok']:
        raise ValueError('E8: formato brasileiro não preservado pelo Excel: ' + report['number_format'])
    if report['display'] == 'divergente':
        raise ValueError('E8: exibição da data diverge de DD/MM/AAAA após reabrir. Confira formato e largura no Excel.')
    # #### denotes a width problem, not an incorrect date. PDF validation still
    # requires the actual formatted date to be present in the exported document.
