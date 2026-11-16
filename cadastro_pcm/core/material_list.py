"""Local, conservative parsing of user-entered material lists; no fiscal inference."""
import re
import unicodedata
from collections import Counter
from decimal import Decimal
from .models import Item, Evidence

SOURCE = 'Lista inserida manualmente'
NCM = r'(?:[0-9]{4}\.[0-9]{2}\.[0-9]{2}|[0-9]{8})'


def clean(text):
    text = text.strip()
    for marker in ('**', '__', '`'):
        if text.startswith(marker) and text.endswith(marker):
            text = text[len(marker):-len(marker)].strip()
    return text


def heading(text):
    return ''.join(c for c in unicodedata.normalize('NFD', clean(text).casefold()) if not unicodedata.combining(c))


def price(text, explicit=False):
    text = clean(text)
    labeled = re.match(r'^(?:pre[cç]o(?: unit[aá]rio)?|valor(?: unit[aá]rio)?)\s*:?\s*', text, re.I)
    if labeled:
        explicit = True; text = text[labeled.end():]
    if text.startswith('R$'):
        explicit = True; text = text[2:].strip()
    if re.fullmatch(r'(?:[0-9]{1,3}(?:\.[0-9]{3})+|[0-9]+),[0-9]{2}', text):
        return Decimal(text.replace('.', '').replace(',', '.'))
    if explicit and re.fullmatch(r'[0-9]+(?:\.[0-9]{2})?', text):
        return Decimal(text)
    return None


def make_item(name, ncm='', amount='', warnings=(), *, explicit_price=True):
    name, ncm, amount = clean(name), clean(ncm), clean(amount)
    original_ncm = ncm
    alerts = list(warnings)
    ncm = re.sub(r'^NCM\s*:?\s*', '', ncm, flags=re.I)
    if ncm and not re.fullmatch(NCM, ncm):
        alerts.append('NCM não reconhecido: revisar o texto de origem.'); ncm = ''
    value = price(amount, explicit_price) if amount else None
    if amount and value is None:
        alerts.append('Preço não reconhecido: revisar o texto de origem.')
    item = Item(name=name, description=name, ncm=ncm, unit_price=value, warnings=alerts)
    for field, excerpt in [('name', name), ('description', name), ('ncm', original_ncm), ('unit_price', amount)]:
        if excerpt:
            item.evidence[field] = Evidence(origin='manual', source=SOURCE, excerpt=excerpt)
    return item


def mark_duplicates(items):
    counts = Counter(' '.join(i.name.split()).casefold() for i in items)
    for item in items:
        message = 'Possível duplicata: confira; ambos os registros serão mantidos.'
        item.warnings = [w for w in item.warnings if w != message]
        if counts[' '.join(item.name.split()).casefold()] > 1:
            item.warnings.append(message)
    return items


def parse_material_list(text):
    items, notices = [], []
    current = None
    columns = None
    def finish():
        nonlocal current
        if current:
            items.append(make_item(**current, explicit_price=False)); current = None
    for number, raw in enumerate(text.splitlines(), 1):
        line = clean(re.sub(r'^\s*(?:[-*+] |[0-9]+[.)] )\s*', '', raw))
        if '\t' in raw:
            line = raw.strip(' \r')  # Retain empty first/last spreadsheet cells.
        if not line or re.fullmatch(r'[\s|:═=*_\-]+', line) or line.startswith('```'):
            continue
        # Tabs/pipes are structural only; commas, dimensions and hyphens are not.
        delimiter = '\t' if '\t' in line else ('|' if '|' in line else None)
        if delimiter:
            finish()
            cells = [clean(c) for c in line.strip('|').split(delimiter)]
            keys = [heading(c) for c in cells]
            names = {'cadastro', 'material', 'nome', 'descricao', 'nome/descricao', 'nome do material'}
            prices = {'preco', 'preco unitario', 'valor', 'valor unitario'}
            if any(k in names for k in keys) and 'ncm' in keys:
                columns = {('name' if k in names else 'ncm' if k == 'ncm' else 'amount'): j
                           for j, k in enumerate(keys) if k in names or k == 'ncm' or k in prices}
                continue
            mapping = columns or {'name': 0, 'ncm': 1, 'amount': 2}
            values = {key: cells[index] if index < len(cells) else '' for key, index in mapping.items()}
            if not values.get('name'):
                notices.append(f'Linha {number}: registro sem descrição; revise a origem.'); continue
            warnings = [] if len(cells) == len(mapping) else ['Colunas ambíguas: confira todos os campos na origem.']
            items.append(make_item(**values, warnings=warnings))
            continue
        if heading(line.lstrip('# ')) in {'lista de materiais', 'materiais', 'cadastro', 'ncm', 'preco unitario'}:
            continue
        code = re.fullmatch(r'(?:NCM\s*:?\s*)?(' + NCM + ')', line, re.I)
        is_ncm = code is not None
        value = price(line)
        if is_ncm or value is not None or re.match(r'^(?:NCM\b|R\$|pre[cç]o\b|valor\b)', line, re.I):
            field = 'ncm' if is_ncm or line.upper().startswith('NCM') else 'amount'
            if current is None:
                notices.append(f'Linha {number}: campo sem material associado; revise a origem.')
            elif current[field]:
                current['warnings'].append(f'Linha {number}: {field} repetido; associação ambígua, revise a origem.')
                current[field] = ''
            else:
                current[field] = code[1] if code else line
            continue
        if current and not current['ncm'] and not current['amount']:
            current['warnings'].append('Descrições consecutivas sem campos: confira se são materiais separados.')
        finish()
        current = dict(name=line, ncm='', amount='', warnings=[])
    finish()
    return mark_duplicates(items), notices
