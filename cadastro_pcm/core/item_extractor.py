"""Estratégias extensíveis: tabelas com cabeçalhos e blocos rotulados.
Textos não reconhecidos permanecem disponíveis para revisão manual.
"""
import re
import unicodedata
from .models import Item, Evidence, decimal_br


def normalize(value):
    return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', value.lower()) if unicodedata.category(c) != 'Mn')).strip()


ALIASES = {
    'name': ['descricao', 'produto', 'material', 'servico', 'descricao do produto'],
    'quantity': ['quantidade', 'qtd', 'qtde', 'qt', 'qt.'],
    'unit_price': ['preco unitario', 'valor unitario', 'vl. unitario', 'vl unit', 'unitario', 'vl. unit.', 'vl unit.', 'preco unit.', 'valor unit.', 'preco unitario (r$)', 'valor unitario (r$)'],
    'total_price': ['preco total', 'valor total', 'total'],
    'ncm': ['ncm'], 'manufacturer': ['fabricante', 'marca'],
    'manufacturer_code': ['codigo do fabricante', 'codigo', 'referencia'],
    'description': ['descricao tecnica'],
}


def field_for(label):
    label = normalize(label).rstrip(': .')
    return next((key for key, names in ALIASES.items() if label in names), None)


def make_item(values, page, excerpt):
    clean, warnings = {}, []
    for key, value in values.items():
        if not value or not str(value).strip():
            continue
        value = str(value).strip()
        try:
            clean[key] = decimal_br(value) if key in ('quantity', 'unit_price', 'total_price') else value
            # Validate fields individually so one malformed NCM does not discard the item.
            Item(**{key: clean[key]})
        except (ValueError, ArithmeticError):
            clean.pop(key, None)
            warnings.append(f'{key}: valor não reconhecido; confira o documento original.')
    if not clean.get('name'):
        return None
    clean.setdefault('description', clean['name'])
    item = Item(**clean, warnings=warnings)
    item.evidence = {key: Evidence(origin='pdf', source=f'{page.source} — página {page.number}', excerpt=excerpt) for key in clean}
    return item


class TableStrategy:
    def extract(self, page):
        result = []
        for table in page.tables:
            header_index = None
            for index, row in enumerate(table[:5]):
                candidates = [field_for(str(cell or '')) for cell in row]
                if 'name' in candidates and sum(key is not None for key in candidates) >= 2:
                    headers, header_index = candidates, index
                    break
            if header_index is None:
                continue
            pending = None
            excerpt = []
            for row in table[header_index + 1:]:
                values = {key: row[i] for i, key in enumerate(headers) if key and i < len(row)}
                name = str(values.get('name') or '').strip()
                if normalize(name).startswith(('total', 'subtotal')):
                    if pending:
                        item = make_item(pending, page, str(excerpt))
                        if item: result.append(item)
                        pending = None
                    continue
                # Only append pure description continuation rows. Never copy numbers between rows.
                other = [v for k, v in values.items() if k not in ('name', 'description') and str(v or '').strip()]
                if not name and pending and values.get('description') and not other:
                    pending['description'] = str(pending.get('description') or pending['name']) + '\n' + str(values['description'])
                    excerpt.append(row)
                    continue
                if pending:
                    item = make_item(pending, page, str(excerpt))
                    if item: result.append(item)
                pending = values if name else None
                excerpt = [row]
            if pending:
                item = make_item(pending, page, str(excerpt))
                if item: result.append(item)
        return result


class LabelStrategy:
    def extract(self, page):
        result, values, lines = [], {}, []
        continuation = None
        # Normalize horizontal whitespace, preserving line and item boundaries.
        for raw in page.text.splitlines():
            line = re.sub(r'[^\S\n]+', ' ', raw).strip()
            match = re.match(r'^([^:]{2,45}):\s*(.*?)\s*$', line)
            key = field_for(match[1]) if match else None
            if not key:
                if continuation and line and ':' not in line and not re.match(r'(?i)^(total|subtotal|frete|impostos|condicoes)\b', normalize(line)):
                    values[continuation] += '\n' + line
                    lines.append(line)
                else:
                    continuation = None
                continue
            if key == 'name' and 'name' in values:
                item = make_item(values, page, '\n'.join(lines))
                if item: result.append(item)
                values, lines = {}, []
            if key != 'name' and 'name' not in values:
                continue  # Never attribute document-level totals to the next product.
            if key in values and values[key] != match[2]:
                values[key] = ''  # Conflicting labels remain pending.
            else:
                values[key] = match[2]
            continuation = key if key in ('name', 'description') else None
            lines.append(line)
        item = make_item(values, page, '\n'.join(lines))
        if item: result.append(item)
        return result


STRATEGIES = [TableStrategy(), LabelStrategy()]


def extract_items(pages):
    result = []
    for page in pages:
        for strategy in STRATEGIES:
            items = strategy.extract(page)
            if items:
                result.extend(items)
                break
    # Flag, never delete or merge, matching product identities across documents/pages.
    seen = {}
    for item in result:
        key = (normalize(item.name), normalize(item.description), item.manufacturer_code)
        if key in seen:
            warning = 'Possível duplicata: confira documento, página, embalagem e quantidade.'
            item.warnings.append(warning)
            if warning not in seen[key].warnings: seen[key].warnings.append(warning)
        else:
            seen[key] = item
    return result
