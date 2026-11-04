"""Estratégias extensíveis: tabelas com cabeçalhos e blocos rotulados.
Textos não reconhecidos permanecem disponíveis para revisão manual.
"""
import re
import unicodedata
from .models import Item, Evidence, decimal_br


def normalize(value):
    return ''.join(c for c in unicodedata.normalize('NFD', value.lower()) if unicodedata.category(c) != 'Mn').strip()


ALIASES = {
    'name': ['descricao', 'produto', 'material', 'servico', 'descricao do produto'],
    'quantity': ['quantidade', 'qtd', 'qtde', 'qt.'],
    'unit_price': ['preco unitario', 'valor unitario', 'vl. unitario', 'vl unit', 'unitario'],
    'total_price': ['preco total', 'valor total', 'total'],
    'ncm': ['ncm'], 'manufacturer': ['fabricante', 'marca'],
    'manufacturer_code': ['codigo do fabricante', 'codigo', 'referencia'],
    'description': ['descricao tecnica'],
}


def field_for(label):
    label = normalize(label).rstrip(': .')
    return next((key for key, names in ALIASES.items() if label in names), None)


def make_item(values, page, excerpt):
    clean = {}
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
    if not clean.get('name'):
        return None
    clean.setdefault('description', clean['name'])
    item = Item(**clean)
    item.evidence = {key: Evidence(origin='pdf', source=f'{page.source} — página {page.number}', excerpt=excerpt) for key in clean}
    return item


class TableStrategy:
    def extract(self, page):
        result = []
        for table in page.tables:
            if not table:
                continue
            headers = [field_for(str(cell or '')) for cell in table[0]]
            if 'name' not in headers:
                continue
            for row in table[1:]:
                values = {key: row[i] for i, key in enumerate(headers) if key and i < len(row)}
                item = make_item(values, page, str(row))
                if item:
                    result.append(item)
        return result


class LabelStrategy:
    def extract(self, page):
        result, values, lines = [], {}, []
        for line in page.text.splitlines():
            match = re.match(r'^\s*([^:]{2,45}):\s*(.*?)\s*$', line)
            key = field_for(match[1]) if match else None
            if not key:
                continue
            if key == 'name' and 'name' in values:
                item = make_item(values, page, '\n'.join(lines))
                if item:
                    result.append(item)
                values, lines = {}, []
            values[key] = match[2]
            lines.append(line)
        item = make_item(values, page, '\n'.join(lines))
        if item:
            result.append(item)
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
    return result
