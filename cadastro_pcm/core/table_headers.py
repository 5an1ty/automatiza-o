"""Semantic column mapping for ruled and word-aligned PDF tables."""
import re
from .product_text import normalize

ALIASES = {
    'name': ['descricao', 'produto', 'material', 'servico', 'descricao do produto'],
    'quantity': ['quantidade', 'qtd', 'qtde', 'qt', 'qt.'],
    'unit_price': ['preco unitario', 'valor unitario', 'vl. unitario', 'vl unit', 'unitario', 'vl. unit.', 'vl unit.', 'preco unit.', 'valor unit.', 'preco unitario (r$)', 'valor unitario (r$)', 'r$ unit', 'r$ unit.', 'r$ unitario'],
    'total_price': ['preco total', 'valor total', 'total', 'r$ total'],
    'ncm': ['ncm'], 'description': ['descricao tecnica'],
    'item_ref': ['item', 'item.', 'n item', 'item nº', 'item n°'],
}


def field_for(label):
    label = normalize(label).rstrip(': .')
    for key, names in ALIASES.items():
        if label in [name.rstrip(': .') for name in names]: return key
    # Only monetary column labels, never IPI/ICMS percentages or an isolated
    # "Impostos" heading. Currency and included-tax suffixes are optional.
    money = re.fullmatch(r'(?:r\$\s*)?(?:(?:valor|preco|vl\.?)\s+)?(unitario|unit\.?|total)(?:\s+(?:c\s*/|com)\s*impostos)?(?:\s*\(r\$\))?', label)
    if money: return 'total_price' if money[1] == 'total' else 'unit_price'
    return None


def data_row(row):
    return any(re.fullmatch(r'\s*(?:R\$\s*)?\d[\d.,\s]*%?\s*', str(cell or '')) for cell in row)


def find_header(table):
    best = None
    for start in range(min(8, len(table))):
        labels = []
        for end in range(start, min(start + 4, len(table))):
            row = table[end]
            if data_row(row): break
            while len(labels) < len(row): labels.append([])
            for col, cell in enumerate(row):
                text = str(cell or '').strip()
                if text and text not in labels[col]: labels[col].append(text)
            combined = [' '.join(parts) for parts in labels]
            headers = [field_for(label) for label in combined]
            recognized = [key for key in headers if key]
            if 'name' not in recognized or len(recognized) < 2 or len(set(recognized)) != len(recognized):
                continue
            score = len(recognized) + 2 * ('unit_price' in recognized) + ('total_price' in recognized)
            score += 0.1 * sum('impostos' in normalize(label) for label in combined)
            candidate = (score, -(end - start), -start)
            if best is None or candidate > best[0]: best = candidate, start, end, headers
    return (best[1], best[2], best[3]) if best else None


def table_confidence(table):
    found = find_header(table)
    if not found: return 0
    start, end, headers = found
    result = sum(key is not None for key in headers)
    if 'unit_price' in headers:
        col = headers.index('unit_price')
        numeric = sum(col < len(row) and bool(re.fullmatch(r'\s*(?:R\$\s*)?\d[\d.\s]*(?:,\d{1,2})?\s*', str(row[col] or ''))) for row in table[end + 1:])
        result += 4 + min(numeric, 3) / 10
    return result
