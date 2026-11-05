"""Estratégias extensíveis: tabelas com cabeçalhos e blocos rotulados.
Textos não reconhecidos permanecem disponíveis para revisão manual.
"""
import re
from .models import Item, Evidence, decimal_br
from .product_text import normalize, commercial_line, clean_description
from .table_headers import field_for, find_header
from .scope_matcher import match_scope


def separate_ncm(text):
    """Only item descriptions: explicit fiscal labels or a separated leading code.

    Never search document-level numbers, manufacturer codes or telephone fields.
    """
    codes = []
    def remove(match):
        codes.append(Item(ncm=match['code']).ncm)
        return ''
    text = re.sub(r'\(?\bNCM\s*:?\s*(?P<code>\d{4}\.\d{2}\.\d{2}|\d{8})(?![\w.])\)?', remove, text, flags=re.I)
    if not codes:
        text = re.sub(r'^\s*(?P<code>\d{4}\.\d{2}\.\d{2}|\d{8})\s*[-–—]\s*(?=\D)', remove, text)
    text = re.sub(r'^[\s\-–—:;]+|[\s\-–—:;]+$', '', text)
    return text.strip(), codes


def make_item(values, page, excerpt):
    clean, warnings = {}, []
    values = dict(values)
    reference = str(values.pop('item_ref', '') or '').strip().rstrip('.')
    name = str(values.get('name') or '').strip()
    if not name or commercial_line(name.splitlines()[0]):
        return None
    for key in ('name', 'description'):
        if values.get(key):
            original = str(values[key])
            values[key] = clean_description(original)
            if values[key] != original.strip():
                warnings.append('Trechos comerciais ou campos secundários foram omitidos; confira a evidência original.')
    codes = []
    for key in ('name', 'description'):
        if values.get(key):
            values[key], found = separate_ncm(str(values[key]))
            codes.extend(found)
    if values.get('ncm'):
        try:
            codes.append(Item(ncm=str(values['ncm']).strip()).ncm)
        except ValueError:
            warnings.append('NCM: valor não reconhecido; confira o documento original.')
            codes.append('')
    if codes:
        if len(set(codes)) == 1 and codes[0]:
            values['ncm'] = codes[0]
        else:
            values['ncm'] = ''
            warnings.append('NCM divergente na linha do item; confira o código fiscal no documento.')
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
    if item.unit_price is None:
        item.warnings.append('Preço unitário não identificado com confiança. Total, IPI e quantidade não substituem o unitário; revise o PDF.')
    if item.quantity is not None and item.unit_price is not None and item.total_price is not None:
        if abs(item.quantity * item.unit_price - item.total_price) > decimal_br('0,01'):
            item.warnings.append('Quantidade × unitário difere do total; confira impostos e descontos. Valores extraídos foram preservados.')
    item.evidence = {key: Evidence(origin='pdf', source=f'{page.source} — página {page.number}', excerpt=excerpt) for key in clean}
    if reference and re.fullmatch(r'\d+(?:\.\d+)*', reference):
        item.evidence['item_ref'] = Evidence(origin='pdf', source=page.source, excerpt=reference)
    return item


class TableStrategy:
    def extract(self, page):
        result = []
        for table in page.tables:
            found = find_header(table)
            if found is None: continue
            header_start, header_index, headers = found
            heading = str(table[header_start:header_index + 1])
            pending = None
            excerpt = []
            for row in table[header_index + 1:]:
                # Repeated column headings are not product rows.
                repeated = [field_for(str(cell or '')) for cell in row]
                if 'name' in repeated and sum(key is not None for key in repeated) >= 2:
                    if pending:
                        item = make_item(pending, page, heading + str(excerpt))
                        if item: result.append(item)
                    pending, excerpt = None, []
                    continue
                values = {key: row[i] for i, key in enumerate(headers) if key and i < len(row)}
                name = str(values.get('name') or '').strip()
                if commercial_line(name) or normalize(name) in ('total', 'subtotal'):
                    if pending:
                        item = make_item(pending, page, heading + str(excerpt))
                        if item: result.append(item)
                        pending = None
                    # Commercial sections terminate item association, including
                    # prices that happen to line up with the product columns.
                    break
                # Only append pure description continuation rows. Never copy numbers between rows.
                other = [v for k, v in values.items() if k not in ('name', 'description') and str(v or '').strip()]
                if not name and pending and values.get('description') and not other:
                    pending['description'] = str(pending.get('description') or pending['name']) + '\n' + str(values['description'])
                    excerpt.append(row)
                    continue
                if pending:
                    item = make_item(pending, page, heading + str(excerpt))
                    if item: result.append(item)
                pending = values if name else None
                excerpt = [row]
            if pending:
                item = make_item(pending, page, heading + str(excerpt))
                if item: result.append(item)
        return result


class LabelStrategy:
    def extract(self, page):
        result, values, lines = [], {}, []
        continuation = None
        commercial_section = False
        # Normalize horizontal whitespace, preserving line and item boundaries.
        for raw in page.text.splitlines():
            line = re.sub(r'[^\S\n]+', ' ', raw).strip()
            if re.fullmatch(r'(?:\d+[.)\s-]*)?(?:itens(?: da cotacao| / servicos)?|materiais|produtos|servicos|relacao de itens)', normalize(line)):
                commercial_section = False
                continuation = None
                continue
            match = re.match(r'^([^:]{2,45}):\s*(.*?)\s*$', line)
            key = field_for(match[1]) if match else None
            if (not key and commercial_line(line)) or (key == 'name' and commercial_line(match[2])):
                item = make_item(values, page, '\n'.join(lines))
                if item: result.append(item)
                values, lines, continuation = {}, [], None
                commercial_section = True
                continue
            if commercial_section:
                if key == 'name' and normalize(match[1]) in ('produto', 'material', 'servico'):
                    commercial_section = False
                else:
                    continue
            if not key:
                if continuation and line and ':' not in line and not re.match(r'(?i)^(total|subtotal|frete|impostos|condicoes)\b', normalize(line)):
                    # Continuation belongs to technical detail, not the short name.
                    values['description'] = values.get('description', values.get('name', '')) + '\n' + line
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
    pages = list(pages)
    result = []
    for page in pages:
        for strategy in STRATEGIES:
            items = strategy.extract(page)
            if items:
                result.extend(items)
                break
    match_scope(result, pages)
    # Flag, never delete or merge, matching product identities across documents/pages.
    seen = {}
    for item in result:
        key = (normalize(item.name), normalize(item.description))
        if key in seen:
            warning = 'Possível duplicata: confira documento, página, embalagem e quantidade.'
            item.warnings.append(warning)
            if warning not in seen[key].warnings: seen[key].warnings.append(warning)
        else:
            seen[key] = item
    return result
