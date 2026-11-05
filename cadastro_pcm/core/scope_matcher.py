"""Join technical scope only by a unique explicit item reference in one PDF."""
from collections import Counter, defaultdict
import re
from .models import Evidence
from .product_text import normalize, commercial_line, clean_description


def match_scope(items, pages):
    scopes = defaultdict(list)
    for page in pages:
        active = False
        reference, lines = '', []
        def flush():
            text = clean_description('\n'.join(lines))
            if reference and text: scopes[(page.source, reference)].append((page, text))
        for raw in page.text.splitlines():
            line = raw.strip()
            title = re.sub(r'^\d+[.)]\s*', '', normalize(line))
            if re.match(r'^(?:escopo(?: de fornecimento)?|especificacoes tecnicas)\b', title):
                flush(); active = True; reference, lines = '', []
                continue
            if not active: continue
            if re.match(r'^(?:imagem|figura|foto|ilustracao)\s*\d+\s*[:.-]', normalize(line)):
                continue
            if commercial_line(title) or re.match(r'^\d+[.)]\s+[^\d]', line):
                flush(); active = False; reference, lines = '', []
                continue
            match = re.match(r'^(?:Item\s+)?(\d+(?:\.\d+)+)\.?\s+(?:[-–—:]\s*)?(.*)$', line, re.I)
            if match:
                flush(); reference, lines = match[1], [match[2]]
            elif reference and line:
                lines.append(re.sub(r'^(?:[-•]\s*|o\s+(?=[A-ZÁ-Ú]))', '', line))
        flush()
    keys = [(item.evidence['item_ref'].source, item.evidence['item_ref'].excerpt)
            for item in items if 'item_ref' in item.evidence]
    counts = Counter(keys)
    for item in items:
        ref = item.evidence.get('item_ref')
        if ref is None: continue
        key = ref.source, ref.excerpt
        candidates = scopes.get(key, [])
        if not candidates: continue
        if len(candidates) != 1 or counts[key] != 1:
            item.warnings.append('Escopo com referência repetida; associação técnica pendente de revisão manual.')
            continue
        page, text = candidates[0]
        if item.description not in ('', item.name):
            item.warnings.append('Descrição técnica já existente; escopo não substituído automaticamente.')
            continue
        item.description = text
        item.evidence['description'] = Evidence(origin='pdf', source=f'{page.source} — página {page.number}',
                                               excerpt=f'Item {ref.excerpt}\n{text}')
