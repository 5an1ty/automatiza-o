"""Conservative positional fallback for commercial tables without ruling lines."""
from bisect import bisect_right
import re
from .product_text import normalize, commercial_line
from .table_headers import field_for, data_row, find_header


def word_tables(page):
    lines = []
    for word in sorted(page.get_text('words'), key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        center = (word[1] + word[3]) / 2
        if lines and abs(center - lines[-1][0]) <= 3:
            lines[-1][1].append(word)
        else:
            lines.append((center, [word]))
    lines = [(y, sorted(words, key=lambda w: w[0])) for y, words in lines]
    result = []
    consumed = -1
    monetary = {'valor', 'preco', 'vl', 'vl.', 'r$'}
    taxes = {'ipi', 'icms', 'iss', 'pis', 'cofins'}
    codes = {'cod', 'cod.', 'codigo', 'referencia'}
    for index, (_, words) in enumerate(lines):
        if index <= consumed or data_row([w[4] for w in words]): continue
        keys = [field_for(w[4]) for w in words]
        fragments = sum(normalize(w[4]) in monetary for w in words)
        if 'name' not in keys or sum(key is not None for key in keys) + fragments < 2: continue
        anchors, labels = [], []
        for word in words:
            text = normalize(word[4]).rstrip(':')
            key = field_for(text)
            # "Valor Unitário" and "R$ Total" are one column, not two.
            modifier = key in ('unit_price', 'total_price') and labels and normalize(labels[-1]) in monetary
            if (key is not None or text in monetary | taxes | codes) and not modifier:
                anchors.append(word[0]); labels.append(word[4])
            elif labels:
                labels[-1] += ' ' + word[4]
        if len(anchors) < 2: continue
        rows = []
        end = index
        previous_y = lines[index][0]
        for next_index in range(index + 1, len(lines)):
            y, following = lines[next_index]
            if y - previous_y > 45: break
            text = ' '.join(w[4] for w in following)
            if commercial_line(re.sub(r'^\d+[.)]\s*', '', text)): break
            row = [''] * len(anchors)
            for word in following:
                col = bisect_right(anchors, word[0] + 3) - 1
                if col < 0: continue
                row[col] = (row[col] + ' ' + word[4]).strip()
            if not rows and not data_row(row):
                # Remaining header lines are aligned to their parent column.
                for col, value in enumerate(row):
                    if value: labels[col] += ' ' + value
            else:
                if not find_header([labels]): break
                name_col = [field_for(label) for label in labels].index('name')
                if rows and row[name_col] and all(not value for col, value in enumerate(row) if col != name_col):
                    rows[-1][name_col] += '\n' + row[name_col]
                else:
                    rows.append(row)
            previous_y, end = y, next_index
        if rows and find_header([labels]):
            # Bound the fallback to this header and its data, not the whole page.
            all_words = [w for _, group in lines[index:end + 1] for w in group]
            box = (min(w[0] for w in all_words), min(w[1] for w in all_words),
                   max(w[2] for w in all_words), max(w[3] for w in all_words))
            result.append((box, [labels] + rows))
            consumed = end
    return result
