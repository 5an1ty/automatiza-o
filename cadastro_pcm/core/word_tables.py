"""Conservative positional fallback for commercial tables without ruling lines."""
from bisect import bisect_right
import re
from .product_text import normalize, commercial_line
from .table_headers import field_for, data_row, find_header


MONETARY = {'valor', 'preco', 'vl', 'vl.', 'r$'}
TAXES = {'ipi', 'icms', 'iss', 'pis', 'cofins'}
CODES = {'cod', 'cod.', 'codigo', 'referencia'}


def header_columns(band):
    runs, loose = [], []
    for y, words in band:
        current = None
        for word in words:
            text = normalize(word[4]).rstrip(':')
            key = field_for(text)
            modifier = key in ('unit_price', 'total_price') and current and normalize(' '.join(w[4] for w in current[1])) in MONETARY
            if (key is not None or text in MONETARY | TAXES | CODES) and not modifier:
                current = [y, [word]]; runs.append(current)
            elif current:
                current[1].append(word)
            else:
                loose.append(word)
    columns = []
    for y, words in runs:
        left, right = min(w[0] for w in words), max(w[2] for w in words)
        # Vertical header fragments overlap horizontally with their parent.
        matches = [col for col in columns if abs(col[0] - y) > 3
                   and min(col[2], right) - max(col[1], left) > 0.5 * min(col[2] - col[1], right - left)]
        if len(matches) > 1: return [], []
        if matches:
            col = matches[0]; col[1] = min(col[1], left); col[2] = max(col[2], right); col[3].extend(words)
        else:
            columns.append([y, left, right, list(words)])
    columns.sort(key=lambda col: (col[1] + col[2]) / 2)
    centers = [(col[1] + col[2]) / 2 for col in columns]
    for word in loose:
        if not centers: break
        center = (word[0] + word[2]) / 2
        nearest = min(range(len(centers)), key=lambda col: abs(centers[col] - center))
        columns[nearest][3].append(word)
    labels = [' '.join(w[4] for w in sorted(col[3], key=lambda w: ((w[1] + w[3]) / 2, w[0]))) for col in columns]
    return centers, labels


def word_tables(page):
    lines = []
    for word in sorted(page.get_text('words'), key=lambda w: ((w[1] + w[3]) / 2, w[0])):
        center = (word[1] + word[3]) / 2
        if lines and abs(center - lines[-1][0]) <= 3:
            lines[-1][1].append(word)
        else:
            lines.append((center, [word]))
    lines = [(y, sorted(words, key=lambda w: w[0])) for y, words in lines]
    result, consumed = [], -1
    for index, (y, words) in enumerate(lines):
        if index <= consumed or data_row([w[4] for w in words]): continue
        keys = [field_for(w[4]) for w in words]
        fragments = sum(normalize(w[4]) in MONETARY for w in words)
        if 'name' not in keys or sum(key is not None for key in keys) + fragments < 2: continue
        start = index
        # A vertically centered Description may sit below monetary/tax headings.
        while start > 0 and y - lines[start - 1][0] <= 15:
            preceding = lines[start - 1][1]
            if data_row([w[4] for w in preceding]) or commercial_line(' '.join(w[4] for w in preceding)): break
            start -= 1
        header_end = index
        while header_end + 1 < min(index + 5, len(lines)):
            next_y, following = lines[header_end + 1]
            if next_y - lines[header_end][0] > 45 or data_row([w[4] for w in following]): break
            if commercial_line(' '.join(w[4] for w in following)): break
            header_end += 1
        centers, labels = header_columns(lines[start:header_end + 1])
        found = find_header([labels])
        if len(centers) < 2 or not found: continue
        boundaries = [(a + b) / 2 for a, b in zip(centers, centers[1:])]
        name_col = found[2].index('name')
        rows, end = [], header_end
        previous_y = lines[header_end][0]
        for next_index in range(header_end + 1, len(lines)):
            next_y, following = lines[next_index]
            if next_y - previous_y > 45: break
            text = ' '.join(w[4] for w in following)
            if commercial_line(re.sub(r'^\d+[.)]\s*', '', text)): break
            row = [''] * len(centers)
            for word in following:
                col = bisect_right(boundaries, (word[0] + word[2]) / 2)
                row[col] = (row[col] + ' ' + word[4]).strip()
            if rows and row[name_col] and all(not value for col, value in enumerate(row) if col != name_col):
                rows[-1][name_col] += '\n' + row[name_col]
            else:
                rows.append(row)
            previous_y, end = next_y, next_index
        if rows:
            all_words = [w for _, group in lines[start:end + 1] for w in group]
            box = (min(w[0] for w in all_words), min(w[1] for w in all_words),
                   max(w[2] for w in all_words), max(w[3] for w in all_words))
            result.append((box, [labels] + rows)); consumed = end
    return result
