from dataclasses import dataclass
from pathlib import Path
import pymupdf
from .table_headers import table_confidence
from .word_tables import word_tables


@dataclass
class Page:
    source: str
    number: int
    text: str
    tables: list[list[list[str | None]]]


def read_pdf(path):
    pages = []
    with pymupdf.open(path) as doc:
        if doc.needs_pass:
            raise ValueError(f'{Path(path).name}: PDF protegido por senha.')
        for index, page in enumerate(doc):
            text = page.get_text(sort=True)
            tables = []
            if text.strip():
                candidates = []
                for priority, options in ((1, {}), (0, {'strategy': 'text', 'min_words_vertical': 2, 'min_words_horizontal': 1})):
                    for table in page.find_tables(**options).tables:
                        cells = table.extract()
                        candidates.append((table_confidence(cells), priority, table.bbox, cells))
                for box, cells in word_tables(page):
                    candidates.append((table_confidence(cells), 0, box, cells))
                selected = []
                for score, priority, box, cells in sorted(candidates, key=lambda c: (c[0], c[1]), reverse=True):
                    rect = pymupdf.Rect(box)
                    # Prefer the strongest header/column mapping for overlapping
                    # detections. Never produce duplicate items from both finders.
                    if any((rect & pymupdf.Rect(other)).get_area() > 0.5 * min(rect.get_area(), pymupdf.Rect(other).get_area()) for other, _ in selected):
                        continue
                    selected.append((box, cells))
                tables = [cells for box, cells in sorted(selected, key=lambda pair: (pair[0][1], pair[0][0]))]
            pages.append(Page(str(Path(path).resolve()), index + 1, text, tables))
    if not any(page.text.strip() for page in pages):
        raise ValueError(f'{Path(path).name}: sem texto extraível. OCR necessário; não disponível nesta versão.')
    return pages
