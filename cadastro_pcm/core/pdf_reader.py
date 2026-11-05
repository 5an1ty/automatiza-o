from dataclasses import dataclass
from pathlib import Path
import pymupdf


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
            tables = [table.extract() for table in page.find_tables().tables] if text.strip() else []
            if text.strip() and not tables:
                # Word-aligned table strategy handles columns without ruling lines.
                tables = [table.extract() for table in page.find_tables(strategy='text', min_words_vertical=2, min_words_horizontal=1).tables]
            pages.append(Page(str(Path(path).resolve()), index + 1, text, tables))
    if not any(page.text.strip() for page in pages):
        raise ValueError(f'{Path(path).name}: sem texto extraível. OCR necessário; não disponível nesta versão.')
    return pages
