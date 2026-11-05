"""Validate proposed limits; never guess that content beyond them is dispensable."""
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.print_settings import PrintArea

PROPOSED = {'Criação de Material Nacional': 'A1:E33', 'Criação de Material Importado': 'A1:E34'}


def print_area(ws):
    if ws.print_area:
        # One contiguous form only; never silently export multiple areas.
        areas = list(PrintArea.from_string(ws.print_area).ranges)
        if len(areas) != 1:
            raise ValueError('Área de impressão deve ser um único intervalo contínuo.')
        area = str(areas[0])
    else:
        area = PROPOSED.get(ws.title)
        if not area:
            raise ValueError('Aba sem intervalo de impressão conhecido; configure um modelo autorizado.')
    left, top, right, bottom = range_boundaries(area)
    if left != 1 or top != 1 or bottom < 33 or right < 5:
        # Existing smaller synthetic/test or non-corporate forms require explicit correction.
        if not ws.print_area:
            raise ValueError('Intervalo preliminar não abrange o formulário.')
    for row in ws:
        for cell in row:
            if cell.value is not None and not (left <= cell.column <= right and top <= cell.row <= bottom):
                raise ValueError(f'Conteúdo em {cell.coordinate} fora da área de impressão. Revise o modelo antes de exportar.')
    for merged in ws.merged_cells.ranges:
        if merged.min_col < left or merged.max_col > right or merged.min_row < top or merged.max_row > bottom:
            raise ValueError('Célula mesclada ultrapassa a área de impressão; revise os limites.')
    return area


def configure_page(ws, area, expected_pages=1):
    ws.PageSetup.PrintArea = area
    ws.PageSetup.PaperSize = 9  # xlPaperA4
    ws.PageSetup.Orientation = 1  # xlPortrait
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = expected_pages


def check_print_scale(ws, area, expected_pages=1):
    """Use Excel's actual point dimensions before fitting; do not shrink blindly.

    Keep existing margins and reject more than 15% reduction. This is a geometry
    guard, not a certificate of visual readability or merged-cell text fit.
    """
    if expected_pages < 1: raise ValueError('Número de páginas inválido.')
    setup = ws.PageSetup
    width = 595.28 - float(setup.LeftMargin) - float(setup.RightMargin)
    height = 841.89 - float(setup.TopMargin) - float(setup.BottomMargin)
    region = ws.Range(area)
    if min(width, height, float(region.Width), float(region.Height)) <= 0:
        raise ValueError('Dimensões ou margens de impressão inválidas.')
    scale = min(1.0, width / float(region.Width), height * expected_pages / float(region.Height))
    if scale < 0.85:
        raise ValueError(f'O formulário exigiria escala de {scale:.0%}; inferior ao limite de 85% para preservar legibilidade. Revise alturas, largura e margens da cópia no Excel.')
    return scale
