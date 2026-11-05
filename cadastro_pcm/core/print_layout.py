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


def configure_page(ws, area):
    ws.PageSetup.PrintArea = area
    ws.PageSetup.PaperSize = 9  # xlPaperA4
    ws.PageSetup.Orientation = 1  # xlPortrait
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    # Do not shrink height blindly. Additional pages fail PDF validation for review.
    ws.PageSetup.FitToPagesTall = False
