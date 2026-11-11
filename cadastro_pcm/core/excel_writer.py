"""Surgical OOXML edits retain unrelated workbook components byte-for-byte."""
import sys
import shutil
import hashlib
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from datetime import date
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.datetime import to_excel
from .print_layout import print_area
from .excel_runtime import excel_session
from .template_reader import read_template, sheet_parts
from .date_style import date_style, DATE_FORMAT
from .excel_date import check_stored_date, check_com_date
from .item_extractor import separate_ncm
from .product_text import clean_description

MAP = {'requester':'C8', 'request_date':'E8', 'area':'C9', 'name':'A12',
       'description':'A14', 'classification':'A16', 'unit':'E16', 'group':'A18',
       'purpose':'C18', 'sap':'E18', 'unit_price':'A20', 'equipment':'C20',
       'imported':'A22', 'agency':'C22', 'lower_ncm':'E22'}
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def use_native_excel():
    return sys.platform == 'win32'


def fingerprint(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def layout_report(ws):
    try:
        return {'status': 'conferir visualmente', 'area': print_area(ws)}
    except ValueError as exc:
        return {'status': 'bloqueado', 'reason': str(exc)}


def inspect_template(path):
    wb = read_template(path)
    try:
        return {'sha256': fingerprint(path), 'mapping_preliminary': MAP,
                'sheets': {ws.title: {
                    'protected': ws.protection.sheet,
                    'merged': [str(r) for r in ws.merged_cells.ranges],
                    'print_area': str(ws.print_area),
                    'print_layout_check': layout_report(ws),
                    'orientation': ws.page_setup.orientation,
                    'paper_size': ws.page_setup.paperSize,
                    'fit_width': ws.page_setup.fitToWidth,
                    'fit_height': ws.page_setup.fitToHeight,
                    'margins': str(ws.page_margins),
                    'cells': {c.coordinate: {'value': str(c.value) if c.value is not None else '',
                              'locked': c.protection.locked, 'type': c.data_type}
                              for row in ws for c in row if c.value is not None or c.coordinate in MAP.values()},
                    'validations': [{'range': str(v.sqref), 'type': v.type, 'formula1': v.formula1,
                                     'formula2': v.formula2} for v in ws.data_validations.dataValidation],
                } for ws in wb.worksheets}}
    finally:
        wb.close()


def allowed_values(wb, ws, formula):
    if not formula:
        raise ValueError('Lista de validação sem origem.')
    if formula.startswith('"') and formula.endswith('"'):
        return formula[1:-1].split(',')
    formula = formula.lstrip('=')
    if formula in wb.defined_names:
        destinations = list(wb.defined_names[formula].destinations)
    elif '!' in formula:
        sheet, reference = formula.rsplit('!', 1)
        destinations = [(sheet.strip("'"), reference)]
    else:
        destinations = [(ws.title, formula)]
    result = []
    try:
        for sheet, reference in destinations:
            cells = wb[sheet][reference]
            if not isinstance(cells, tuple):
                cells = ((cells,),)
            for row in cells:
                for cell in row:
                    if cell.data_type == 'f':
                        raise ValueError('Lista calculada exige validação corporativa.')
                    if cell.value is not None:
                        result.append(str(cell.value))
    except (KeyError, TypeError, AttributeError) as exc:
        raise ValueError('Lista de validação não resolvida; solicite modelo autorizado.') from exc
    return result


def validate_cells(wb, ws, values):
    if ws['C16'].value not in (None, ''):
        raise ValueError('Campo Tipo deve estar vazio no modelo autorizado.')
    for address, value in values.items():
        cell = ws[address]
        if isinstance(cell, MergedCell):
            raise ValueError(f'{address}: endereço não é a âncora da célula mesclada.')
        if cell.data_type == 'f':
            raise ValueError(f'{address}: fórmula institucional não pode ser alterada.')
        if ws.protection.sheet and cell.protection.locked:
            raise ValueError(f'{address}: campo bloqueado; solicite modelo autorizado.')
        for validation in ws.data_validations.dataValidation:
            if address not in validation.sqref:
                continue
            if value in ('', None) and validation.allowBlank:
                continue
            if validation.type == 'list':
                if str(value).casefold() not in {v.casefold() for v in allowed_values(wb, ws, validation.formula1)}:
                    raise ValueError(f'{address}: valor {value!r} não permitido na lista.')
            elif validation.type:
                raise ValueError(f'{address}: validação {validation.type} requer suporte específico; gravação bloqueada.')


def check_template(template, item, sheet, authorized_sha):
    """Same preflight for button explanations and writing; never modifies input."""
    item.require_ready()
    template = Path(template)
    if not authorized_sha or fingerprint(template) != authorized_sha:
        raise ValueError('Inspecione e autorize o mapeamento deste modelo nas Configurações.')
    wb = read_template(template)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError('Aba selecionada ausente no modelo.')
        ws = wb[sheet]
        print_area(ws)
        if (item.imported == 'Sim') != ('importado' in sheet.lower()):
            raise ValueError('Item Importado? incompatível com a aba selecionada.')
        description = clean_description(item.description) or item.name
        # Old records used this longer label. Normalize it before extracting the
        # existing code and re-add exactly one fiscal line.
        description = re.sub(r'(?i)NCM\s+fiscal\s+informado\s*:', 'NCM:', description)
        description, codes = separate_ncm(description)
        if codes and any(code != item.ncm for code in codes):
            raise ValueError('NCM da descrição diverge do campo fiscal; revise o item.')
        if item.ncm:
            description += '\nNCM: ' + item.ncm
        values = {address: getattr(item, field) for field, address in MAP.items()}
        values['A14'] = description
        if item.price_fill_later:
            values['A20'] = None
        # An empty optional field already empty in the template needs no edit.
        # Preserve it instead of reassigning a blank through a restrictive list.
        values = {address: value for address, value in values.items()
                  if value not in ('', None) or ws[address].value not in ('', None)}
        validate_cells(wb, ws, values)
        epoch = wb.epoch
    finally:
        wb.close()
    return values, epoch


def write_excel(template, destination, item, sheet, authorized_sha):
    template, destination = Path(template), Path(destination)
    if template.resolve() == destination.resolve() or destination.exists():
        raise ValueError('Destino existente ou igual ao modelo; gravação recusada.')
    values, epoch = check_template(template, item, sheet, authorized_sha)
    if use_native_excel():
        result = write_excel_com(template, destination, values, sheet, epoch)
        try:
            verify_request_date(result, sheet, item.request_date)
        except Exception:
            destination.unlink(missing_ok=True)
            raise
        return result
    with ZipFile(template) as source:
        part = sheet_parts(source)[sheet]
        original_xml = source.read(part).decode('utf-8')
        root = ET.fromstring(original_xml)
        data = root.find(f'{{{NS}}}sheetData')
        styles_xml = None
        for address, value in values.items():
            row_number = ''.join(c for c in address if c.isdigit())
            row = next((r for r in data if r.attrib['r'] == row_number), None)
            if row is None:
                row = ET.SubElement(data, f'{{{NS}}}row', r=row_number)
                data[:] = sorted(data, key=lambda r: int(r.attrib['r']))
            cell = next((c for c in row if c.attrib['r'] == address), None)
            if cell is None:
                cell = ET.SubElement(row, f'{{{NS}}}c', r=address)
            for child in list(cell):
                if child.tag in {f'{{{NS}}}v', f'{{{NS}}}is', f'{{{NS}}}f'}:
                    cell.remove(child)
            if isinstance(value, date):
                if address == 'E8': styles_xml = date_style(source, cell)
                value = to_excel(value, epoch)
            if value is None:
                cell.attrib.pop('t', None)
            elif isinstance(value, (float, int)) or address == 'A20':
                cell.attrib['t'] = 'n'
                ET.SubElement(cell, f'{{{NS}}}v').text = str(value)
            else:
                cell.attrib['t'] = 'inlineStr'
                text = ET.SubElement(ET.SubElement(cell, f'{{{NS}}}is'), f'{{{NS}}}t')
                text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                text.text = str(value or '')
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
            row[:] = sorted(row, key=lambda c: column_index_from_string(coordinate_from_string(c.attrib['r'])[0]))
        # Replace only destination cells. Reserializing the full sheet can drop
        # namespace declarations referenced by mc:Ignorable / Excel extensions.
        # Keep the original root, x14 validations, drawings and all other XML.
        for address in values:
            cell = next(c for row in data for c in row if c.attrib['r'] == address)
            replacement = ET.tostring(cell, encoding='unicode')
            # Lazy attributes are essential: a greedy match can consume '/'
            # from an empty cell and then eat the following nonempty cell.
            pattern = rf'<(?:\w+:)?c\b(?=[^>]*\br=[\"\']{address}[\"\'])[^>]*?(?:/>|>.*?</(?:\w+:)?c>)'
            original_xml, count = re.subn(pattern, lambda _: replacement, original_xml, count=1, flags=re.S)
            if not count:
                row_number = re.search(r'\d+', address).group()
                row_pattern = rf'(<(?:\w+:)?row\b(?=[^>]*\br=[\"\']{row_number}[\"\'])[^>]*?)(/>|>.*?</(?:\w+:)?row>)'
                def append_cell(match):
                    if match[2] == '/>':
                        prefix = re.match(r'<(\w+:)?row', match[1]).group(1) or ''
                        return match[1] + '>' + replacement + f'</{prefix}row>'
                    # Insert in column order, before the first later cell.
                    from openpyxl.utils.cell import column_index_from_string
                    column = column_index_from_string(re.match(r'[A-Z]+', address).group())
                    body = match[2]
                    for later in re.finditer(r'<(?:\w+:)?c\b[^>]*\br=[\"\']([A-Z]+)\d+[\"\']', body):
                        if column_index_from_string(later[1]) > column:
                            return match[1] + body[:later.start()] + replacement + body[later.start():]
                    return match[1] + re.sub(r'(</(?:\w+:)?row>)$', lambda m: replacement + m[1], body)
                original_xml, count = re.subn(row_pattern, append_cell, original_xml, count=1, flags=re.S)
                if not count:
                    opening = re.search(r'<((?:\w+:)?sheetData)\b[^>]*>', original_xml)
                    if not opening:
                        raise ValueError('Estrutura sheetData ausente no modelo.')
                    prefix = opening[1].removesuffix('sheetData')
                    new_row = f'<{prefix}row r="{row_number}">{replacement}</{prefix}row>'
                    later = next((m for m in re.finditer(r'<(?:\w+:)?row\b[^>]*\br=[\"\'](\d+)[\"\']', original_xml)
                                  if int(m[1]) > int(row_number)), None)
                    if opening[0].endswith('/>'):
                        original_xml = original_xml[:opening.start()] + opening[0][:-2] + '>' + new_row + f'</{opening[1]}>' + original_xml[opening.end():]
                    else:
                        position = later.start() if later else original_xml.index(f'</{opening[1]}>')
                        original_xml = original_xml[:position] + new_row + original_xml[position:]
        destination.parent.mkdir(parents=True, exist_ok=True)
        with destination.open('xb') as handle, ZipFile(handle, 'w', ZIP_DEFLATED) as output:
            for entry in source.infolist():
                content = (original_xml.encode('utf-8') if entry.filename == part else
                           styles_xml if entry.filename == 'xl/styles.xml' and styles_xml else source.read(entry.filename))
                output.writestr(entry, content)
    try:
        verify_request_date(destination, sheet, item.request_date)
    except Exception:
        destination.unlink(missing_ok=True)
        raise
    return destination


def verify_request_date(path, sheet, expected):
    wb = read_template(path)
    try:
        cell = wb[sheet]['E8']
        check_stored_date(cell, expected, wb.epoch)
    finally:
        wb.close()


def write_excel_com(template, destination, values, sheet, epoch):
    """Write only on a fresh copy using native Excel. Failed copies are removed."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open('xb') as output, template.open('rb') as source:
        shutil.copyfileobj(source, output)
    try:
        with excel_session(destination, read_only=False) as (_, book):
            if book.ReadOnly:
                raise ValueError('Cópia aberta somente para leitura. Verifique permissões da pasta.')
            ws = book.Worksheets(sheet)
            if ws.Range('C16').Value not in (None, ''):
                raise ValueError('Campo Tipo deve estar vazio.')
            for address in values:
                cell = ws.Range(address)
                if cell.HasFormula or (ws.ProtectContents and cell.Locked):
                    raise ValueError(f'{address}: campo protegido ou fórmula; solicite modelo autorizado.')
                if cell.MergeCells and cell.MergeArea.Cells(1, 1).Address != cell.Address:
                    raise ValueError(f'{address}: endereço não é âncora da mesclagem.')
            # Native validation includes extended validations not exposed by openpyxl.
            validations = set()
            try:
                validated = ws.Cells.SpecialCells(-4174)  # xlCellTypeAllValidation
            except Exception as exc:
                # Error 1004 is expected only when no validated cells exist.
                info = getattr(exc, 'excepinfo', None)
                if not info or (info[5] & 0xffff) != 1004:
                    raise ValueError('Não foi possível verificar as validações nativas do Excel.') from exc
            else:
                for address in values:
                    if ws.Application.Intersect(ws.Range(address), validated) is not None:
                        validations.add(address)
            for address, value in values.items():
                if isinstance(value, date):
                    value = to_excel(value, epoch)
                elif address == 'A20' and value is not None:
                    value = float(value)
                if address == 'E8':
                    ws.Range(address).Value2 = float(value)
                    ws.Range(address).NumberFormat = DATE_FORMAT
                    check_com_date(ws.Range(address), values['E8'], epoch, phase='antes de salvar', repair=True)
                else:
                    ws.Range(address).Value = value
                if address in validations and not ws.Range(address).Validation.Value:
                    raise ValueError(f'{address}: valor rejeitado pela validação do Excel.')
            book.Save()
        if 'E8' in values:
            with excel_session(destination, read_only=True) as (_, reopened):
                check_com_date(reopened.Worksheets(sheet).Range('E8'), values['E8'], epoch, phase='após reabrir')
        return destination
    except Exception:
        destination.unlink(missing_ok=True)
        raise
