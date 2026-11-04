"""Surgical OOXML edits retain unrelated workbook components byte-for-byte."""
import hashlib
import json
import posixpath
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from datetime import date
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.datetime import to_excel

MAP = {'requester':'C8', 'request_date':'E8', 'area':'C9', 'name':'A12',
       'description':'A14', 'classification':'A16', 'unit':'E16', 'group':'A18',
       'purpose':'C18', 'sap':'E18', 'unit_price':'A20', 'equipment':'C20',
       'imported':'A22', 'agency':'C22', 'lower_ncm':'E22'}
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def fingerprint(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def inspect_template(path):
    wb = load_workbook(path, data_only=False)
    try:
        return {'sha256': fingerprint(path), 'mapping_preliminary': MAP,
                'sheets': {ws.title: {
                    'protected': ws.protection.sheet,
                    'merged': [str(r) for r in ws.merged_cells.ranges],
                    'print_area': str(ws.print_area),
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
                if str(value) not in allowed_values(wb, ws, validation.formula1):
                    raise ValueError(f'{address}: valor {value!r} não permitido na lista.')
            elif validation.type:
                raise ValueError(f'{address}: validação {validation.type} requer suporte específico; gravação bloqueada.')


def write_excel(template, destination, item, sheet, authorized_sha):
    item.require_ready()
    template, destination = Path(template), Path(destination)
    if template.resolve() == destination.resolve() or destination.exists():
        raise ValueError('Destino existente ou igual ao modelo; gravação recusada.')
    if not authorized_sha or fingerprint(template) != authorized_sha:
        raise ValueError('Inspecione e autorize o mapeamento deste modelo nas Configurações.')
    wb = load_workbook(template, data_only=False)
    try:
        if sheet not in wb.sheetnames:
            raise ValueError('Aba selecionada ausente no modelo.')
        ws = wb[sheet]
        if not ws.print_area:
            raise ValueError('Modelo sem área de impressão; solicite configuração corporativa.')
        if (item.imported == 'Sim') != ('importado' in sheet.lower()):
            raise ValueError('Item Importado? incompatível com a aba selecionada.')
        description = item.description
        if item.manufacturer:
            description += '\nFabricante: ' + item.manufacturer
        if item.manufacturer_code:
            description += '\nCódigo do fabricante: ' + item.manufacturer_code
        if item.ncm:
            description += '\nNCM fiscal informado: ' + item.ncm
        values = {address: getattr(item, field) for field, address in MAP.items()}
        values['A14'] = description
        validate_cells(wb, ws, values)
        epoch = wb.epoch
    finally:
        wb.close()
    with ZipFile(template) as source:
        workbook = ET.fromstring(source.read('xl/workbook.xml'))
        rels = ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))
        relation = next(s for s in workbook.find(f'{{{NS}}}sheets') if s.attrib['name'] == sheet)
        rid = relation.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        target = next(r.attrib['Target'] for r in rels if r.attrib['Id'] == rid)
        part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
        root = ET.fromstring(source.read(part))
        data = root.find(f'{{{NS}}}sheetData')
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
                value = to_excel(value, epoch)
            if isinstance(value, (float, int)) or address == 'A20':
                cell.attrib['t'] = 'n'
                ET.SubElement(cell, f'{{{NS}}}v').text = str(value)
            else:
                cell.attrib['t'] = 'inlineStr'
                text = ET.SubElement(ET.SubElement(cell, f'{{{NS}}}is'), f'{{{NS}}}t')
                text.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                text.text = str(value or '')
            from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
            row[:] = sorted(row, key=lambda c: column_index_from_string(coordinate_from_string(c.attrib['r'])[0]))
        destination.parent.mkdir(parents=True, exist_ok=True)
        with destination.open('xb') as handle, ZipFile(handle, 'w', ZIP_DEFLATED) as output:
            for entry in source.infolist():
                output.writestr(entry, ET.tostring(root, encoding='utf-8', xml_declaration=True) if entry.filename == part else source.read(entry.filename))
    return destination
