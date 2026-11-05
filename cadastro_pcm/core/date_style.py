"""Append a date style while preserving all existing workbook styles."""
from copy import deepcopy
import re
from xml.etree import ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DATE_FORMAT = 'dd/mm/yyyy'


def append_element(xml, tag, element, count):
    fragment = ET.tostring(element, encoding='unicode')
    pattern = rf'(<((?:\w+:)?{tag})\b[^>]*?)(/>|>.*?</(?:\w+:)?{tag}>)'
    def replace(match):
        opening = re.sub(r'\bcount=["\']\d+["\']', f'count="{count}"', match[1])
        if not re.search(r'\bcount=', opening): opening += f' count="{count}"'
        if match[3] == '/>': return opening + '>' + fragment + f'</{match[2]}>'
        return opening + re.sub(r'(</(?:\w+:)?' + tag + '>)$', lambda m: fragment + m[1], match[3])
    updated, found = re.subn(pattern, replace, xml, count=1, flags=re.S)
    if found: return updated
    if tag != 'numFmts': raise ValueError('Modelo sem estilos de célula válidos.')
    opening = re.search(r'<((?:\w+:)?styleSheet)\b[^>]*>', xml)
    if not opening: raise ValueError('Modelo sem estrutura de estilos válida.')
    prefix = opening[1].removesuffix('styleSheet')
    return xml[:opening.end()] + f'<{prefix}numFmts count="1">{fragment}</{prefix}numFmts>' + xml[opening.end():]


def date_style(archive, cell):
    xml = archive.read('xl/styles.xml').decode('utf-8')
    root = ET.fromstring(xml)
    formats = root.find(f'{{{NS}}}numFmts')
    formats = list(formats) if formats is not None else []
    number_id = next((int(f.attrib['numFmtId']) for f in formats if f.attrib.get('formatCode') == DATE_FORMAT), None)
    if number_id is None:
        number_id = max([163] + [int(f.attrib['numFmtId']) for f in formats]) + 1
        xml = append_element(xml, 'numFmts', ET.Element(f'{{{NS}}}numFmt', numFmtId=str(number_id), formatCode=DATE_FORMAT), len(formats) + 1)
    styles = root.find(f'{{{NS}}}cellXfs')
    if styles is None: raise ValueError('Modelo sem estilos de célula válidos.')
    style = deepcopy(styles[int(cell.attrib.get('s', '0'))])
    style.set('numFmtId', str(number_id)); style.set('applyNumberFormat', '1')
    for index, existing in enumerate(styles):
        if ET.tostring(existing) == ET.tostring(style):
            cell.set('s', str(index))
            return xml.encode('utf-8')
    cell.set('s', str(len(styles)))
    return append_element(xml, 'cellXfs', style, len(styles) + 1).encode('utf-8')
