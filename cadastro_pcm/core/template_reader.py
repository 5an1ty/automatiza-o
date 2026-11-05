"""Read standard and Excel x14 validations without ever saving the template."""
import posixpath
import warnings
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM = 'http://schemas.microsoft.com/office/excel/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def sheet_parts(archive):
    workbook = ET.fromstring(archive.read('xl/workbook.xml'))
    relations = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    targets = {r.attrib['Id']: r.attrib['Target'] for r in relations}
    return {
        sheet.attrib['name']: (target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target))
        for sheet in workbook.find(f'{{{MAIN}}}sheets')
        for target in [targets[sheet.attrib[f'{{{REL}}}id']]]
    }


def read_template(path):
    # openpyxl omits x14 validations; load them separately into the read-only model.
    # The original package, including all extensions, remains the source for writes.
    with warnings.catch_warnings():
        warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed')
        wb = load_workbook(path, data_only=False)
    try:
        with ZipFile(path) as archive:
            for title, part in sheet_parts(archive).items():
                root = ET.fromstring(archive.read(part))
                for element in root.findall(f'.//{{{X14}}}dataValidation'):
                    attrs = dict(element.attrib)
                    attrs = {key: value for key, value in attrs.items() if key in DataValidation.__attrs__}
                    for key in ('allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage'):
                        if key in attrs: attrs[key] = attrs[key] in ('1', 'true')
                    sqref = element.findtext(f'{{{XM}}}sqref')
                    if not sqref:
                        raise ValueError('Validação estendida sem destinos; solicite modelo autorizado.')
                    validation = DataValidation(**attrs, sqref=sqref)
                    for field in ('formula1', 'formula2'):
                        node = element.find(f'{{{X14}}}{field}')
                        if node is not None:
                            setattr(validation, field, node.findtext(f'{{{XM}}}f'))
                    wb[title].add_data_validation(validation)
        return wb
    except Exception:
        wb.close()
        raise
