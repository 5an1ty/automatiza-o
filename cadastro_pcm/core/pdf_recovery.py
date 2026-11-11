"""Recovery from an app-generated workbook; recheck each manually edited version."""
import tempfile
from pathlib import Path
from datetime import date
from openpyxl import Workbook
from .models import Item
from .template_reader import read_template
from .print_layout import print_area, PROPOSED
from .excel_writer import fingerprint
from .pdf_exporter import export_pdf
from .pdf_diagnostics import diagnostic_payload


def existing_workbook(path):
    path = Path(path)
    record = Item.model_validate_json((path.parent / 'registro.json').read_text(encoding='utf-8'))
    if not record.approved:
        raise ValueError('O registro local não contém aprovação humana.')
    if path.name != 'cadastro.xlsx':
        raise ValueError('Selecione cadastro.xlsx junto do registro.json criado pelo aplicativo.')
    book = read_template(path)
    try:
        sheets = {}
        for name in book.sheetnames:
            if name not in PROPOSED: continue
            sheet = book[name]
            try:
                print_area(sheet)
                if not sheet['A12'].value or sheet['C16'].value not in (None, ''):
                    continue
                if sheet['E8'].value is None:
                    continue
                sheets[name] = str(sheet['A12'].value)
            except ValueError:
                continue
        if not sheets:
            raise ValueError('Nenhuma aba de formulário válido encontrada. Confira identificação, Tipo, data e área de impressão.')
        return {'sheets': sheets, 'sha256': fingerprint(path)}
    finally:
        book.close()


def export_existing(path, destination, sheet, expected_sha, expected_pages=1):
    info = existing_workbook(path)
    if info['sha256'] != expected_sha:
        raise ValueError('O Excel mudou desde a conferência. Selecione e confira a nova versão.')
    if sheet not in info['sheets']:
        raise ValueError('Aba não validada para exportação.')
    return export_pdf(path, destination, sheet, info['sheets'][sheet], expected_pages)


def synthetic_export_test():
    """Only synthetic material, cleaned on both success and failure; report returned."""
    report = {}
    temporary = tempfile.TemporaryDirectory(prefix='pcm_teste_pdf_')
    result = {'ok': False, 'report': report}
    try:
        folder = Path(temporary.name); source = folder / 'sintetico.xlsx'
        book = Workbook(); sheet = book.active; sheet.title = 'Criação de Material Nacional'
        sheet['A1'] = 'Diagnostico PCM'; sheet['A12'] = 'Material sintetico PCM'
        sheet['E8'] = date.today(); sheet['E8'].number_format = 'dd/mm/yyyy'
        sheet['A33'] = 'Controle sintetico'; sheet.print_area = 'A1:E33'
        book.save(source); book.close()
        try:
            export_pdf(source, folder / 'teste.pdf', sheet.title, 'Material sintetico PCM', report=report)
            result = {'ok': True, 'report': report}
        except Exception as exc:
            payload = diagnostic_payload(exc)
            if not payload:
                from .pdf_diagnostics import exception_codes
                report.update(stage='indisponivel', exceptions=exception_codes(exc))
            result = {'ok': False, 'report': payload['content'] if payload else report}
    finally:
        try:
            temporary.cleanup()
        except OSError as exc:
            from .pdf_diagnostics import exception_codes
            result['ok'] = False
            result['report']['temporary_cleanup_error'] = exception_codes(exc)
    return result
