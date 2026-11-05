"""Explicit local Windows integration check. Reports only E8, never other cells."""
import argparse
from datetime import date
import json
from pathlib import Path
import shutil
import sys
import tempfile
from uuid import uuid4
from openpyxl.utils.datetime import to_excel
from .date_style import DATE_FORMAT
from .excel_date import inspect_com_date, check_com_date, DateValidationError
from .excel_runtime import excel_session
from .excel_writer import fingerprint, validate_cells, verify_request_date
from .pdf_exporter import export_pdf
from .template_reader import read_template


def diagnose_date(template, sheet, expected, output_dir, with_pdf=False):
    if sys.platform != 'win32': raise RuntimeError('Este teste de integração requer Windows e Microsoft Excel real.')
    template = Path(template).resolve()
    before = fingerprint(template)
    wb = read_template(template)
    try:
        validate_cells(wb, wb[sheet], {'E8': expected})
        epoch = wb.epoch
    finally:
        wb.close()
    folder = Path(output_dir); folder.mkdir(parents=True, exist_ok=True)
    target = folder / f'data_solicitacao_{uuid4().hex[:10]}.xlsx'
    with target.open('xb') as dst, template.open('rb') as src: shutil.copyfileobj(src, dst)
    report = {'expected_date': expected.isoformat(), 'ok': False,
              'pdf': 'não executado' if with_pdf else 'não solicitado'}
    phase = 'gravação'
    try:
        with excel_session(target, read_only=False) as (_, book):
            if book.ReadOnly: raise ValueError('Cópia aberta somente para leitura.')
            cell = book.Worksheets(sheet).Range('E8')
            if cell.HasFormula or (book.Worksheets(sheet).ProtectContents and cell.Locked):
                raise ValueError('E8 protegida ou com fórmula; original e proteção preservados.')
            cell.Value2 = float(to_excel(expected, epoch))
            cell.NumberFormat = DATE_FORMAT
            report['before_save'] = check_com_date(cell, expected, epoch, phase='diagnóstico antes de salvar',
                                                   report_dir=folder, repair=True)
            book.Save()
        phase = 'reabertura'
        with excel_session(target, read_only=True) as (_, book):
            report['after_reopen'] = inspect_com_date(book.Worksheets(sheet).Range('E8'), expected, epoch)
            check_com_date(book.Worksheets(sheet).Range('E8'), expected, epoch,
                           phase='diagnóstico após reabrir', report_dir=folder)
        verify_request_date(target, sheet, expected)
        if with_pdf:
            phase = 'exportação PDF'
            report['pdf'] = 'falhou'
            export_pdf(target, target.with_suffix('.pdf'), sheet, '', 1)
            report['pdf'] = 'exportado e data brasileira conferida'
        report['ok'] = True
    except Exception as exc:
        report['failed_phase'] = phase
        report['error'] = exc.reason if isinstance(exc, DateValidationError) else 'Operação local não concluída: ' + type(exc).__name__
        if isinstance(exc, DateValidationError): report['technical_report'] = exc.report_path.name
    report['source_unchanged'] = fingerprint(template) == before
    report['copy'] = target.name
    report_path = target.with_suffix('.json')
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    return report, report_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template', required=True)
    parser.add_argument('--sheet', default='Criação de Material Nacional')
    parser.add_argument('--date', type=date.fromisoformat, default=date.today())
    parser.add_argument('--output-dir')
    parser.add_argument('--export-pdf', action='store_true')
    args = parser.parse_args()
    try:
        folder = args.output_dir or tempfile.mkdtemp(prefix='pcm_data_')
        report, path = diagnose_date(args.template, args.sheet, args.date, folder, args.export_pdf)
        print(('PASSOU' if report['ok'] and report['source_unchanged'] else 'FALHOU') + f' — relatório local: {path}')
        print('A cópia é de diagnóstico, não um cadastro aprovado. Não publique os arquivos corporativos.')
        return 0 if report['ok'] and report['source_unchanged'] else 1
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == '__main__': raise SystemExit(main())
