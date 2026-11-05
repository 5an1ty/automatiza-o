"""Local-only readiness checks. Reports contain statuses, never workbook content."""
import importlib.util
import json
import platform
import sys
import tempfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from .excel_writer import MAP, fingerprint
from .excel_runtime import excel_session
from .pdf_exporter import export_pdf, validate_pdf
from .print_layout import print_area, configure_page

REQUIRED_SHEETS = ('Criação de Material Nacional', 'Criação de Material Importado', 'Instruções Gerais', 'Planilha auxiliar')


def diagnose(settings):
    checks = []
    def add(name, status, message):
        checks.append({'verificacao': name, 'status': status, 'mensagem': message})
    add('Python', 'ok' if sys.version_info >= (3, 11) else 'erro', platform.python_version())
    windows = sys.platform == 'win32'
    add('Windows', 'ok' if windows else 'indisponível', 'Windows detectado.' if windows else 'Excel COM não pode ser validado neste sistema.')
    pywin = importlib.util.find_spec('win32com') is not None if windows else False
    add('pywin32', 'ok' if pywin else 'indisponível', 'Disponível.' if pywin else 'Requer Windows e instalação pelo iniciar.bat.')
    try:
        output = Path(settings.output)
        output.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='pcm_check_', dir=output):
            pass
        writable = True
        add('Pasta de saída', 'ok', 'Criação e remoção de pasta temporária confirmadas.')
    except OSError:
        writable = False
        add('Pasta de saída', 'erro', 'Não foi possível gravar. Confira permissões e espaço disponível.')
    template_ok = False
    try:
        if not settings.template or not Path(settings.template).is_file():
            raise ValueError('missing')
        wb = load_workbook(settings.template, data_only=False)
        try:
            missing = [name for name in REQUIRED_SHEETS if name not in wb.sheetnames]
            add('Abas corporativas', 'erro' if missing else 'ok', 'Abas ausentes: ' + ', '.join(missing) if missing else 'Todas as abas esperadas foram encontradas.')
            selected = wb[settings.sheet]
            blocked = []
            for address in MAP.values():
                cell = selected[address]
                if isinstance(cell, MergedCell) or cell.data_type == 'f' or (selected.protection.sheet and cell.protection.locked):
                    blocked.append(address)
            add('Campos editáveis', 'erro' if blocked else 'ok', 'Revisar células: ' + ', '.join(blocked) if blocked else 'Destinos sem bloqueios, fórmulas ou âncoras inválidas.')
            empty_type = selected['C16'].value in (None, '')
            add('Campo Tipo', 'ok' if empty_type else 'erro', 'Vazio.' if empty_type else 'Modelo contém Tipo preenchido; use modelo limpo.')
            try:
                area = print_area(selected)
                layout_ok = True
                add('Área de impressão', 'ok', area + ' — ainda exige conferência visual de legibilidade e objetos.')
            except ValueError:
                layout_ok = False
                add('Área de impressão', 'erro', 'Há conteúdo ou mesclagem fora dos limites. Consulte o relatório de inspeção.')
            authorized = bool(settings.authorized_template_sha256 and fingerprint(settings.template) == settings.authorized_template_sha256)
            add('Autorização do modelo', 'ok' if authorized else 'pendente', 'Hash corresponde ao autorizado.' if authorized else 'Inspecione e autorize o modelo atual nas configurações.')
            template_ok = not missing and not blocked and empty_type and layout_ok and authorized
            add('Assinaturas e layout', 'pendente', 'Confirmar manualmente posições vazias, controle de alterações, imagens e ausência de cortes no PDF. Não homologado automaticamente.')
        finally:
            wb.close()
    except Exception:
        # Do not put raw exception strings, paths, names or cell text into reports.
        add('Modelo Excel', 'erro', 'Modelo ausente, inválido ou aba selecionada indisponível. Confira as configurações.')
    if windows and pywin:
        try:
            with excel_session() as (excel, _):
                book = excel.Workbooks.Add()
                try:
                    ws = book.Worksheets(1)
                    ws.Range('A1').Value = 'Diagnóstico PCM'
                    configure_page(ws, 'A1:E33')
                    if writable:
                        with tempfile.TemporaryDirectory(prefix='pcm_check_', dir=settings.output) as folder:
                            target = Path(folder) / 'teste.pdf'
                            ws.ExportAsFixedFormat(0, str(target), 0, True, False)
                            validate_pdf(target, 'Diagnóstico PCM', 1)
                        add('PDF sintético via Excel', 'ok', 'Exportação real de uma página conferida; arquivo temporário removido.')
                finally:
                    book.Close(False)
            add('Microsoft Excel', 'ok', 'Instância isolada iniciada e encerrada.')
        except Exception:
            add('Microsoft Excel / PDF de teste', 'erro', 'Falha no Excel ou na exportação. Verifique ativação, instalação e impressora padrão; nenhum sucesso foi presumido.')
        if template_ok and writable:
            try:
                with tempfile.TemporaryDirectory(prefix='pcm_check_', dir=settings.output) as folder:
                    export_pdf(settings.template, Path(folder) / 'formulario.pdf', settings.sheet, '', settings.expected_pages)
                add('PDF do modelo sem preenchimento', 'ok', 'Exportação real com páginas conferidas; original aberto somente para leitura e PDF temporário removido. Revisão visual continua pendente.')
            except Exception:
                add('PDF do modelo sem preenchimento', 'erro', 'Exportação ou paginação falhou. Revise o modelo e a impressão no Excel.')
        else:
            add('PDF do modelo', 'não executado', 'Corrija modelo, autorização e saída antes de repetir.')
    else:
        add('Excel e exportação real', 'não executado', 'Execute este diagnóstico no Windows com Excel e pywin32.')
    return {'versao_relatorio': 1, 'verificacoes': checks, 'homologacao_visual': 'pendente'}


def save_report(report, path):
    Path(path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')


def summary(report):
    return '\n\n'.join(f"[{check['status'].upper()}] {check['verificacao']}\n{check['mensagem']}" for check in report['verificacoes'])
