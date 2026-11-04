import json
import logging
import os
from datetime import date
from pathlib import Path
from PySide6.QtCore import QThread, Signal, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QTableWidget, QTableWidgetItem, QMessageBox,
    QFileDialog, QDialog, QFormLayout, QLineEdit, QTextEdit, QCheckBox,
    QDialogButtonBox, QLabel, QComboBox)
from cadastro_pcm.core.models import Item, Evidence, decimal_br, money_br
from cadastro_pcm.core.settings import load_settings, save_settings
from cadastro_pcm.core.pdf_reader import read_pdf
from cadastro_pcm.core.item_extractor import extract_items
from cadastro_pcm.core.excel_writer import inspect_template, write_excel
from cadastro_pcm.core.pdf_exporter import export_pdf
from cadastro_pcm.core.file_manager import new_output
from cadastro_pcm.core.research import STATUS, apply_suggestion


class Worker(QThread):
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, function):
        super().__init__()
        self.function = function

    def run(self):
        try:
            self.done.emit(self.function())
        except Exception as exc:
            logging.exception('Falha na operação')
            self.failed.emit(str(exc))


LABELS = {'name':'Nome', 'description':'Descrição técnica', 'quantity':'Quantidade',
    'unit_price':'Preço unitário', 'total_price':'Preço total (referência)', 'ncm':'NCM fiscal',
    'manufacturer':'Fabricante', 'manufacturer_code':'Código fabricante',
    'requester':'Requisitante', 'area':'Área', 'request_date':'Data (AAAA-MM-DD)',
    'classification':'Classificação', 'unit':'Grupo UM', 'group':'Grupo de Itens',
    'purpose':'Finalidade', 'sap':'Código SAP', 'equipment':'Equip. Utilização',
    'imported':'Item Importado?', 'agency':'Órgão fiscalizador', 'lower_ncm':'NCM inferior'}


class Editor(QDialog):
    def __init__(self, item, parent):
        super().__init__(parent)
        self.setWindowTitle('Revisar item e evidências')
        self.resize(800, 720)
        from PySide6.QtWidgets import QScrollArea
        layout = QVBoxLayout(self)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        body = QWidget(); form = QFormLayout(body); self.inputs = {}
        for key, label in LABELS.items():
            value = getattr(item, key)
            control = QTextEdit() if key == 'description' else QLineEdit()
            text = '' if value is None else str(value)
            if key == 'description':
                control.setPlainText(text); control.setMaximumHeight(90)
            else:
                control.setText(text)
            evidence = item.evidence.get(key)
            origin = evidence.origin if evidence else ('pendente' if not text else 'regra')
            control.setToolTip(evidence.model_dump_json(indent=2) if evidence else origin)
            form.addRow(f'{label} [{origin}]', control)
            self.inputs[key] = control
        scroll.setWidget(body); layout.addWidget(scroll)
        self.confirm = QCheckBox('Revisei todos os dados, fontes e estimativas; aprovo este item.')
        layout.addWidget(self.confirm)
        layout.addWidget(QLabel('Tipo e assinaturas permanecem vazios. NCM ausente continua pendente.'))
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.save(item)); buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def save(self, item):
        try:
            data = item.model_dump()
            for key, control in self.inputs.items():
                value = control.toPlainText() if isinstance(control, QTextEdit) else control.text().strip()
                if key in ('quantity', 'unit_price', 'total_price'):
                    value = decimal_br(value) if value else None
                if str(value) != str(getattr(item, key)):
                    data['evidence'][key] = Evidence(origin='manual', source='Revisão do usuário').model_dump()
                data[key] = value
            data['approved'] = self.confirm.isChecked()
            self.result_item = Item.model_validate(data)
            self.accept()
        except Exception as exc:
            QMessageBox.warning(self, 'Dados inválidos', str(exc))


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = load_settings(); self.items = []; self.outputs = {}; self.worker = None
        self.setWindowTitle('Cadastro PCM — revisão de materiais e serviços')
        self.resize(1250, 780); self.setAcceptDrops(True)
        body = QWidget(); self.setCentralWidget(body); layout = QVBoxLayout(body)
        layout.addWidget(QLabel('Importe os PDFs, revise cada item e gere os documentos usando o modelo oficial.'))
        self.documents = QListWidget(); self.documents.setMaximumHeight(110); layout.addWidget(self.documents)
        actions = QHBoxLayout(); layout.addLayout(actions)
        for title, handler in [('Adicionar PDFs', self.add_files), ('Remover PDF', self.remove_file), ('Ver texto do PDF', self.view_source),
                               ('Processar', self.process), ('Adicionar item manual', self.add_item),
                               ('Configurações', self.configure)]:
            button = QPushButton(title); button.clicked.connect(handler); actions.addWidget(button)
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(['Descrição', 'Quantidade', 'Preço unitário', 'NCM', 'Fabricante', 'Origem', 'Validação'])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.cellDoubleClicked.connect(lambda *_: self.edit())
        layout.addWidget(self.table)
        actions = QHBoxLayout(); layout.addLayout(actions)
        for title, handler in [('Editar item', self.edit), ('Registrar pesquisa', self.research),
                               ('Gerar Excel', lambda: self.generate(False)), ('Gerar Excel + PDF', lambda: self.generate(True)),
                               ('Visualizar PDF', self.preview), ('Imprimir…', self.print_pdf), ('Abrir saída', self.open_output)]:
            button = QPushButton(title); button.clicked.connect(handler); actions.addWidget(button)
        self.status = QLabel('Pronto. ' + STATUS); self.status.setWordWrap(True); layout.addWidget(self.status)
        if not self.settings.requester:
            from PySide6.QtCore import QTimer
            QTimer.singleShot(0, self.configure)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()

    def dropEvent(self, event):
        self.add_paths([url.toLocalFile() for url in event.mimeData().urls()])

    def add_paths(self, paths):
        existing = {self.documents.item(i).text() for i in range(self.documents.count())}
        for path in paths:
            path = str(Path(path).resolve())
            if Path(path).suffix.lower() == '.pdf' and path not in existing:
                self.documents.addItem(path); existing.add(path)

    def add_files(self):
        self.add_paths(QFileDialog.getOpenFileNames(self, 'Selecionar PDFs', '', 'PDF (*.pdf)')[0])

    def view_source(self):
        entry = self.documents.currentItem()
        if entry is None:
            return
        def done(pages):
            dialog = QDialog(self); dialog.setWindowTitle('Texto extraído — revisão manual'); dialog.resize(850, 700)
            layout = QVBoxLayout(dialog); text = QTextEdit(); text.setReadOnly(True)
            text.setPlainText('\n\n'.join(f'Página {p.number}\n{p.text}' for p in pages))
            layout.addWidget(text); dialog.exec()
        self.run_job(lambda: read_pdf(entry.text()), done)

    def remove_file(self):
        self.documents.takeItem(self.documents.currentRow())

    def run_job(self, function, callback):
        if self.worker is not None:
            QMessageBox.information(self, 'Aguarde', 'Uma operação está em andamento.'); return
        self.centralWidget().setEnabled(False)
        self.status.setText('Processando…')
        self.worker = Worker(function)
        self.worker.done.connect(callback)
        self.worker.failed.connect(lambda error: QMessageBox.warning(self, 'Operação interrompida', error))
        self.worker.finished.connect(self.job_finished)
        self.worker.start()

    def job_finished(self):
        self.worker.deleteLater(); self.worker = None
        self.centralWidget().setEnabled(True)
        self.status.setText('Operação finalizada. Confira os resultados e mensagens.')

    def process(self):
        paths = [self.documents.item(i).text() for i in range(self.documents.count())]
        if not paths: return
        if self.items and QMessageBox.question(self, 'Substituir lista?', 'Substituir os itens atuais? Alterações não geradas serão descartadas.') != QMessageBox.Yes: return
        def task():
            items, warnings = [], []
            for path in paths:
                try:
                    pages = read_pdf(path); found = extract_items(pages)
                    if not found:
                        warnings.append(f'{Path(path).name}: layout não reconhecido. Consulte o texto extraído e adicione itens manualmente.')
                    if any(not page.text.strip() for page in pages):
                        warnings.append(f'{Path(path).name}: há páginas sem texto; revise as páginas digitalizadas.')
                    items.extend(found)
                except Exception as exc:
                    warnings.append(str(exc))
            return items, warnings
        def done(result):
            self.items, warnings = result; self.outputs = {}
            for item in self.items: self.apply_defaults(item)
            self.refresh()
            if warnings: QMessageBox.warning(self, 'Revisão necessária', '\n'.join(warnings))
        self.run_job(task, done)

    def apply_defaults(self, item):
        for key in ('requester', 'area', 'classification', 'unit', 'group', 'purpose'):
            if key not in item.evidence:
                setattr(item, key, getattr(self.settings, key))
                item.evidence[key] = Evidence(origin='regra', source='Configurações locais')
        item.request_date = date.today() if self.settings.date_policy == 'atual' else self.settings.fixed_date

    def refresh(self):
        self.table.setRowCount(len(self.items))
        for row, item in enumerate(self.items):
            values = [item.name, str(item.quantity or ''), money_br(item.unit_price), item.ncm or 'Pendente', item.manufacturer,
                      '; '.join(sorted({e.source for e in item.evidence.values() if e.source})), 'Aprovado' if item.approved else 'Revisão pendente']
            for col, value in enumerate(values): self.table.setItem(row, col, QTableWidgetItem(value))
        self.table.resizeColumnsToContents(); self.table.setColumnWidth(5, 250)

    def selected(self):
        row = self.table.currentRow()
        if row < 0: raise ValueError('Selecione um item.')
        return row, self.items[row]

    def add_item(self):
        item = Item(); self.apply_defaults(item)
        editor = Editor(item, self)
        if editor.exec(): self.items.append(editor.result_item); self.refresh()

    def edit(self):
        try:
            row, item = self.selected(); editor = Editor(item, self)
            if editor.exec():
                self.items[row] = editor.result_item; self.outputs.pop(row, None); self.refresh()
        except ValueError as exc: QMessageBox.warning(self, 'Seleção', str(exc))

    def research(self):
        try: row, item = self.selected()
        except ValueError as exc: QMessageBox.warning(self, 'Seleção', str(exc)); return
        dialog = QDialog(self); dialog.setWindowTitle('Registrar fonte pesquisada manualmente')
        form = QFormLayout(dialog); form.addRow(QLabel(STATUS))
        field = QComboBox()
        for key in ('ncm', 'unit_price', 'manufacturer', 'manufacturer_code', 'description'): field.addItem(LABELS[key], key)
        value = QLineEdit(); url = QLineEdit(); estimated = QCheckBox('Valor estimado — exige aprovação')
        form.addRow('Campo', field); form.addRow('Novo valor', value); form.addRow('URL da fonte', url); form.addRow(estimated)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel); form.addRow(buttons)
        buttons.accepted.connect(dialog.accept); buttons.rejected.connect(dialog.reject)
        if dialog.exec():
            try:
                key = field.currentData(); val = decimal_br(value.text()) if key == 'unit_price' else value.text()
                if QMessageBox.question(self, 'Confirmar substituição', f'Substituir {LABELS[key]} e registrar a fonte?') != QMessageBox.Yes: return
                self.items[row] = apply_suggestion(item, key, val, url.text(), estimated.isChecked()); self.outputs.pop(row, None); self.refresh()
            except Exception as exc: QMessageBox.warning(self, 'Pesquisa', str(exc))

    def configure(self):
        dialog = QDialog(self); dialog.setWindowTitle('Configurações'); form = QFormLayout(dialog); controls = {}
        names = {'requester':'Requisitante padrão', 'area':'Área', 'date_policy':'Data: atual ou fixa', 'fixed_date':'Data fixa AAAA-MM-DD',
                 'output':'Pasta de saída', 'template':'Caminho do modelo Excel', 'sheet':'Aba do formulário',
                 'classification':'Classificação', 'unit':'Grupo UM', 'group':'Grupo de Itens', 'purpose':'Finalidade', 'expected_pages':'Páginas esperadas'}
        for key, label in names.items():
            control = QLineEdit(str(getattr(self.settings, key))); controls[key] = control; form.addRow(label, control)
        pick = QPushButton('Selecionar modelo…'); form.addRow(pick)
        pick.clicked.connect(lambda: controls['template'].setText(QFileDialog.getOpenFileName(dialog, 'Modelo oficial', '', 'Excel (*.xlsx)')[0] or controls['template'].text()))
        inspect = QPushButton('Inspecionar modelo e salvar relatório…'); form.addRow(inspect)
        def inspect_now():
            try:
                report = inspect_template(controls['template'].text())
                path = QFileDialog.getSaveFileName(dialog, 'Salvar inspeção', 'inspecao_modelo.json', 'JSON (*.json)')[0]
                if path:
                    Path(path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
                    QMessageBox.information(dialog, 'Inspeção', 'Relatório salvo. Confira todas as abas, endereços, listas, assinaturas e área de impressão antes de autorizar.')
            except Exception as exc: QMessageBox.warning(dialog, 'Modelo', str(exc))
        inspect.clicked.connect(inspect_now)
        authorize = QCheckBox('Conferi o relatório: mapeamento correto, Tipo e assinaturas vazios, área de impressão adequada.')
        form.addRow(authorize)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel); form.addRow(buttons)
        def save():
            try:
                data = self.settings.model_dump()
                for key, control in controls.items(): data[key] = control.text().strip()
                if data['date_policy'] not in ('atual', 'fixa'): raise ValueError('Política deve ser atual ou fixa.')
                if not data['requester']: raise ValueError('Informe seu nome como requisitante.')
                if int(data['expected_pages']) < 1: raise ValueError('Número de páginas inválido.')
                if authorize.isChecked():
                    report = inspect_template(data['template'])
                    if data['sheet'] not in report['sheets']: raise ValueError('Aba não encontrada.')
                    data['authorized_template_sha256'] = report['sha256']
                elif data['template'] != self.settings.template or data['sheet'] != self.settings.sheet:
                    data['authorized_template_sha256'] = ''
                self.settings = type(self.settings).model_validate(data); save_settings(self.settings); dialog.accept()
            except Exception as exc: QMessageBox.warning(dialog, 'Configurações', str(exc))
        buttons.accepted.connect(save); buttons.rejected.connect(dialog.reject); dialog.exec()

    def generate(self, pdf):
        try:
            row, original = self.selected(); item = original.model_copy(deep=True); item.require_ready()
            settings = self.settings.model_copy(deep=True)
            def task():
                folder = new_output(settings.output, item.name)
                xlsx = write_excel(settings.template, folder / 'cadastro.xlsx', item, settings.sheet, settings.authorized_template_sha256)
                (folder / 'registro.json').write_text(item.model_dump_json(indent=2), encoding='utf-8')
                if pdf: export_pdf(xlsx, folder / 'cadastro.pdf', settings.sheet, item.name, settings.expected_pages)
                return folder
            def done(folder):
                self.outputs[row] = folder
                QMessageBox.information(self, 'Arquivos gerados', f'{folder}\nConfira visualmente o PDF antes de imprimir: texto, Tipo e assinaturas.')
            self.run_job(task, done)
        except Exception as exc: QMessageBox.warning(self, 'Geração', str(exc))

    def pdf_path(self):
        row, _ = self.selected(); path = self.outputs.get(row, Path('/__ausente__')) / 'cadastro.pdf'
        if not path.is_file(): raise ValueError('Gere o PDF deste item antes de continuar.')
        return path

    def preview(self):
        try:
            from PySide6.QtPdf import QPdfDocument
            from PySide6.QtPdfWidgets import QPdfView
            dialog = QDialog(self); dialog.setWindowTitle('Visualizar PDF — confira cortes e campos manuais'); dialog.resize(850, 900)
            layout = QVBoxLayout(dialog); doc = QPdfDocument(dialog); doc.load(str(self.pdf_path()))
            view = QPdfView(dialog); view.setDocument(doc); view.setPageMode(QPdfView.PageMode.MultiPage); layout.addWidget(view)
            dialog.exec()
        except Exception as exc: QMessageBox.warning(self, 'PDF', str(exc))

    def print_pdf(self):
        try:
            if os.name != 'nt': raise ValueError('Impressão disponível no Windows.')
            # Open the PDF viewer. The user invokes its print dialog explicitly.
            QMessageBox.information(self, 'Impressão', 'O PDF será aberto no visualizador do Windows. Use Ctrl+P para escolher impressora e confirmar a impressão.')
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.pdf_path())))
        except Exception as exc: QMessageBox.warning(self, 'Impressão', str(exc))

    def open_output(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(self.settings.output).resolve())))

    def closeEvent(self, event):
        if self.worker is not None:
            QMessageBox.information(self, 'Aguarde', 'Aguarde o término da operação antes de fechar.'); event.ignore()
        else: event.accept()
