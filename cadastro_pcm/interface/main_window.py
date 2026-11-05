import json
import logging
import os
from datetime import date
from pathlib import Path
from PySide6.QtCore import QThread, Signal, QUrl, Qt
from PySide6.QtGui import QDesktopServices, QColor
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QTableWidget, QTableWidgetItem, QMessageBox,
    QFileDialog, QDialog, QFormLayout, QLineEdit, QTextEdit, QCheckBox,
    QDialogButtonBox, QLabel, QComboBox, QSplitter, QHeaderView, QProgressBar, QTabWidget, QListWidgetItem)
from .theme import QSS
from .components import label, badge, card, button, item_status
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
        layout.setContentsMargins(24, 24, 24, 24); layout.setSpacing(14)
        layout.addWidget(label('Revise antes de gerar', 'title'))
        layout.addWidget(label('Confira os dados e suas fontes. Passe o cursor sobre um campo para ver a evidência.', 'muted'))
        tabs = QTabWidget(); layout.addWidget(tabs, 1)
        forms = {}
        for title in ('Material e valores', 'Cadastro e solicitação'):
            scroll = QScrollArea(); scroll.setWidgetResizable(True)
            body = QWidget(); form = QFormLayout(body); form.setSpacing(14)
            form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
            scroll.setWidget(body); tabs.addTab(scroll, title); forms[title] = form
        self.inputs = {}
        for key, field_label in LABELS.items():
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
            group = 'Material e valores' if key in ('name', 'description', 'quantity', 'unit_price', 'total_price', 'ncm', 'manufacturer', 'manufacturer_code') else 'Cadastro e solicitação'
            forms[group].addRow(f'{field_label} · {origin}', control)
            self.inputs[key] = control
        self.confirm = QCheckBox('Revisei todos os dados, fontes e estimativas; aprovo este item.')
        layout.addWidget(self.confirm)
        layout.addWidget(QLabel('Tipo e assinaturas permanecem vazios. NCM ausente continua pendente.'))
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(lambda: self.save(item)); buttons.rejected.connect(self.reject)
        buttons.button(QDialogButtonBox.Save).setText('Salvar revisão')
        buttons.button(QDialogButtonBox.Save).setProperty('primary', True)
        buttons.button(QDialogButtonBox.Cancel).setText('Cancelar')
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
        self.setStyleSheet(QSS)
        self.setWindowTitle('Cadastro PCM')
        self.resize(1440, 960); self.setMinimumSize(1050, 760); self.setAcceptDrops(True)
        body = QWidget(); self.setCentralWidget(body); layout = QVBoxLayout(body)
        layout.setContentsMargins(24, 20, 24, 16); layout.setSpacing(14)
        header = QHBoxLayout(); headings = QVBoxLayout()
        headings.addWidget(label('Cadastro PCM', 'title'))
        headings.addWidget(label('Da cotação ao cadastro. Revise com confiança, gere com precisão.', 'muted'))
        header.addLayout(headings); header.addStretch()
        header.addWidget(button('Configurações', self.configure)); layout.addLayout(header)
        indicators = QHBoxLayout()
        self.model_indicator = badge('Modelo não configurado', 'warning')
        self.output_indicator = label('', 'muted')
        indicators.addWidget(self.model_indicator); indicators.addWidget(self.output_indicator, 1)
        layout.addLayout(indicators)
        self.steps = label('01  Configurar modelo     →     02  Adicionar PDFs     →     03  Processar     →     04  Revisar     →     05  Gerar e conferir', 'muted')
        layout.addWidget(self.steps)
        intake, incoming = card('Documentos de entrada')
        entry = QHBoxLayout()
        from PySide6.QtWidgets import QFrame
        drop = QFrame(); drop.setObjectName('dropZone'); drop_layout = QVBoxLayout(drop)
        drop_layout.addWidget(label('Solte seus PDFs aqui', 'section'))
        drop_layout.addWidget(label('Cotações e documentos de fornecedores. Você também pode selecionar arquivos.', 'muted'))
        drop_layout.addWidget(button('＋  Adicionar PDFs', self.add_files, True))
        entry.addWidget(drop, 1)
        file_area = QVBoxLayout()
        self.documents = QListWidget(); self.documents.setMinimumHeight(90); self.documents.setMaximumHeight(125)
        file_area.addWidget(self.documents)
        file_actions = QHBoxLayout()
        self.remove_button = button('Remover', self.remove_file)
        self.source_button = button('Ver texto', self.view_source)
        file_actions.addWidget(self.remove_button); file_actions.addWidget(self.source_button); file_actions.addStretch()
        file_area.addLayout(file_actions); entry.addLayout(file_area, 2); incoming.addLayout(entry)
        layout.addWidget(intake)
        review_header = QHBoxLayout()
        review_header.addWidget(label('Revisão dos itens', 'section')); review_header.addStretch()
        review_header.addWidget(button('Adicionar item manual', self.add_item))
        self.process_button = button('Processar PDFs', self.process, True)
        review_header.addWidget(self.process_button); layout.addLayout(review_header)
        split = QSplitter(Qt.Horizontal)
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(['Descrição', 'Quantidade', 'Preço unitário', 'NCM', 'Fabricante', 'Origem', 'Validação'])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True); self.table.setShowGrid(False)
        self.table.verticalHeader().hide(); self.table.verticalHeader().setDefaultSectionSize(48)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.cellDoubleClicked.connect(lambda *_: self.edit())
        self.table.itemSelectionChanged.connect(self.update_details)
        split.addWidget(self.table)
        details, detail_layout = card('Detalhes do item')
        details.setMinimumWidth(280); details.setMaximumWidth(430)
        self.detail_badge = badge('Selecione um item')
        detail_layout.addWidget(self.detail_badge)
        self.details = QTextEdit(); self.details.setReadOnly(True)
        self.details.setPlaceholderText('Os dados completos, fontes e pendências do item selecionado aparecem aqui.')
        detail_layout.addWidget(self.details)
        self.edit_button = button('Editar e revisar item', self.edit)
        self.research_button = button('Registrar pesquisa', self.research)
        detail_layout.addWidget(self.edit_button); detail_layout.addWidget(self.research_button)
        split.addWidget(details); split.setSizes([950, 330]); layout.addWidget(split, 1)
        actions_card, actions = card()
        row = QHBoxLayout(); row.addWidget(label('Gerar documentos', 'section')); row.addStretch()
        self.excel_button = button('Gerar Excel', lambda: self.generate(False))
        self.pdf_button = button('Gerar Excel + PDF', lambda: self.generate(True), True)
        row.addWidget(self.excel_button); row.addWidget(self.pdf_button); actions.addLayout(row)
        row = QHBoxLayout(); row.addWidget(label('Conferir e finalizar', 'muted')); row.addStretch()
        self.preview_button = button('Visualizar PDF', self.preview)
        self.print_button = button('Imprimir…', self.print_pdf)
        row.addWidget(self.preview_button); row.addWidget(self.print_button)
        row.addWidget(button('Abrir saída', self.open_output)); actions.addLayout(row); layout.addWidget(actions_card)
        self.progress = QProgressBar(); self.progress.setRange(0, 0); self.progress.setTextVisible(False); self.progress.hide()
        layout.addWidget(self.progress)
        self.status = label('Nenhum PDF carregado · Nenhum item identificado ainda', 'muted')
        layout.addWidget(self.status)
        self.documents.currentRowChanged.connect(lambda *_: self.update_actions())
        self.update_indicators(); self.update_actions()
        if not self.settings.requester:
            from PySide6.QtCore import QTimer
            QTimer.singleShot(0, self.configure)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls(): event.acceptProposedAction()

    def dropEvent(self, event):
        self.add_paths([url.toLocalFile() for url in event.mimeData().urls()])

    def add_paths(self, paths):
        existing = {self.document_path(i) for i in range(self.documents.count())}
        for path in paths:
            path = str(Path(path).resolve())
            if Path(path).suffix.lower() == '.pdf' and Path(path).is_file() and path not in existing:
                entry = QListWidgetItem()
                entry.setData(Qt.UserRole, path)
                self.documents.addItem(entry); self.set_document_status(entry, 'Pendente')
                existing.add(path)
        self.update_summary(); self.update_actions()

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
        self.run_job(lambda: read_pdf(entry.data(Qt.UserRole)), done)

    def remove_file(self):
        self.documents.takeItem(self.documents.currentRow())
        self.update_summary(); self.update_actions()

    def run_job(self, function, callback):
        if self.worker is not None:
            QMessageBox.information(self, 'Aguarde', 'Uma operação está em andamento.'); return
        self.centralWidget().setEnabled(False)
        self.progress.show()
        self.status.setText('Trabalhando nos documentos. Aguarde um instante…')
        self.worker = Worker(function)
        self.worker.done.connect(callback)
        self.last_error = ''
        self.worker.failed.connect(self.show_job_error)
        self.worker.finished.connect(self.job_finished)
        self.worker.start()

    def job_finished(self):
        self.worker.deleteLater(); self.worker = None
        self.centralWidget().setEnabled(True)
        self.progress.hide()
        self.update_actions()
        if getattr(self, 'last_error', ''):
            self.status.setText('Erro · ' + self.last_error)
        else:
            self.update_summary()

    def process(self):
        paths = [self.document_path(i) for i in range(self.documents.count())]
        if not paths: return
        if self.items and QMessageBox.question(self, 'Substituir lista?', 'Substituir os itens atuais? Alterações não geradas serão descartadas.') != QMessageBox.Yes: return
        def task():
            items, warnings, statuses = [], [], {}
            for path in paths:
                try:
                    pages = read_pdf(path); found = extract_items(pages)
                    if not found:
                        warnings.append(f'{Path(path).name}: layout não reconhecido. Consulte o texto extraído e adicione itens manualmente.')
                    if any(not page.text.strip() for page in pages):
                        warnings.append(f'{Path(path).name}: há páginas sem texto; revise as páginas digitalizadas.')
                    items.extend(found)
                    statuses[path] = 'Extraído' if found else 'Incompleto'
                except Exception as exc:
                    warnings.append(str(exc))
                    statuses[path] = 'Erro'
            return items, warnings, statuses
        def done(result):
            self.items, warnings, statuses = result; self.outputs = {}
            for index in range(self.documents.count()):
                entry = self.documents.item(index)
                self.set_document_status(entry, statuses.get(self.document_path(index), 'Pendente'))
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

    def document_path(self, index):
        return self.documents.item(index).data(Qt.UserRole)

    def set_document_status(self, entry, status):
        path = Path(entry.data(Qt.UserRole))
        size = path.stat().st_size / 1024 if path.exists() else 0
        entry.setText(f'{path.name}   ·   {size:.0f} KB   ·   {status}')
        entry.setToolTip(str(path))
        entry.setForeground(QColor({'Erro': '#ffa6b4', 'Incompleto': '#f2cb83', 'Extraído': '#88e2bb'}.get(status, '#b8c9e5')))

    def update_indicators(self):
        exists = bool(self.settings.template and Path(self.settings.template).is_file())
        authorized = exists and bool(self.settings.authorized_template_sha256)
        text = ('Modelo configurado · conferir alterações' if authorized else 'Modelo selecionado · autorize nas configurações') if exists else 'Configure o modelo Excel'
        self.model_indicator.setText(text)
        self.model_indicator.setProperty('tone', 'success' if authorized else 'warning')
        self.model_indicator.style().unpolish(self.model_indicator); self.model_indicator.style().polish(self.model_indicator)
        self.model_indicator.setToolTip(self.settings.template or 'Nenhum modelo selecionado')
        output = Path(self.settings.output)
        self.output_indicator.setText('Saída · ' + (output.name or str(output)))
        self.output_indicator.setToolTip(str(output))

    def update_summary(self):
        count = self.documents.count()
        ready = sum(item_status(item)[0] == 'Pronto para gerar' for item in self.items)
        files = f'{count} PDF' + ('s carregados' if count != 1 else ' carregado')
        items = f'{len(self.items)} itens · {ready} prontos para geração' if self.items else 'Nenhum item identificado ainda'
        self.status.setText(files + '  ·  ' + items)

    def update_actions(self):
        row = self.table.currentRow()
        selected = 0 <= row < len(self.items)
        ready = selected and item_status(self.items[row])[0] == 'Pronto para gerar'
        configured = bool(self.settings.template and Path(self.settings.template).is_file() and self.settings.authorized_template_sha256)
        pdf = selected and row in self.outputs and (self.outputs[row] / 'cadastro.pdf').is_file()
        self.edit_button.setEnabled(selected); self.research_button.setEnabled(selected)
        self.excel_button.setEnabled(ready and configured); self.pdf_button.setEnabled(ready and configured)
        for widget in (self.excel_button, self.pdf_button):
            widget.setToolTip('Revise e aprove o item e configure um modelo autorizado para gerar.' if not (ready and configured) else 'Gerar documentos do item selecionado')
        self.preview_button.setEnabled(bool(pdf)); self.print_button.setEnabled(bool(pdf))
        self.process_button.setEnabled(self.documents.count() > 0)
        self.remove_button.setEnabled(self.documents.currentRow() >= 0)
        self.source_button.setEnabled(self.documents.currentRow() >= 0)

    def update_details(self):
        row = self.table.currentRow()
        if 0 <= row < len(self.items):
            item = self.items[row]; status, tone = item_status(item)
            self.detail_badge.setText(status); self.detail_badge.setProperty('tone', tone)
            lines = [item.name, '', 'DESCRIÇÃO TÉCNICA', item.description or 'Pendente', '',
                     f'Quantidade: {item.quantity if item.quantity is not None else "Pendente"}',
                     f'Preço unitário: {money_br(item.unit_price)}', f'NCM fiscal: {item.ncm or "Pendente"}',
                     f'Fabricante: {item.manufacturer or "Pendente"}', '', 'ORIGEM E EVIDÊNCIAS']
            for field, evidence in item.evidence.items():
                lines.extend([f'{LABELS.get(field, field)} · {evidence.origin}', evidence.source, evidence.excerpt])
                if evidence.estimated: lines.append('Estimativa — requer aprovação humana')
            lines.extend(['', 'OBSERVAÇÕES', 'Tipo e assinaturas permanecem vazios.'])
            if not item.ncm: lines.append('NCM fiscal pendente: confira a aplicabilidade antes de gerar.')
            if not item.approved: lines.append('Revise as informações e aprove no editor.')
            self.details.setPlainText('\n'.join(lines))
        else:
            self.detail_badge.setText('Selecione um item'); self.detail_badge.setProperty('tone', 'neutral'); self.details.clear()
        self.detail_badge.style().unpolish(self.detail_badge); self.detail_badge.style().polish(self.detail_badge)
        self.update_actions()

    def refresh(self):
        selected = self.table.currentRow()
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.items))
        for row, item in enumerate(self.items):
            status, tone = item_status(item)
            values = [item.name, str(item.quantity or ''), money_br(item.unit_price), item.ncm or 'Pendente', item.manufacturer or '—',
                      '; '.join(sorted({e.source for e in item.evidence.values() if e.source})), status]
            for col, value in enumerate(values):
                cell = QTableWidgetItem(value); cell.setToolTip(value)
                if col == 6:
                    cell.setForeground(QColor({'success': '#88e2bb', 'warning': '#f2cb83', 'neutral': '#a7ccff'}[tone]))
                if col in (1, 2): cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table.setItem(row, col, cell)
        for col, width in enumerate((220, 100, 125, 105, 125, 160, 170)):
            self.table.setColumnWidth(col, width)
        if self.items: self.table.selectRow(min(max(selected, 0), len(self.items) - 1))
        self.table.blockSignals(False)
        self.update_details(); self.update_summary()

    def show_job_error(self, error):
        self.last_error = error
        QMessageBox.warning(self, 'Não foi possível concluir', error)

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
        dialog = QDialog(self); dialog.setWindowTitle('Configurações'); dialog.resize(780, 720)
        outer = QVBoxLayout(dialog); outer.setContentsMargins(24, 24, 24, 24); outer.setSpacing(16)
        outer.addWidget(label('Prepare seu espaço de trabalho', 'title'))
        outer.addWidget(label('Escolha o modelo oficial e defina os padrões para os próximos itens.', 'muted'))
        tabs = QTabWidget(); outer.addWidget(tabs)
        forms = {}
        for title in ('Modelo e saída', 'Solicitante e data', 'Padrões do cadastro'):
            panel = QWidget(); panel_form = QFormLayout(panel); panel_form.setSpacing(14)
            panel_form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
            forms[title] = panel_form; tabs.addTab(panel, title)
        form = forms['Modelo e saída']; controls = {}
        names = {'requester':'Requisitante padrão', 'area':'Área', 'date_policy':'Data: atual ou fixa', 'fixed_date':'Data fixa AAAA-MM-DD',
                 'output':'Pasta de saída', 'template':'Caminho do modelo Excel', 'sheet':'Aba do formulário',
                 'classification':'Classificação', 'unit':'Grupo UM', 'group':'Grupo de Itens', 'purpose':'Finalidade', 'expected_pages':'Páginas esperadas'}
        for key, field_label in names.items():
            control = QLineEdit(str(getattr(self.settings, key)))
            if key == 'date_policy':
                control = QComboBox(); control.addItems(['atual', 'fixa']); control.setCurrentText(self.settings.date_policy)
            controls[key] = control
            section = 'Solicitante e data' if key in ('requester', 'area', 'date_policy', 'fixed_date') else ('Padrões do cadastro' if key in ('classification', 'unit', 'group', 'purpose') else 'Modelo e saída')
            control.setToolTip(str(getattr(self.settings, key)))
            forms[section].addRow(field_label, control)
        pick = QPushButton('Selecionar modelo…'); form.addRow(pick)
        pick.clicked.connect(lambda: controls['template'].setText(QFileDialog.getOpenFileName(dialog, 'Modelo oficial', '', 'Excel (*.xlsx)')[0] or controls['template'].text()))
        output_pick = QPushButton('Escolher pasta de saída…'); form.addRow(output_pick)
        output_pick.clicked.connect(lambda: controls['output'].setText(QFileDialog.getExistingDirectory(dialog, 'Pasta de saída', controls['output'].text()) or controls['output'].text()))
        form.addRow(label('O modelo original é preservado. Confira o relatório antes de autorizar o preenchimento.', 'muted'))
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
        authorize = QCheckBox('Conferi o relatório e autorizo este modelo.'); authorize.setToolTip('Mapeamento correto, Tipo e assinaturas vazios e área de impressão adequada.')
        form.addRow(authorize)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel); outer.addWidget(buttons)
        buttons.button(QDialogButtonBox.Save).setText('Salvar configurações')
        buttons.button(QDialogButtonBox.Save).setProperty('primary', True)
        buttons.button(QDialogButtonBox.Cancel).setText('Cancelar')
        def save():
            try:
                data = self.settings.model_dump()
                for key, control in controls.items(): data[key] = control.currentText() if isinstance(control, QComboBox) else control.text().strip()
                if data['date_policy'] not in ('atual', 'fixa'): raise ValueError('Política deve ser atual ou fixa.')
                if not data['requester']: raise ValueError('Informe seu nome como requisitante.')
                if int(data['expected_pages']) < 1: raise ValueError('Número de páginas inválido.')
                if authorize.isChecked():
                    report = inspect_template(data['template'])
                    if data['sheet'] not in report['sheets']: raise ValueError('Aba não encontrada.')
                    data['authorized_template_sha256'] = report['sha256']
                elif data['template'] != self.settings.template or data['sheet'] != self.settings.sheet:
                    data['authorized_template_sha256'] = ''
                self.settings = type(self.settings).model_validate(data); save_settings(self.settings); self.update_indicators(); self.update_actions(); self.status.setText('Configurações salvas. Os padrões serão usados nos próximos itens.'); dialog.accept()
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
                self.update_actions()
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
