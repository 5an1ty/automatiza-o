"""Editable, explicitly confirmed preview, sharing the application's item model."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QLabel, QPushButton, QMessageBox)
from cadastro_pcm.core.material_list import parse_material_list, make_item, mark_duplicates
from cadastro_pcm.core.models import money_br


class MaterialListDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Importar lista de materiais'); self.resize(950, 720)
        self.result_items = []; self.parsed = []; self.notices = []
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel('Cole descrições, NCMs e preços. Dados informados pelo usuário; NCM não verificado.'))
        self.text = QTextEdit(); self.text.setAcceptRichText(False)
        self.text.setPlaceholderText('Material | NCM | Preço unitário'); layout.addWidget(self.text, 1)
        self.analyze_button = QPushButton('Analisar lista'); layout.addWidget(self.analyze_button)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(['Material', 'NCM', 'Preço unitário', 'Situação'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 360); layout.addWidget(self.table, 1)
        self.summary = QLabel('0 itens'); self.summary.setWordWrap(True); layout.addWidget(self.summary)
        row = QHBoxLayout(); self.add_button = QPushButton('Adicionar à revisão')
        self.add_button.setEnabled(False); row.addWidget(self.add_button)
        cancel = QPushButton('Cancelar'); row.addWidget(cancel); layout.addLayout(row)
        self.analyze_button.clicked.connect(self.analyze)
        self.add_button.clicked.connect(self.confirm); cancel.clicked.connect(self.reject)
        self.text.textChanged.connect(self.invalidate)
        self.table.itemChanged.connect(self.update_status)

    def invalidate(self):
        self.add_button.setEnabled(False)
        self.summary.setText('Texto alterado. Clique em Analisar lista novamente.')

    def rows(self):
        return mark_duplicates([make_item(*(self.table.item(row, col).text() for col in range(3)),
                                         warnings=self.parsed[row].warnings)
                                for row in range(self.table.rowCount())])

    def analyze(self):
        self.parsed, self.notices = parse_material_list(self.text.toPlainText())
        self.table.blockSignals(True); self.table.setRowCount(len(self.parsed))
        for row, item in enumerate(self.parsed):
            for col, text in enumerate((item.name, item.ncm, money_br(item.unit_price) if item.unit_price is not None else '', '')):
                cell = QTableWidgetItem(text)
                if col == 3: cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(row, col, cell)
        self.table.blockSignals(False); self.update_status()
        self.add_button.setEnabled(bool(self.parsed))

    def update_status(self):
        self.table.blockSignals(True)
        for row, item in enumerate(self.rows()):
            messages = item.warnings + ([] if item.ncm else ['NCM pendente']) + ([] if item.unit_price is not None else ['Preço pendente'])
            if not item.name: messages.append('Nome pendente')
            self.table.item(row, 3).setText('; '.join(messages) or 'Revisar e aprovar')
        self.table.blockSignals(False)
        self.summary.setText(f'{self.table.rowCount()} itens. Possíveis duplicatas são mantidas. ' + ' '.join(self.notices))

    def confirm(self):
        items = self.rows()
        if not items or any(not item.name.strip() for item in items):
            QMessageBox.warning(self, 'Importar lista', 'Informe o nome de cada material antes de adicionar.'); return
        self.result_items = items
        self.accept()
