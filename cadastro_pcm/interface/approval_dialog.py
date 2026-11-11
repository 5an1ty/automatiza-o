"""Bounded approval dialogs: fixed summary/footer and a scrollable detail list."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QDialog, QVBoxLayout, QListWidget, QListWidgetItem, QDialogButtonBox
from .components import label


class ApprovalDialog(QDialog):
    ALLOW_INCOMPLETE = 2
    COMPLETE_ONLY = 1

    def __init__(self, items, parent=None, *, result=False, approved_count=0):
        super().__init__(parent)
        self.setWindowTitle('Aprovação concluída' if result else 'Aprovação do material' if len(items) == 1 else 'Aprovação em lote')
        layout = QVBoxLayout(self)
        complete = sum(not item.pending_fields() for item in items)
        pending = len(items) - complete
        authorized = sum(item.approved and bool(item.pending_fields()) and not item.blocking_fields() for item in items)
        summary = (f'{approved_count} materiais aprovados\n{authorized} aprovados com pendências\n'
                   f'{len(items) - approved_count} não aprovados nesta operação' if result else
                   f'{len(items)} materiais selecionados\n{complete} completos\n{pending} com pendências')
        self.summary = label(summary)
        layout.addWidget(self.summary)
        self.details = QListWidget()
        self.details.setWordWrap(True)
        self.details.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.details.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        for item in items:
            missing = item.pending_fields()
            structural = item.structural_pending_fields()
            if structural:
                status = 'Bloqueado: identificação ou requisitante ausente/inválido'
            elif item.approved and missing and not item.blocking_fields():
                status = 'Aprovado com pendências'
            elif item.approved and not missing:
                status = 'Aprovado'
            else:
                status = 'Pendente de aprovação' if missing else 'Completo, apto para aprovação'
            text = f'{item.name or "Sem identificação"}\n{status}\n' + ('Pendências: ' + '; '.join(missing) if missing else 'Sem pendências')
            row = QListWidgetItem(text); row.setToolTip(text); self.details.addItem(row)
        layout.addWidget(self.details, 1)
        self.buttons = QDialogButtonBox()
        if result:
            self.buttons.addButton('Fechar', QDialogButtonBox.RejectRole)
            self.buttons.rejected.connect(self.reject)
        else:
            layout.addWidget(label('Deseja permitir materiais incompletos?' if pending else 'Confirma a revisão e aprovação dos materiais completos?'))
            layout.addWidget(label('Preço e NCM ausentes continuarão vazios. Após editar o Excel, exporte um novo PDF: o anterior não é atualizado.', 'muted'))
            yes = self.buttons.addButton('Sim, aprovar incompletos', QDialogButtonBox.ActionRole)
            yes.setEnabled(bool(pending))
            yes.clicked.connect(lambda: self.done(self.ALLOW_INCOMPLETE))
            no = self.buttons.addButton('Não, somente completos' if pending else 'Aprovar completos', QDialogButtonBox.ActionRole)
            no.clicked.connect(lambda: self.done(self.COMPLETE_ONLY))
            cancel = self.buttons.addButton('Cancelar', QDialogButtonBox.RejectRole)
            cancel.setDefault(True)
            self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        # Honor the active monitor rather than growing with item count.
        screen = parent.screen() if parent is not None else self.screen()
        area = screen.availableGeometry()
        if area.width() < 760:
            self.buttons.setOrientation(Qt.Vertical)
        self.setMaximumSize(area.width(), area.height())
        self.resize(min(850, area.width()), min(560, area.height()))

    def showEvent(self, event):
        super().showEvent(event)
        area = self.screen().availableGeometry()
        self.resize(min(self.width(), area.width()), min(self.height(), area.height()))
        self.move(max(area.left(), min(self.x(), area.right() - self.width() + 1)),
                  max(area.top(), min(self.y(), area.bottom() - self.height() + 1)))
