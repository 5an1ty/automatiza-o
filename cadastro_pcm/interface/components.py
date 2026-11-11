from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt


def label(text, role=None):
    widget = QLabel(text)
    widget.setTextFormat(Qt.PlainText)
    widget.setWordWrap(True)
    if role:
        widget.setObjectName(role)
    return widget


def badge(text, tone='neutral'):
    widget = label(text, 'badge')
    widget.setProperty('tone', tone)
    return widget


def card(title=None):
    frame = QFrame(); frame.setObjectName('card')
    layout = QVBoxLayout(frame); layout.setContentsMargins(18, 16, 18, 16); layout.setSpacing(12)
    if title:
        layout.addWidget(label(title, 'section'))
    return frame, layout


def button(text, handler, primary=False):
    widget = QPushButton(text)
    widget.setProperty('primary', primary)
    widget.setCursor(Qt.PointingHandCursor)
    widget.clicked.connect(handler)
    return widget


def item_status(item):
    if item.pending_fields():
        return 'Incompleto', 'warning'
    if item.price_fill_later:
        return ('Preço a preencher · aprovado', 'success') if item.approved else ('Preço a preencher · aprovar', 'warning')
    if item.approved:
        return 'Pronto para gerar', 'success'
    if any(e.origin == 'manual' for e in item.evidence.values()):
        return 'Revisado · aprovar', 'warning'
    if any(e.origin == 'pdf' for e in item.evidence.values()):
        return 'Extraído · revisar', 'neutral'
    return 'Pendente', 'warning'
