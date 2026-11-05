"""Paleta compartilhada por janelas, diálogos e componentes."""
QSS = '''
QWidget { background: #11151e; color: #e4e9f2; font-family: 'Segoe UI', sans-serif; font-size: 13px; }
QMainWindow, QDialog { background: #11151e; }
QFrame#card { background: #1a202c; border: 1px solid #2c3545; border-radius: 12px; }
QFrame#dropZone { background: #182331; border: 1px dashed #537799; border-radius: 12px; }
QLabel { background: transparent; }
QLabel#title { font-size: 28px; font-weight: 700; }
QLabel#section { font-size: 16px; font-weight: 600; }
QLabel#muted { color: #9daec5; }
QLabel#badge { border-radius: 7px; padding: 5px 9px; background: #283348; color: #b8c9e5; }
QLabel#badge[tone='success'] { background: #183a34; color: #88e2bb; }
QLabel#badge[tone='warning'] { background: #443722; color: #f2cb83; }
QLabel#badge[tone='error'] { background: #472831; color: #ffa6b4; }
QPushButton { background: #283345; border: 1px solid #3a4960; border-radius: 8px; padding: 9px 14px; font-weight: 600; }
QPushButton:hover { background: #364862; border-color: #7291ba; }
QPushButton:pressed { background: #1f2b3d; }
QPushButton[primary='true'] { background: #83b5ff; color: #101d32; border-color: #83b5ff; }
QPushButton[primary='true']:hover { background: #a7ccff; }
QPushButton:disabled { background: #202633; color: #627188; border-color: #293343; }
QLineEdit, QTextEdit, QComboBox, QListWidget { background: #141b27; border: 1px solid #344155; border-radius: 7px; padding: 8px; selection-background-color: #365a86; }
QLineEdit:focus, QTextEdit:focus, QComboBox:focus { border-color: #83b5ff; }
QComboBox QAbstractItemView { background: #1a2433; selection-background-color: #365a86; }
QTableWidget { background: #171e2a; alternate-background-color: #1c2533; border: 1px solid #303b4d; border-radius: 9px; gridline-color: #2b3545; selection-background-color: #304f77; selection-color: #ffffff; }
QTableWidget::item { padding: 10px; border-bottom: 1px solid #263142; }
QTableWidget::item:selected { background: #304f77; }
QHeaderView::section { background: #242f40; color: #b6c6db; border: none; padding: 12px 9px; font-weight: 600; }
QListWidget::item { padding: 7px; border-radius: 5px; }
QListWidget::item:selected { background: #304f77; }
QTabWidget::pane { border: 1px solid #303b4d; border-radius: 8px; padding: 10px; }
QTabBar::tab { background: #202938; padding: 10px 16px; margin-right: 4px; border-radius: 6px; color: #9daec5; }
QTabBar::tab:selected { background: #324762; color: #e4e9f2; }
QScrollArea { border: none; }
QScrollBar:vertical { background: #171e2a; width: 10px; }
QScrollBar::handle:vertical { background: #43536c; border-radius: 5px; min-height: 30px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QCheckBox { spacing: 8px; padding: 5px; }
QCheckBox::indicator { width: 18px; height: 18px; }
QProgressBar { border: none; background: #202c3c; max-height: 4px; }
QProgressBar::chunk { background: #83b5ff; }
QToolTip { background: #28354a; color: #edf3ff; border: 1px solid #5d7699; padding: 8px; }
'''
