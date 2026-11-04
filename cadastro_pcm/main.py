import logging
import sys
from PySide6.QtWidgets import QApplication, QMessageBox
from cadastro_pcm.core.settings import settings_path


def main():
    folder = settings_path().parent
    folder.mkdir(parents=True, exist_ok=True)
    from logging.handlers import RotatingFileHandler
    logging.basicConfig(level=logging.INFO, handlers=[RotatingFileHandler(folder / 'cadastro.log', maxBytes=2_000_000, backupCount=3, encoding='utf-8')])
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    try:
        from cadastro_pcm.interface.main_window import Window
        window = Window(); window.show()
    except Exception as exc:
        logging.exception('Inicialização')
        QMessageBox.critical(None, 'Não foi possível iniciar', str(exc))
        return 1
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
