"""Application entry point."""
import sys

from PySide6.QtWidgets import QApplication

from screenshot_to_word.main_window import MainWindow


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("Screenshot to Word")
    app.setOrganizationName("ScreenshotToWord")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
