"""Application entry point."""
import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from screenshot_to_word.main_window import MainWindow


def main() -> None:
    # Give Windows a stable application identity so the taskbar groups this app
    # under its own icon instead of the generic Python interpreter icon.
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "ScreenshotToWord.DesktopApp"
            )
        except (AttributeError, OSError):
            pass

    app = QApplication(sys.argv)
    app.setApplicationName("Screenshot to Word")
    app.setOrganizationName("ScreenshotToWord")

    assets_dir = Path(__file__).resolve().parent / "assets"
    icon_path = assets_dir / "app_icon.ico"
    if not icon_path.exists():
        icon_path = assets_dir / "app_icon.png"
    app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.setWindowIcon(QIcon(str(icon_path)))
    window.show()
    sys.exit(app.exec())
