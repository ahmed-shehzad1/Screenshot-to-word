"""Main GUI, floating heading bar, insertion controls, and preferences."""
from PySide6.QtCore import Qt, QSettings, QUrl
from PySide6.QtGui import QDesktopServices
import win32com.client

from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFormLayout, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QSpinBox,
    QVBoxLayout, QWidget, QFrame,
)

from screenshot_to_word.clipboard_monitor import ClipboardMonitor, current_clipboard_image
from screenshot_to_word.word_writer import WordAutomationError, WordWriter


class SettingsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Screenshot to Word — Settings")
        self.setMinimumWidth(340)
        form = QFormLayout()
        self.interval = QSpinBox()
        self.interval.setRange(100, 5000)
        self.interval.setSingleStep(100)
        self.interval.setSuffix(" ms")
        self.interval.setValue(int(settings.value("monitor_interval_ms", 300)))
        self.width = QDoubleSpinBox()
        self.width.setRange(2.0, 6.5)
        self.width.setSingleStep(0.25)
        self.width.setSuffix(" inches")
        self.width.setValue(float(settings.value("image_width_inches", 6.3)))
        self.width.setToolTip("Maximum image width. Images keep their original proportions.")
        self.mode = QComboBox()
        self.mode.addItem("Ask before inserting when an error occurs", "warn")
        self.mode.addItem("Show errors in status only", "quiet")
        idx = self.mode.findData(settings.value("error_mode", "warn"))
        if idx >= 0:
            self.mode.setCurrentIndex(idx)
        form.addRow("Clipboard check interval:", self.interval)
        form.addRow("Maximum screenshot width:", self.width)
        form.addRow("Error notifications:", self.mode)
        save = QPushButton("Save settings")
        cancel = QPushButton("Cancel")
        save.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)
        buttons = QHBoxLayout()
        buttons.addStretch(1)
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addLayout(buttons)

    def save_settings(self):
        self.settings.setValue("monitor_interval_ms", self.interval.value())
        self.settings.setValue("image_width_inches", self.width.value())
        self.settings.setValue("error_mode", self.mode.currentData())


class FloatingHeadingBar(QDialog):
    """Small always-on-top companion bar for task headings and quick insertion."""
    insert_heading_requested = None
    insert_screenshot_requested = None

    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("Screenshot to Word — Quick Bar")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setMinimumWidth(300)
        self.heading_input = QLineEdit()
        self.heading_input.setPlaceholderText("Task heading (e.g. Task 3 — Row Reduction)")
        self.heading_button = QPushButton("Insert Heading")
        self.screenshot_button = QPushButton("Insert Screenshot")
        self.check_word_button = QPushButton("W")
        self.check_word_button.setFixedSize(34, 34)
        self.check_word_button.setToolTip("Open Microsoft Word / check the active document")
        self.check_word_button.setAccessibleName("Check Word")
        self.heading_button.clicked.connect(self._emit_heading)
        self.screenshot_button.clicked.connect(self._emit_screenshot)
        # Show Word only while the user holds this compact preview button.
        self.check_word_button.pressed.connect(self._emit_check_word_pressed)
        self.check_word_button.released.connect(self._emit_check_word_released)
        self.heading_input.returnPressed.connect(self._emit_heading)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(QLabel("Quick actions — stays on top"))
        layout.addWidget(self.heading_input)
        row = QHBoxLayout()
        row.addWidget(self.heading_button)
        row.addWidget(self.screenshot_button)
        row.addWidget(self.check_word_button)
        layout.addLayout(row)
        self._heading_callback = None
        self._screenshot_callback = None
        self._check_word_callback = None
        self._check_word_release_callback = None
        self.setStyleSheet("""
            QDialog { background: #111827; color: #e5e7eb; }
            QLabel { color: #f8fafc; font-weight: 700; }
            QLineEdit { color: #f8fafc; background: #0b1220; border: 1px solid #475569; border-radius: 7px; padding: 8px; }
            QPushButton { color: #f8fafc; background: #263449; border: 1px solid #475569; border-radius: 7px; padding: 8px 10px; font-weight: 600; }
            QPushButton:hover { background: #344761; }
        """)

    def set_callbacks(
        self, heading_callback, screenshot_callback,
        check_word_callback=None, check_word_release_callback=None,
    ):
        self._heading_callback = heading_callback
        self._screenshot_callback = screenshot_callback
        self._check_word_callback = check_word_callback
        self._check_word_release_callback = check_word_release_callback

    def _emit_check_word_pressed(self):
        if self._check_word_callback:
            self._check_word_callback()

    def _emit_check_word_released(self):
        if self._check_word_release_callback:
            self._check_word_release_callback()

    def _emit_heading(self):
        if self._heading_callback:
            self._heading_callback(self.heading_input.text())

    def _emit_screenshot(self):
        if self._screenshot_callback:
            self._screenshot_callback()


class AboutDialog(QDialog):
    """Polished About and open-source support dialog."""
    GITHUB_PROFILE = "https://github.com/ahmed-shehzad1"
    REPOSITORY_URL = "https://github.com/ahmed-shehzad1/Screenshot-to-word"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Screenshot to Word")
        self.setMinimumWidth(500)
        self.setStyleSheet("""
            QDialog { background: #0b1220; color: #e5edf8; }
            QLabel { color: #cbd5e1; font-size: 10pt; }
            QLabel#heroName { color: #f8fafc; font-size: 21pt; font-weight: 800; }
            QLabel#eyebrow { color: #60a5fa; font-size: 9pt; font-weight: 700; letter-spacing: 1px; }
            QFrame#panel { background: #111c2e; border: 1px solid #27364d; border-radius: 14px; }
            QPushButton { color: #f8fafc; background: #243650; border: 1px solid #3a506f; border-radius: 9px; padding: 11px 14px; font-weight: 700; }
            QPushButton:hover { background: #304967; }
            QPushButton#primary { background: #2563eb; border-color: #3b82f6; }
            QPushButton#primary:hover { background: #1d4ed8; }
            QPushButton#quiet { background: transparent; border-color: #334155; }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(22, 22, 22, 20)
        root.setSpacing(14)

        hero = QFrame()
        hero.setObjectName("panel")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(18, 18, 18, 18)
        logo = QLabel("SW")
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo.setFixedSize(66, 66)
        logo.setStyleSheet("background:#2563eb; color:white; border-radius:16px; font-size:21pt; font-weight:900;")
        intro = QVBoxLayout()
        eyebrow = QLabel("THE DEVELOPER")
        eyebrow.setObjectName("eyebrow")
        name = QLabel("Ahmed Shahzad")
        name.setObjectName("heroName")
        subtitle = QLabel("Student · Air University, Islamabad · 21")
        subtitle.setWordWrap(True)
        intro.addWidget(eyebrow)
        intro.addWidget(name)
        intro.addWidget(subtitle)
        hero_layout.addWidget(logo)
        hero_layout.addSpacing(14)
        hero_layout.addLayout(intro, 1)
        root.addWidget(hero)

        story_panel = QFrame()
        story_panel.setObjectName("panel")
        story_layout = QVBoxLayout(story_panel)
        story_layout.setContentsMargins(18, 16, 18, 16)
        story_layout.setSpacing(8)
        story_heading = QLabel("Why I built this")
        story_heading.setStyleSheet("color:#f8fafc; font-size:12pt; font-weight:800;")
        story = QLabel("I love coding things that can help people. Screenshot to Word is a small project to make capturing and organizing screenshots in Microsoft Word simpler and faster.")
        story.setWordWrap(True)
        story.setMinimumHeight(54)
        story_layout.addWidget(story_heading)
        story_layout.addWidget(story)
        root.addWidget(story_panel)

        support_panel = QFrame()
        support_panel.setObjectName("panel")
        support_layout = QVBoxLayout(support_panel)
        support_layout.setContentsMargins(18, 16, 18, 18)
        support_layout.setSpacing(9)
        support_heading = QLabel("Enjoying Screenshot to Word?")
        support_heading.setStyleSheet("color:#f8fafc; font-size:13pt; font-weight:800;")
        support_copy = QLabel("Thanks for trying it! Following me and starring the repository helps other people discover the project and encourages me to build more useful tools.")
        support_copy.setWordWrap(True)
        support_layout.addWidget(support_heading)
        support_layout.addWidget(support_copy)
        buttons = QHBoxLayout()
        profile_button = QPushButton("↗  Follow on GitHub")
        profile_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.GITHUB_PROFILE)))
        star_button = QPushButton("★  Star the repository")
        star_button.setObjectName("primary")
        star_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.REPOSITORY_URL)))
        buttons.addWidget(profile_button)
        buttons.addWidget(star_button)
        support_layout.addLayout(buttons)
        root.addWidget(support_panel)

        footer = QHBoxLayout()
        footer.addStretch(1)
        close_button = QPushButton("Done")
        close_button.setObjectName("quiet")
        close_button.clicked.connect(self.accept)
        footer.addWidget(close_button)
        root.addLayout(footer)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ScreenshotToWord", "ScreenshotToWord")
        self.setWindowTitle("Screenshot to Word")
        # Replace the default Windows title bar with a branded, draggable title bar.
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint)
        self.setMinimumSize(620, 600)
        self.resize(760, 700)
        self._drag_position = None
        self.writer = WordWriter()
        self.monitor = ClipboardMonitor(int(self.settings.value("monitor_interval_ms", 300)), self)
        self.monitor.image_detected.connect(self._insert_image)
        self.monitor.status_changed.connect(self._set_status)
        self._auto_enabled = True

        titlebar = QWidget()
        titlebar.setObjectName("customTitleBar")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(14, 8, 8, 8)
        titlebar_layout.setSpacing(10)
        brand_mark = QLabel("SW")
        brand_mark.setObjectName("brandMark")
        brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_mark.setFixedSize(34, 34)
        brand_name = QLabel("Screenshot to Word")
        brand_name.setObjectName("brandName")
        brand_hint = QLabel("SCREEN CAPTURE WORKFLOW")
        brand_hint.setObjectName("brandHint")
        brand_text = QVBoxLayout()
        brand_text.setSpacing(0)
        brand_text.addWidget(brand_name)
        brand_text.addWidget(brand_hint)
        titlebar_layout.addWidget(brand_mark)
        titlebar_layout.addLayout(brand_text)
        titlebar_layout.addStretch(1)
        self.minimize_button = QPushButton("—")
        self.minimize_button.setObjectName("windowControl")
        self.minimize_button.setFixedSize(36, 30)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.maximize_button = QPushButton("□")
        self.maximize_button.setObjectName("windowControl")
        self.maximize_button.setFixedSize(36, 30)
        self.maximize_button.clicked.connect(self._toggle_maximize)
        close_button = QPushButton("×")
        close_button.setObjectName("closeControl")
        close_button.setFixedSize(36, 30)
        close_button.clicked.connect(self.close)
        titlebar_layout.addWidget(self.minimize_button)
        titlebar_layout.addWidget(self.maximize_button)
        titlebar_layout.addWidget(close_button)

        title = QLabel("Screenshot to Word")
        title.setObjectName("appTitle")
        subtitle = QLabel("Append Snipping Tool screenshots to your open Microsoft Word document.")
        subtitle.setWordWrap(True)
        self.document_label = QLabel("Word document: checking…")
        self.document_label.setWordWrap(True)
        self.status_label = QLabel("Status: ready")
        self.status_label.setWordWrap(True)
        self.clipboard_label = QLabel("Clipboard: not checked yet")
        self.clipboard_label.setWordWrap(True)

        self.auto_checkbox = QCheckBox("Automatically insert each new clipboard screenshot")
        self.auto_checkbox.setChecked(True)
        self.auto_checkbox.toggled.connect(self._toggle_auto)
        self.start_button = QPushButton("Start automatic insertion")
        self.start_button.clicked.connect(self._start_stop_monitoring)
        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self._pause_resume)
        self.pause_button.setEnabled(False)
        self.manual_button = QPushButton("Insert Screenshot Now")
        self.manual_button.clicked.connect(self._manual_insert)
        self.heading_button = QPushButton("Insert Task Heading")
        self.heading_button.clicked.connect(self._insert_heading_from_main)
        self.quickbar_button = QPushButton("Show Floating Heading Bar")
        self.quickbar_button.clicked.connect(self._toggle_quickbar)
        self.settings_button = QPushButton("Settings")
        self.settings_button.clicked.connect(self._open_settings)
        self.about_button = QPushButton("About / Support")
        self.about_button.clicked.connect(self._show_about)

        controls = QHBoxLayout()
        controls.addWidget(self.start_button)
        controls.addWidget(self.pause_button)
        manual_row = QHBoxLayout()
        manual_row.addWidget(self.manual_button)
        manual_row.addWidget(self.heading_button)
        tools_row = QHBoxLayout()
        tools_row.addWidget(self.quickbar_button)
        tools_row.addWidget(self.settings_button)
        tools_row.addWidget(self.about_button)

        note = QLabel("Tip: open the intended Word document first. Screenshots are inserted locally; save the document in Word when ready.")
        note.setWordWrap(True)
        note.setObjectName("note")

        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(18, 0, 18, 18)
        layout.addWidget(titlebar)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(4)
        layout.addWidget(self.document_label)
        layout.addWidget(self.auto_checkbox)
        layout.addLayout(controls)
        layout.addLayout(manual_row)
        layout.addLayout(tools_row)
        layout.addWidget(self.clipboard_label)
        layout.addWidget(self.status_label)
        layout.addWidget(note)
        layout.addStretch(1)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        # Premium dark UI with high-contrast controls and clearer visual grouping.
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #111827; color: #e5e7eb; font-size: 10pt; }
            QWidget#customTitleBar { background: #0b1220; border-bottom: 1px solid #293548; }
            QLabel#brandMark { background: #2563eb; color: white; border-radius: 9px; font-size: 12pt; font-weight: 900; }
            QLabel#brandName { color: #f8fafc; font-size: 10pt; font-weight: 800; }
            QLabel#brandHint { color: #64748b; font-size: 7pt; font-weight: 700; }
            QPushButton#windowControl { background: transparent; border: none; border-radius: 5px; padding: 2px; color: #cbd5e1; font-size: 12pt; }
            QPushButton#windowControl:hover { background: #263449; }
            QPushButton#closeControl { background: transparent; border: none; border-radius: 5px; padding: 2px; color: #cbd5e1; font-size: 17pt; }
            QPushButton#closeControl:hover { background: #dc2626; color: white; }
            QWidget#centralwidget { background: #111827; }
            QLabel { color: #cbd5e1; }
            QLabel#appTitle { font-size: 25px; font-weight: 800; color: #f8fafc; padding-bottom: 2px; }
            QLabel#note { color: #94a3b8; background: #172033; border: 1px solid #293548; border-radius: 8px; padding: 10px; }
            QPushButton { color: #f8fafc; background: #263449; border: 1px solid #3a4b63; border-radius: 8px; padding: 10px 12px; font-weight: 600; }
            QPushButton:hover { background: #344761; border-color: #647da0; }
            QPushButton:pressed { background: #1d293b; }
            QPushButton:disabled { color: #718096; background: #1b2534; border-color: #273447; }
            QCheckBox { color: #e2e8f0; spacing: 8px; padding: 8px 0; }
            QCheckBox::indicator { width: 17px; height: 17px; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { color: #f8fafc; background: #0b1220; selection-background-color: #2563eb; padding: 8px; border: 1px solid #37465c; border-radius: 7px; }
            QComboBox QAbstractItemView { background: #172033; color: #f8fafc; selection-background-color: #2563eb; }
            QScrollBar:vertical { background: #111827; width: 10px; margin: 0; }
            QScrollBar::handle:vertical { background: #37465c; border-radius: 4px; min-height: 24px; }
            QToolTip { color: #f8fafc; background: #0b1220; border: 1px solid #475569; }
        """)
        self.quickbar = FloatingHeadingBar()
        self.quickbar.set_callbacks(
            self._insert_heading,
            self._manual_insert,
            self._show_word_preview,
            self._hide_word_preview,
        )
        self._refresh_document()
        self._update_auto_controls()

    def _toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
            self.maximize_button.setText("□")
        else:
            self.showMaximized()
            self.maximize_button.setText("❐")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            child = self.childAt(event.position().toPoint())
            # Drag only when the click is on the title-bar background/labels.
            if child is not None and (child.objectName() in {"customTitleBar", "brandName", "brandHint", "brandMark"}):
                self._drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._drag_position is not None and event.buttons() & Qt.MouseButton.LeftButton and not self.isMaximized():
            self.move(event.globalPosition().toPoint() - self._drag_position)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_position = None
        super().mouseReleaseEvent(event)

    def _show_about(self):
        AboutDialog(self).exec()

    def _toggle_auto(self, checked):
        self._auto_enabled = checked
        if checked:
            if not self.monitor.running:
                self._start_monitoring()
        else:
            self._stop_monitoring()
            self._set_status("Automatic insertion disabled. Use Insert Screenshot Now for manual insertion.")
        self._update_auto_controls()

    def _update_auto_controls(self):
        self.start_button.setEnabled(self._auto_enabled)
        self.pause_button.setEnabled(self._auto_enabled and self.monitor.running)
        self.pause_button.setText("Pause" if self.monitor.running else "Resume")
        self.start_button.setText("Stop automatic insertion" if self.monitor.running else "Start automatic insertion")

    def _start_monitoring(self):
        if not self._auto_enabled:
            return
        try:
            name = self.writer.active_document_name()
        except WordAutomationError as exc:
            QMessageBox.warning(self, "Word not ready", str(exc))
            return
        self.monitor.set_interval(int(self.settings.value("monitor_interval_ms", 300)))
        self.monitor.start()
        self._set_status(f"Automatic insertion is running for {name}.")
        self._update_auto_controls()

    def _stop_monitoring(self):
        self.monitor.stop()
        self._update_auto_controls()

    def _start_stop_monitoring(self):
        if self.monitor.running:
            self._stop_monitoring()
        else:
            self._start_monitoring()

    def _pause_resume(self):
        if self.monitor.running:
            self.monitor.stop()
            self._set_status("Automatic insertion paused. Use Resume or manual insertion.")
            self._update_auto_controls()
        else:
            self._start_monitoring()

    def _insert_image(self, image):
        try:
            width = float(self.settings.value("image_width_inches", 6.3))
            document_name = self.writer.append_image(image, max_width_inches=width)
            self._set_status(f"Inserted screenshot into {document_name}.")
            self._refresh_document()
        except WordAutomationError as exc:
            self._show_error("Screenshot not inserted", str(exc))
            self._set_status(str(exc))

    def _manual_insert(self):
        image = current_clipboard_image()
        if image is None:
            self.clipboard_label.setText("Clipboard: no image found. Take a screenshot first.")
            self._set_status("No clipboard image available to insert.")
            return
        self.clipboard_label.setText(f"Clipboard: image ready ({image.width} × {image.height})")
        self._insert_image(image)

    def _insert_heading_from_main(self):
        from PySide6.QtWidgets import QInputDialog
        heading, ok = QInputDialog.getText(self, "Insert Task Heading", "Heading text:")
        if ok:
            self._insert_heading(heading)

    def _insert_heading(self, heading):
        try:
            name = self.writer.append_heading(heading)
            self._set_status(f"Inserted heading into {name}.")
            self._refresh_document()
        except WordAutomationError as exc:
            self._show_error("Heading not inserted", str(exc))
            self._set_status(str(exc))

    def _toggle_quickbar(self):
        if self.quickbar.isVisible():
            self.quickbar.hide()
            self.quickbar_button.setText("Show Floating Heading Bar")
        else:
            self.quickbar.show()
            self.quickbar.raise_()
            self.quickbar_button.setText("Hide Floating Heading Bar")

    def _open_settings(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            dialog.save_settings()
            self.monitor.set_interval(int(self.settings.value("monitor_interval_ms", 300)))
            self._set_status("Settings saved.")

    def _show_error(self, title, message):
        if self.settings.value("error_mode", "warn") == "warn":
            QMessageBox.warning(self, title, message)

    def _show_word_preview(self):
        """Show Word while the floating W button is held down."""
        try:
            try:
                self._word_preview_app = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                self._word_preview_app = win32com.client.Dispatch("Word.Application")
            word = self._word_preview_app
            word.Visible = True
            if word.Documents.Count == 0:
                word.Documents.Add()
            try:
                word.WindowState = 0  # wdWindowStateNormal
                word.Activate()
                if word.ActiveWindow is not None:
                    word.ActiveWindow.Activate()
            except Exception:
                pass
            self._refresh_document()
            self._set_status("Hold the W button to preview Word; release it to minimize Word.")
        except Exception:
            self._word_preview_app = None
            self._set_status("Could not show Microsoft Word. Check that desktop Word is installed.")

    def _hide_word_preview(self):
        """Minimize Word immediately when the floating W button is released."""
        word = getattr(self, "_word_preview_app", None)
        if word is None:
            return
        try:
            word.WindowState = 2  # wdWindowStateMinimize
            try:
                if word.ActiveWindow is not None:
                    word.ActiveWindow.WindowState = 2
            except Exception:
                pass
        except Exception:
            pass

    def _refresh_document(self):
        try:
            name = self.writer.active_document_name()
            self.document_label.setText(f"Word document: {name}")
        except WordAutomationError:
            self.document_label.setText("Word document: none detected — open desktop Word and a document.")

    def _set_status(self, message):
        self.status_label.setText(f"Status: {message}")

    def changeEvent(self, event):
        # The quick bar is an independent top-level window, so it remains visible
        # when the main window is minimized. Re-raise it if it was already shown.
        super().changeEvent(event)
        if event.type() == event.Type.WindowStateChange and self.isMinimized():
            if self.quickbar.isVisible():
                self.quickbar.show()
                self.quickbar.raise_()

    def closeEvent(self, event):
        self.monitor.stop()
        self.quickbar.close()
        event.accept()
