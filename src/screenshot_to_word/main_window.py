"""Main GUI, floating heading bar, insertion controls, and preferences."""
from PySide6.QtCore import Qt, QSettings
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFormLayout, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QSpinBox,
    QVBoxLayout, QWidget,
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
        self.setMinimumWidth(360)
        self.heading_input = QLineEdit()
        self.heading_input.setPlaceholderText("Task heading (e.g. Task 3 — Row Reduction)")
        self.heading_button = QPushButton("Insert Heading")
        self.screenshot_button = QPushButton("Insert Screenshot")
        self.heading_button.clicked.connect(self._emit_heading)
        self.screenshot_button.clicked.connect(self._emit_screenshot)
        self.heading_input.returnPressed.connect(self._emit_heading)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(QLabel("Quick actions — stays on top"))
        layout.addWidget(self.heading_input)
        row = QHBoxLayout()
        row.addWidget(self.heading_button)
        row.addWidget(self.screenshot_button)
        layout.addLayout(row)
        self._heading_callback = None
        self._screenshot_callback = None

    def set_callbacks(self, heading_callback, screenshot_callback):
        self._heading_callback = heading_callback
        self._screenshot_callback = screenshot_callback

    def _emit_heading(self):
        if self._heading_callback:
            self._heading_callback(self.heading_input.text())

    def _emit_screenshot(self):
        if self._screenshot_callback:
            self._screenshot_callback()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ScreenshotToWord", "ScreenshotToWord")
        self.setWindowTitle("Screenshot to Word")
        self.setMinimumSize(500, 460)
        self.writer = WordWriter()
        self.monitor = ClipboardMonitor(int(self.settings.value("monitor_interval_ms", 300)), self)
        self.monitor.image_detected.connect(self._insert_image)
        self.monitor.status_changed.connect(self._set_status)
        self._auto_enabled = True

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
        self.check_button = QPushButton("Check Word")
        self.check_button.clicked.connect(self._refresh_document)

        controls = QHBoxLayout()
        controls.addWidget(self.start_button)
        controls.addWidget(self.pause_button)
        controls.addWidget(self.check_button)
        manual_row = QHBoxLayout()
        manual_row.addWidget(self.manual_button)
        manual_row.addWidget(self.heading_button)
        tools_row = QHBoxLayout()
        tools_row.addWidget(self.quickbar_button)
        tools_row.addWidget(self.settings_button)

        note = QLabel("Tip: open the intended Word document first. Screenshots are inserted locally; save the document in Word when ready.")
        note.setWordWrap(True)
        note.setObjectName("note")

        layout = QVBoxLayout()
        layout.setSpacing(12)
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
        self.setStyleSheet("""
            QMainWindow { background: #f7f8fa; color: #172033; }
            QWidget { font-size: 10pt; color: #172033; }
            QLabel { color: #30394a; }
            QCheckBox { color: #30394a; }
            QLabel#appTitle { font-size: 22px; font-weight: 700; color: #172033; }
            QLabel#note { color: #5b6474; }
            QPushButton { padding: 8px 10px; border: 1px solid #d4d9e2; border-radius: 6px; background: white; }
            QPushButton:hover { background: #eef3fb; }
            QPushButton:disabled { color: #9aa2af; background: #f0f2f5; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { color: #172033; background: #ffffff; padding: 6px; border: 1px solid #d4d9e2; border-radius: 5px; }
            QCheckBox { padding: 4px 0; }
        """)
        self.quickbar = FloatingHeadingBar(self)
        self.quickbar.set_callbacks(self._insert_heading, self._manual_insert)
        self._refresh_document()
        self._update_auto_controls()

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

    def _refresh_document(self):
        try:
            name = self.writer.active_document_name()
            self.document_label.setText(f"Word document: {name}")
        except WordAutomationError:
            self.document_label.setText("Word document: none detected — open desktop Word and a document.")

    def _set_status(self, message):
        self.status_label.setText(f"Status: {message}")

    def closeEvent(self, event):
        self.monitor.stop()
        self.quickbar.close()
        event.accept()
