"""Small GUI for starting/stopping clipboard monitoring."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
    QPushButton, QVBoxLayout, QWidget,
)

from screenshot_to_word.clipboard_monitor import ClipboardMonitor
from screenshot_to_word.word_writer import WordAutomationError, WordWriter


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Screenshot to Word")
        self.setMinimumWidth(440)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, False)
        self.writer = WordWriter()
        self.monitor = ClipboardMonitor(parent=self)
        self.monitor.image_detected.connect(self._insert_image)
        self.monitor.status_changed.connect(self._set_status)

        title = QLabel("Screenshot to Word")
        title.setStyleSheet("font-size: 20px; font-weight: 600;")
        subtitle = QLabel("Automatically append clipboard screenshots to your open Word document.")
        subtitle.setWordWrap(True)

        self.document_label = QLabel("Word document: checking…")
        self.document_label.setWordWrap(True)
        self.status_label = QLabel("Status: paused")
        self.status_label.setWordWrap(True)

        self.start_button = QPushButton("Start monitoring")
        self.start_button.clicked.connect(self._toggle_monitoring)
        self.refresh_button = QPushButton("Check Word")
        self.refresh_button.clicked.connect(self._refresh_document)

        self.auto_save_note = QLabel(
            "Note: screenshots are inserted into the active Word document. "
            "Save your document in Word when you are ready."
        )
        self.auto_save_note.setWordWrap(True)
        self.auto_save_note.setStyleSheet("color: #666;")

        buttons = QHBoxLayout()
        buttons.addWidget(self.start_button)
        buttons.addWidget(self.refresh_button)
        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addSpacing(4)
        layout.addWidget(self.document_label)
        layout.addLayout(buttons)
        layout.addWidget(self.status_label)
        layout.addWidget(self.auto_save_note)
        layout.addStretch(1)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self._refresh_document()

    def _toggle_monitoring(self):
        if self.monitor.running:
            self.monitor.stop()
            self.start_button.setText("Start monitoring")
        else:
            try:
                self.writer.active_document_name()
            except WordAutomationError as exc:
                QMessageBox.warning(self, "Word not ready", str(exc))
                return
            self.monitor.start()
            self.start_button.setText("Stop monitoring")

    def _insert_image(self, image):
        try:
            document_name = self.writer.append_image(image)
            self._set_status(f"Inserted screenshot into {document_name}.")
            self._refresh_document()
        except WordAutomationError as exc:
            self._set_status(str(exc))
            QMessageBox.warning(self, "Screenshot not inserted", str(exc))

    def _refresh_document(self):
        try:
            name = self.writer.active_document_name()
            self.document_label.setText(f"Word document: {name}")
        except WordAutomationError:
            self.document_label.setText("Word document: none detected — open Word and a document.")

    def _set_status(self, message):
        self.status_label.setText(f"Status: {message}")

    def closeEvent(self, event):
        self.monitor.stop()
        event.accept()
