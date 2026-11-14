"""Main GUI, floating heading bar, insertion controls, and preferences."""
from PySide6.QtCore import Qt, QSettings, QUrl
from PySide6.QtGui import QDesktopServices, QMovie, QPixmap, QIcon
from pathlib import Path

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
        self.setMinimumWidth(410)
        self.setStyleSheet("""QDialog { background:#f4f0e7; color:#28332f; } QLabel { color:#45544c; } QPushButton { background:#456b60; color:#fffdf8; border:0; border-radius:7px; padding:8px 12px; }""")
        form = QFormLayout()

        self.interval = QSpinBox()
        self.interval.setRange(100, 5000)
        self.interval.setSingleStep(100)
        self.interval.setSuffix(" ms")
        self.interval.setValue(int(settings.value("monitor_interval_ms", 300)))

        self.width = QDoubleSpinBox()
        self.width.setRange(2.0, 12.0)
        self.width.setSingleStep(0.25)
        self.width.setSuffix(" inches")
        self.width.setValue(float(settings.value("image_width_inches", 6.3)))
        self.width.setToolTip("Maximum width. The image will keep its original proportions.")

        self.height = QDoubleSpinBox()
        self.height.setRange(2.0, 16.0)
        self.height.setSingleStep(0.25)
        self.height.setSuffix(" inches")
        self.height.setValue(float(settings.value("image_height_inches", 9.0)))
        self.height.setToolTip("Maximum height. The image will keep its original proportions.")

        self.alignment = QComboBox()
        self.alignment.addItem("Left", "left")
        self.alignment.addItem("Center", "center")
        self.alignment.addItem("Right", "right")
        idx = self.alignment.findData(settings.value("image_alignment", "center"))
        self.alignment.setCurrentIndex(max(0, idx))

        self.mode = QComboBox()
        self.mode.addItem("Ask before inserting when an error occurs", "warn")
        self.mode.addItem("Show errors in status only", "quiet")
        idx = self.mode.findData(settings.value("error_mode", "warn"))
        if idx >= 0:
            self.mode.setCurrentIndex(idx)

        form.addRow("Clipboard check interval:", self.interval)
        form.addRow("Maximum screenshot width:", self.width)
        form.addRow("Maximum screenshot height:", self.height)
        form.addRow("Screenshot alignment:", self.alignment)
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
        self.settings.setValue("image_height_inches", self.height.value())
        self.settings.setValue("image_alignment", self.alignment.currentData())
        self.settings.setValue("error_mode", self.mode.currentData())


class FloatingHeadingBar(QDialog):
    """Small always-on-top companion bar for task headings and quick insertion."""
    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Tool | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("Screenshot to Word — Quick Bar")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setMinimumWidth(300)
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
        self.setStyleSheet("""
            QDialog { background: #f4efe5; color: #28332f; }
            QLabel { color: #28332f; font-weight: 700; }
            QLineEdit { color: #28332f; background: #fffdf8; border: 1px solid #c8c3b6; border-radius: 7px; padding: 8px; }
            QPushButton { color: #fffdf8; background: #42665d; border: 1px solid #36574f; border-radius: 7px; padding: 8px 10px; font-weight: 600; }
            QPushButton:hover { background: #52796e; }
        """)

    def set_callbacks(self, heading_callback, screenshot_callback):
        self._heading_callback = heading_callback
        self._screenshot_callback = screenshot_callback

    def _emit_heading(self):
        if self._heading_callback:
            self._heading_callback(self.heading_input.text())

    def _emit_screenshot(self):
        if self._screenshot_callback:
            self._screenshot_callback()


class AboutDialog(QDialog):
    """Compact, theme-matched About and open-source support dialog."""
    GITHUB_PROFILE = "https://github.com/ahmed-shehzad1"
    REPOSITORY_URL = "https://github.com/ahmed-shehzad1/Screenshot-to-word"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Screenshot to Word")
        self.setWindowIcon(QIcon(str(Path(__file__).resolve().parent / "assets" / "app_icon.png")))
        self.setMinimumSize(520, 440)
        self.resize(560, 480)
        self.setStyleSheet("""
            QDialog { background:#f4f0e7; color:#28332f; }
            QLabel { color:#58665d; font-size:10pt; }
            QLabel#eyebrow { color:#668477; font-size:8pt; font-weight:800; letter-spacing:1px; }
            QLabel#heroName { color:#26352f; font-size:20pt; font-weight:800; }
            QLabel#sectionTitle { color:#2e443a; font-size:12pt; font-weight:800; }
            QFrame#panel { background:#fffdf8; border:1px solid #ded8cc; border-radius:14px; }
            QFrame#heroPanel { background:#e9e2d5; border:1px solid #d8d0c2; border-radius:16px; }
            QPushButton { color:#3a4942; background:#fffdf8; border:1px solid #d7d1c6; border-radius:9px; padding:10px 13px; font-weight:700; }
            QPushButton:hover { background:#f0ece3; }
            QPushButton#primary { color:#fffdf8; background:#456b60; border-color:#456b60; }
            QPushButton#primary:hover { background:#557b6f; }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(22, 20, 22, 18)
        root.setSpacing(13)

        hero = QFrame()
        hero.setObjectName("heroPanel")
        hero_row = QHBoxLayout(hero)
        hero_row.setContentsMargins(18, 16, 18, 16)
        hero_row.setSpacing(15)
        logo = QLabel()
        logo.setFixedSize(88, 88)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pixmap = QPixmap(str(Path(__file__).resolve().parent / "assets" / "app_icon.png"))
        if not pixmap.isNull():
            logo.setPixmap(pixmap.scaled(88, 88, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        intro = QVBoxLayout()
        intro.setSpacing(4)
        eyebrow = QLabel("A SMALL TOOL FOR BETTER DOCUMENTATION")
        eyebrow.setObjectName("eyebrow")
        name = QLabel("Screenshot to Word")
        name.setObjectName("heroName")
        subtitle = QLabel("Built by Ahmed Shahzad · Air University, Islamabad")
        subtitle.setWordWrap(True)
        intro.addWidget(eyebrow)
        intro.addWidget(name)
        intro.addWidget(subtitle)
        intro.addStretch(1)
        hero_row.addWidget(logo)
        hero_row.addLayout(intro, 1)
        root.addWidget(hero)

        story_panel = QFrame()
        story_panel.setObjectName("panel")
        story_layout = QVBoxLayout(story_panel)
        story_layout.setContentsMargins(17, 15, 17, 15)
        story_layout.setSpacing(7)
        story_heading = QLabel("Why I built this")
        story_heading.setObjectName("sectionTitle")
        story = QLabel("I love coding things that can help people. Screenshot to Word makes it easier to capture, arrange, and document university lab work directly in Microsoft Word.")
        story.setWordWrap(True)
        story.setMinimumHeight(52)
        story_layout.addWidget(story_heading)
        story_layout.addWidget(story)
        root.addWidget(story_panel)

        support_panel = QFrame()
        support_panel.setObjectName("panel")
        support_layout = QVBoxLayout(support_panel)
        support_layout.setContentsMargins(17, 15, 17, 16)
        support_layout.setSpacing(8)
        support_heading = QLabel("Support the project")
        support_heading.setObjectName("sectionTitle")
        support_copy = QLabel("If this app is useful, you can help it reach more people by following the developer or starring the GitHub repository. Thank you for your support!")
        support_copy.setWordWrap(True)
        support_layout.addWidget(support_heading)
        support_layout.addWidget(support_copy)
        buttons = QHBoxLayout()
        buttons.setSpacing(9)
        profile_button = QPushButton("↗  Follow on GitHub")
        profile_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.GITHUB_PROFILE)))
        star_button = QPushButton("★  Star repository")
        star_button.setObjectName("primary")
        star_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.REPOSITORY_URL)))
        buttons.addWidget(profile_button)
        buttons.addWidget(star_button)
        support_layout.addLayout(buttons)
        root.addWidget(support_panel)

        footer = QHBoxLayout()
        footer.addStretch(1)
        close_button = QPushButton("Done")
        close_button.clicked.connect(self.accept)
        footer.addWidget(close_button)
        root.addLayout(footer)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ScreenshotToWord", "ScreenshotToWord")
        self.setWindowTitle("Screenshot to Word")
        self.app_icon_path = Path(__file__).resolve().parent / "assets" / "app_icon.png"
        self.setWindowIcon(QIcon(str(self.app_icon_path)))
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint)
        self.setMinimumSize(850, 620)
        self.resize(1040, 720)
        self._drag_position = None
        self.writer = WordWriter()
        self.monitor = ClipboardMonitor(int(self.settings.value("monitor_interval_ms", 300)), self)
        self.monitor.image_detected.connect(self._insert_image)
        self.monitor.status_changed.connect(self._set_status)
        self._auto_enabled = True

        titlebar = QWidget()
        titlebar.setObjectName("customTitleBar")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(18, 8, 12, 8)
        titlebar_layout.setSpacing(12)
        brand_mark = QLabel()
        brand_mark.setObjectName("brandMark")
        brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_mark.setFixedSize(42, 42)
        brand_pixmap = QPixmap(str(self.app_icon_path))
        if not brand_pixmap.isNull():
            brand_mark.setPixmap(brand_pixmap.scaled(42, 42, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        brand_name = QLabel("Screenshot to Word")
        brand_name.setObjectName("brandName")
        brand_hint = QLabel("A CALMER WAY TO DOCUMENT YOUR WORK")
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

        # Main split layout: illustrated study panel + practical workspace.
        body = QHBoxLayout()
        body.setContentsMargins(18, 16, 18, 18)
        body.setSpacing(18)

        art_panel = QFrame()
        art_panel.setObjectName("studyPanel")
        art_panel.setMinimumWidth(280)
        art_panel.setMaximumWidth(340)
        art_layout = QVBoxLayout(art_panel)
        art_layout.setContentsMargins(12, 12, 12, 16)
        art_layout.setSpacing(10)
        self.study_art = QLabel()
        self.study_art.setObjectName("studyArt")
        self.study_art.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.study_art.setMinimumHeight(240)
        movie_path = Path(__file__).resolve().parent / "assets" / "study_loop.gif"
        self.study_movie = QMovie(str(movie_path))
        if self.study_movie.isValid():
            self.study_art.setMovie(self.study_movie)
            self.study_movie.start()
        else:
            self.study_art.setText("YOUR STUDY WORKSPACE")
        art_layout.addWidget(self.study_art, 1)
        art_caption = QLabel("MAKE ROOM FOR\nBETTER WORK.")
        art_caption.setObjectName("artCaption")
        art_caption.setWordWrap(True)
        art_layout.addWidget(art_caption)
        art_copy = QLabel("Capture your progress, organize each task, and keep your lab documentation in one place.")
        art_copy.setObjectName("artCopy")
        art_copy.setWordWrap(True)
        art_layout.addWidget(art_copy)
        body.addWidget(art_panel)

        workspace = QWidget()
        workspace.setObjectName("workspace")
        content = QVBoxLayout(workspace)
        content.setContentsMargins(4, 2, 4, 2)
        content.setSpacing(13)
        title = QLabel("Your documentation workspace")
        title.setObjectName("appTitle")
        subtitle = QLabel("Send screenshots from your clipboard straight into your open Word document.")
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        self.document_label = QLabel("Word document: checking…")
        self.document_label.setObjectName("documentCard")
        self.document_label.setWordWrap(True)
        self.auto_checkbox = QCheckBox("Automatically insert each new clipboard screenshot")
        self.auto_checkbox.setChecked(True)
        self.auto_checkbox.toggled.connect(self._toggle_auto)
        self.start_button = QPushButton("Start automatic insertion")
        self.start_button.setObjectName("primaryButton")
        self.start_button.clicked.connect(self._start_stop_monitoring)
        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self._pause_resume)
        self.pause_button.setEnabled(False)
        controls = QHBoxLayout()
        controls.addWidget(self.start_button, 2)
        controls.addWidget(self.pause_button, 1)
        self.manual_button = QPushButton("Insert Screenshot Now")
        self.manual_button.setObjectName("primaryButton")
        self.manual_button.clicked.connect(self._manual_insert)
        self.heading_button = QPushButton("Insert Task Heading")
        self.heading_button.clicked.connect(self._insert_heading_from_main)
        manual_row = QHBoxLayout()
        manual_row.addWidget(self.manual_button)
        manual_row.addWidget(self.heading_button)
        self.quickbar_button = QPushButton("Show Floating Heading Bar")
        self.quickbar_button.clicked.connect(self._toggle_quickbar)
        self.settings_button = QPushButton("Settings")
        self.settings_button.clicked.connect(self._open_settings)
        self.about_button = QPushButton("About / Support")
        self.about_button.clicked.connect(self._show_about)
        tools_row = QHBoxLayout()
        tools_row.addWidget(self.quickbar_button)
        tools_row.addWidget(self.settings_button)
        tools_row.addWidget(self.about_button)
        self.clipboard_label = QLabel("Clipboard: not checked yet")
        self.clipboard_label.setObjectName("infoText")
        self.clipboard_label.setWordWrap(True)
        self.status_label = QLabel("Status: ready")
        self.status_label.setObjectName("statusCard")
        self.status_label.setWordWrap(True)
        note = QLabel("Before starting, open the Word document where you want your screenshots to appear. Remember to save it in Word.")
        note.setObjectName("note")
        note.setWordWrap(True)

        content.addWidget(title)
        content.addWidget(subtitle)
        content.addWidget(self.document_label)
        content.addWidget(self.auto_checkbox)
        content.addLayout(controls)
        content.addLayout(manual_row)
        content.addLayout(tools_row)
        content.addWidget(self.clipboard_label)
        content.addWidget(self.status_label)
        content.addWidget(note)
        content.addStretch(1)
        body.addWidget(workspace, 1)

        root = QVBoxLayout()
        root.setSpacing(0)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(titlebar)
        root.addLayout(body, 1)
        container = QWidget()
        container.setLayout(root)
        self.setCentralWidget(container)

        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f4f0e7; color: #28332f; font-size: 10pt; }
            QWidget#customTitleBar { background: #eee8dc; border-bottom: 1px solid #d8d0c2; }
            QLabel#brandMark { background: transparent; border-radius: 11px; }
            QLabel#brandName { color: #26352f; font-size: 10pt; font-weight: 800; }
            QLabel#brandHint { color: #7a8278; font-size: 7pt; font-weight: 700; }
            QPushButton#windowControl { background: transparent; border: none; border-radius: 7px; padding: 2px; color: #47564f; font-size: 12pt; }
            QPushButton#windowControl:hover { background: #d9d1c4; }
            QPushButton#closeControl { background: transparent; border: none; border-radius: 7px; padding: 2px; color: #47564f; font-size: 17pt; }
            QPushButton#closeControl:hover { background: #b94c42; color: white; }
            QFrame#studyPanel { background: #e8dfd0; border: 1px solid #d8cdbb; border-radius: 16px; }
            QLabel#studyArt { background: #d8c8b2; border-radius: 11px; }
            QLabel#artCaption { color: #31473f; font-size: 18pt; font-weight: 900; letter-spacing: 1px; }
            QLabel#artCopy { color: #68746b; font-size: 9pt; }
            QLabel#appTitle { color: #26352f; font-size: 23px; font-weight: 800; }
            QLabel#subtitle { color: #737c72; }
            QLabel#documentCard { background: #fffdf8; border: 1px solid #ded8cc; border-radius: 10px; padding: 13px; color: #42564d; font-weight: 600; }
            QLabel#statusCard { background: #e6eee8; border: 1px solid #cfddd2; border-radius: 9px; padding: 10px; color: #3d6254; }
            QLabel#infoText { color: #778078; }
            QLabel#note { color: #777d73; background: #eee8dc; border: 1px solid #ded6c8; border-radius: 9px; padding: 11px; }
            QPushButton { color: #3a4942; background: #fffdf8; border: 1px solid #d7d1c6; border-radius: 8px; padding: 11px 12px; font-weight: 650; }
            QPushButton:hover { background: #f0ece3; border-color: #aab8ad; }
            QPushButton:pressed { background: #e6e0d4; }
            QPushButton#primaryButton { color: #fffdf8; background: #456b60; border-color: #456b60; }
            QPushButton#primaryButton:hover { background: #557b6f; }
            QPushButton:disabled { color: #9a9e96; background: #ece8df; border-color: #e1dbd0; }
            QCheckBox { color: #45544c; spacing: 8px; padding: 8px 0; }
            QCheckBox::indicator { width: 17px; height: 17px; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { color: #28332f; background: #fffdf8; selection-background-color: #78998b; padding: 8px; border: 1px solid #d4cec2; border-radius: 7px; }
            QComboBox QAbstractItemView { background: #fffdf8; color: #28332f; selection-background-color: #dce8df; }
            QToolTip { color: #28332f; background: #fffdf8; border: 1px solid #cfc8bb; }
        """)

        self.quickbar = FloatingHeadingBar()
        self.quickbar.set_callbacks(self._insert_heading, self._manual_insert)
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
            # Allow dragging from any empty/title-bar label area, but not its buttons.
            probe = child
            in_titlebar = False
            while probe is not None:
                if probe.objectName() == "customTitleBar":
                    in_titlebar = True
                    break
                probe = probe.parentWidget()
            if in_titlebar and not isinstance(child, QPushButton):
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
            height = float(self.settings.value("image_height_inches", 9.0))
            alignment = str(self.settings.value("image_alignment", "center"))
            document_name = self.writer.append_image(
                image, max_width_inches=width,
                max_height_inches=height, alignment=alignment,
            )
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
