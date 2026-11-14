"""Main GUI, floating heading bar, insertion controls, and preferences."""
import ctypes

from PySide6.QtCore import Qt, QSettings, QUrl
from PySide6.QtGui import QDesktopServices
import win32com.client

from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFormLayout, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton, QSpinBox,
    QVBoxLayout, QWidget, QFrame, QListWidget, QListWidgetItem,
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
            QDialog { background: #F5F0E7; color: #211E1A; }
            QLabel { color: #33271F; font-weight: 700; }
            QLineEdit { color: #29241F; background: #FFFCF6; border: 1px solid #DCCDBB; border-radius: 7px; padding: 8px; }
            QPushButton { color: #33271F; background: #FFFCF6; border: 1px solid #DCCDBB; border-radius: 7px; padding: 8px 10px; font-weight: 700; }
            QPushButton:hover { background: #F1E5D6; }
            QSpinBox, QDoubleSpinBox, QComboBox { color: #29241F; background: #FFFCF6; border: 1px solid #DCCDBB; border-radius: 7px; padding: 7px; }
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
    """Warm, editorial About dialog with a visible heart motif."""
    GITHUB_PROFILE = "https://github.com/ahmed-shehzad1"
    REPOSITORY_URL = "https://github.com/ahmed-shehzad1/Screenshot-to-word"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Screenshot to Word")
        self.setMinimumWidth(500)
        self.setStyleSheet("""
            QDialog { background: #F5F0E7; color: #211E1A; }
            QLabel { color: #5F584F; font-size: 10pt; }
            QLabel#eyebrow { color: #C85A2B; font-size: 9pt; font-weight: 800; letter-spacing: 2px; }
            QLabel#heroName { color: #211E1A; font-size: 23pt; font-weight: 900; }
            QLabel#heart { color: #D85C32; background: #FBE1D4; border: 1px solid #EFC4B1; border-radius: 22px; font-size: 30pt; font-weight: 700; }
            QFrame#panel { background: #FFFCF6; border: 1px solid #E4DACE; border-radius: 14px; }
            QPushButton { color: #33271F; background: #EFE6D9; border: 1px solid #DCCDBB; border-radius: 9px; padding: 11px 14px; font-weight: 700; }
            QPushButton:hover { background: #E7D8C5; }
            QPushButton#primary { color: #FFF9F2; background: #C85A2B; border-color: #C85A2B; }
            QPushButton#primary:hover { background: #A94720; }
            QPushButton#quiet { background: transparent; border-color: transparent; }
        """)
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 20)
        root.setSpacing(14)

        eyebrow = QLabel("A SMALL TOOL, MADE WITH CARE")
        eyebrow.setObjectName("eyebrow")
        root.addWidget(eyebrow)

        hero = QFrame()
        hero.setObjectName("panel")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(18, 20, 18, 20)
        hero_layout.setSpacing(16)
        heart = QLabel("♥")
        heart.setObjectName("heart")
        heart.setAlignment(Qt.AlignmentFlag.AlignCenter)
        heart.setFixedSize(76, 76)
        intro = QVBoxLayout()
        intro.setSpacing(4)
        name = QLabel("Screenshot to Word")
        name.setObjectName("heroName")
        subtitle = QLabel("Built by Ahmed Shahzad")
        tagline = QLabel("A simpler place for the screenshots that matter.")
        tagline.setWordWrap(True)
        intro.addWidget(name)
        intro.addWidget(subtitle)
        intro.addSpacing(3)
        intro.addWidget(tagline)
        hero_layout.addWidget(heart, 0, Qt.AlignmentFlag.AlignTop)
        hero_layout.addLayout(intro, 1)
        root.addWidget(hero)

        story_panel = QFrame()
        story_panel.setObjectName("panel")
        story_layout = QVBoxLayout(story_panel)
        story_layout.setContentsMargins(18, 16, 18, 17)
        story_layout.setSpacing(7)
        heading = QLabel("Why I built this")
        heading.setStyleSheet("color:#211E1A; font-size:13pt; font-weight:850;")
        story = QLabel("I enjoy building useful little tools. Screenshot to Word makes it easier to collect screenshots and keep them organized in a Microsoft Word document, without interrupting your workflow.")
        story.setWordWrap(True)
        story.setMinimumHeight(52)
        story_layout.addWidget(heading)
        story_layout.addWidget(story)
        root.addWidget(story_panel)

        support_panel = QFrame()
        support_panel.setObjectName("panel")
        support_layout = QVBoxLayout(support_panel)
        support_layout.setContentsMargins(18, 16, 18, 18)
        support_layout.setSpacing(8)
        support_heading = QLabel("If this tool helped you")
        support_heading.setStyleSheet("color:#211E1A; font-size:13pt; font-weight:850;")
        support_copy = QLabel("A follow or a repository star is a small gesture that helps this project reach more people. Thanks for giving it a try.")
        support_copy.setWordWrap(True)
        support_layout.addWidget(support_heading)
        support_layout.addWidget(support_copy)
        buttons = QHBoxLayout()
        profile_button = QPushButton("↗  Follow on GitHub")
        profile_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.GITHUB_PROFILE)))
        star_button = QPushButton("♥  Star the project")
        star_button.setObjectName("primary")
        star_button.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(self.REPOSITORY_URL)))
        buttons.addWidget(profile_button)
        buttons.addWidget(star_button)
        support_layout.addLayout(buttons)
        root.addWidget(support_panel)

        footer = QHBoxLayout()
        footer.addStretch(1)
        close_button = QPushButton("Close")
        close_button.setObjectName("quiet")
        close_button.clicked.connect(self.accept)
        footer.addWidget(close_button)
        root.addLayout(footer)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("ScreenshotToWord", "ScreenshotToWord")
        self.setWindowTitle("Screenshot to Word")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint)
        self.setMinimumSize(660, 700)
        self.resize(820, 860)
        self._drag_position = None
<<<<<<< HEAD
        self._word_preview_app = None
        self._word_preview_hwnd = None
=======
        self._activity_items = []
>>>>>>> 98c23b237bb1aeacf08cf46546e6e12c5dbd9c9d
        self.writer = WordWriter()
        self.monitor = ClipboardMonitor(int(self.settings.value("monitor_interval_ms", 300)), self)
        self.monitor.image_detected.connect(self._insert_image)
        self.monitor.status_changed.connect(self._set_status)
        self._auto_enabled = True

        titlebar = QWidget()
        titlebar.setObjectName("customTitleBar")
        titlebar_layout = QHBoxLayout(titlebar)
        titlebar_layout.setContentsMargins(18, 10, 10, 10)
        titlebar_layout.setSpacing(10)
        brand_mark = QLabel("S/")
        brand_mark.setObjectName("brandMark")
        brand_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_mark.setFixedSize(38, 38)
        brand_name = QLabel("SCREENSHOT TO WORD")
        brand_name.setObjectName("brandName")
        brand_hint = QLabel("A DESKTOP TOOL FOR YOUR WORKFLOW")
        brand_hint.setObjectName("brandHint")
        brand_text = QVBoxLayout()
        brand_text.setSpacing(2)
        brand_text.addWidget(brand_name)
        brand_text.addWidget(brand_hint)
        titlebar_layout.addWidget(brand_mark)
        titlebar_layout.addLayout(brand_text)
        titlebar_layout.addStretch(1)
        self.minimize_button = QPushButton("—")
        self.minimize_button.setObjectName("windowControl")
        self.minimize_button.setFixedSize(34, 30)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.maximize_button = QPushButton("□")
        self.maximize_button.setObjectName("windowControl")
        self.maximize_button.setFixedSize(34, 30)
        self.maximize_button.clicked.connect(self._toggle_maximize)
        close_button = QPushButton("×")
        close_button.setObjectName("closeControl")
        close_button.setFixedSize(34, 30)
        close_button.clicked.connect(self.close)
        titlebar_layout.addWidget(self.minimize_button)
        titlebar_layout.addWidget(self.maximize_button)
        titlebar_layout.addWidget(close_button)

        eyebrow = QLabel("CAPTURE  /  COLLECT  /  CONTINUE")
        eyebrow.setObjectName("eyebrow")
        title = QLabel("Keep the good\n bits in one place.")
        title.setObjectName("appTitle")
        subtitle = QLabel("Send clipboard screenshots straight into your open Word document.")
        subtitle.setObjectName("appSubtitle")
        subtitle.setWordWrap(True)

        self.document_label = QLabel("Checking for an open Word document…")
        self.document_label.setObjectName("documentValue")
        self.document_state = QLabel("WORD DESTINATION")
        self.document_state.setObjectName("cardEyebrow")
        document_card = QFrame()
        document_card.setObjectName("paperCard")
        document_layout = QVBoxLayout(document_card)
        document_layout.setContentsMargins(18, 15, 18, 16)
        document_layout.setSpacing(6)
        document_layout.addWidget(self.document_state)
        document_layout.addWidget(self.document_label)
        self.auto_checkbox = QCheckBox("Automatically insert each new clipboard screenshot")
        self.auto_checkbox.setChecked(True)
        self.auto_checkbox.toggled.connect(self._toggle_auto)
        document_layout.addWidget(self.auto_checkbox)

        monitor_card = QFrame()
        monitor_card.setObjectName("monitorCard")
        monitor_layout = QVBoxLayout(monitor_card)
        monitor_layout.setContentsMargins(18, 16, 18, 17)
        monitor_layout.setSpacing(10)
        monitor_top = QHBoxLayout()
        monitor_title_box = QVBoxLayout()
        monitor_title_box.setSpacing(3)
        monitor_eyebrow = QLabel("AUTOMATION")
        monitor_eyebrow.setObjectName("cardEyebrow")
        monitor_title = QLabel("Clipboard monitor")
        monitor_title.setObjectName("sectionTitle")
        monitor_title_box.addWidget(monitor_eyebrow)
        monitor_title_box.addWidget(monitor_title)
        self.monitor_badge = QLabel("●  READY")
        self.monitor_badge.setObjectName("statusBadge")
        self.monitor_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        monitor_top.addLayout(monitor_title_box, 1)
        monitor_top.addWidget(self.monitor_badge, 0, Qt.AlignmentFlag.AlignTop)
        monitor_layout.addLayout(monitor_top)
        monitor_copy = QLabel("New screenshots are added when they appear on your clipboard.")
        monitor_copy.setObjectName("mutedCopy")
        monitor_copy.setWordWrap(True)
        monitor_layout.addWidget(monitor_copy)
        controls = QHBoxLayout()
        controls.setSpacing(8)
        self.start_button = QPushButton("Start monitoring  →")
        self.start_button.setObjectName("primaryButton")
        self.start_button.clicked.connect(self._start_stop_monitoring)
        self.pause_button = QPushButton("Pause")
        self.pause_button.setObjectName("secondaryButton")
        self.pause_button.clicked.connect(self._pause_resume)
        self.pause_button.setEnabled(False)
        controls.addWidget(self.start_button, 2)
        controls.addWidget(self.pause_button, 1)
        monitor_layout.addLayout(controls)

        action_heading = QLabel("QUICK ACTIONS")
        action_heading.setObjectName("cardEyebrow")
        self.manual_button = QPushButton("＋  Insert screenshot now")
        self.manual_button.setObjectName("actionButton")
        self.manual_button.clicked.connect(self._manual_insert)
        self.heading_button = QPushButton("Tt  Insert task heading")
        self.heading_button.setObjectName("actionButton")
        self.heading_button.clicked.connect(self._insert_heading_from_main)
        self.quickbar_button = QPushButton("↗  Floating heading bar")
        self.quickbar_button.setObjectName("actionButton")
        self.quickbar_button.clicked.connect(self._toggle_quickbar)
        self.settings_button = QPushButton("⚙  Settings")
        self.settings_button.setObjectName("actionButton")
        self.settings_button.clicked.connect(self._open_settings)
        self.about_button = QPushButton("♥  About & support")
        self.about_button.setObjectName("actionButton")
        self.about_button.clicked.connect(self._show_about)
        action_grid = QHBoxLayout()
        action_grid.setSpacing(8)
        action_left = QVBoxLayout()
        action_left.setSpacing(8)
        action_right = QVBoxLayout()
        action_right.setSpacing(8)
        action_left.addWidget(self.manual_button)
        action_left.addWidget(self.quickbar_button)
        action_right.addWidget(self.heading_button)
        action_right.addWidget(self.settings_button)
        action_grid.addLayout(action_left, 1)
        action_grid.addLayout(action_right, 1)

        activity_card = QFrame()
        activity_card.setObjectName("paperCard")
        activity_layout = QVBoxLayout(activity_card)
        activity_layout.setContentsMargins(18, 14, 18, 12)
        activity_layout.setSpacing(8)
        activity_top = QHBoxLayout()
        activity_title = QLabel("Recent activity")
        activity_title.setObjectName("sectionTitle")
        activity_caption = QLabel("LATEST FIRST")
        activity_caption.setObjectName("cardEyebrow")
        activity_top.addWidget(activity_title)
        activity_top.addStretch(1)
        activity_top.addWidget(activity_caption)
        activity_layout.addLayout(activity_top)
        self.activity_list = QListWidget()
        self.activity_list.setObjectName("activityList")
        self.activity_list.setMinimumHeight(96)
        self.activity_list.setMaximumHeight(140)
        self.activity_list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.activity_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        activity_layout.addWidget(self.activity_list)
        self.status_label = QLabel("Ready when you are.")
        self.status_label.setObjectName("statusLine")
        self.status_label.setWordWrap(True)
        self.clipboard_label = QLabel("Clipboard: waiting for an image")
        self.clipboard_label.setObjectName("mutedCopy")
        self.clipboard_label.setWordWrap(True)
        activity_layout.addWidget(self.status_label)
        activity_layout.addWidget(self.clipboard_label)

        note = QLabel("Tip  /  Open the document you want to use before starting. Word remains in control of saving your work.")
        note.setObjectName("footerNote")
        note.setWordWrap(True)

        layout = QVBoxLayout()
        layout.setSpacing(12)
        layout.setContentsMargins(20, 0, 20, 18)
        layout.addWidget(titlebar)
        layout.addWidget(eyebrow)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(document_card)
        layout.addWidget(monitor_card)
        layout.addWidget(action_heading)
        layout.addLayout(action_grid)
        layout.addWidget(self.about_button)
        layout.addWidget(activity_card)
        layout.addWidget(note)
        container = QWidget()
        container.setObjectName("mainSurface")
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.setStyleSheet("""
            QMainWindow, QWidget#mainSurface { background: #F5F0E7; color: #211E1A; font-size: 10pt; }
            QWidget#customTitleBar { background: #F5F0E7; border-bottom: 1px solid #E4DACE; }
            QLabel#brandMark { background: #C85A2B; color: #FFF9F2; border-radius: 11px; font-size: 16pt; font-weight: 900; }
            QLabel#brandName { color: #211E1A; font-size: 9pt; font-weight: 900; letter-spacing: 1px; }
            QLabel#brandHint { color: #8D8174; font-size: 7pt; font-weight: 800; }
            QPushButton#windowControl { background: transparent; border: none; border-radius: 6px; padding: 2px; color: #62594F; font-size: 12pt; }
            QPushButton#windowControl:hover { background: #E9DFD2; }
            QPushButton#closeControl { background: transparent; border: none; border-radius: 6px; color: #62594F; font-size: 17pt; }
            QPushButton#closeControl:hover { background: #C85A2B; color: white; }
            QLabel#eyebrow { color: #C85A2B; font-size: 8pt; font-weight: 900; letter-spacing: 2px; padding-top: 5px; }
            QLabel#appTitle { color: #211E1A; font-family: Georgia, 'Times New Roman'; font-size: 27pt; font-weight: 700; padding: 0; }
            QLabel#appSubtitle { color: #6F665C; font-size: 10pt; padding-bottom: 3px; }
            QFrame#paperCard { background: #FFFCF6; border: 1px solid #E4DACE; border-radius: 13px; }
            QFrame#monitorCard { background: #EDE3D5; border: 1px solid #E0D1BF; border-radius: 13px; }
            QLabel#cardEyebrow { color: #A45A39; font-size: 8pt; font-weight: 900; letter-spacing: 1.5px; }
            QLabel#documentValue { color: #29241F; font-size: 11pt; font-weight: 750; }
            QLabel#sectionTitle { color: #211E1A; font-size: 14pt; font-weight: 850; }
            QLabel#mutedCopy { color: #766B60; font-size: 9pt; }
            QLabel#statusBadge { background: #F9F0E3; color: #A45A39; border: 1px solid #DDC4AC; border-radius: 9px; padding: 6px 9px; font-size: 8pt; font-weight: 900; }
            QLabel#statusLine { color: #33271F; font-size: 9pt; font-weight: 700; }
            QLabel#footerNote { color: #8B7D6E; font-size: 8pt; padding: 2px 1px 0 1px; }
            QPushButton { color: #33271F; background: #FFFCF6; border: 1px solid #DCCDBB; border-radius: 9px; padding: 10px 12px; font-weight: 750; text-align: left; }
            QPushButton:hover { background: #F1E5D6; border-color: #C8AD93; }
            QPushButton:pressed { background: #E7D4C0; }
            QPushButton:disabled { color: #A99D90; background: #E9E0D5; border-color: #DDD2C5; }
            QPushButton#primaryButton { color: #FFF9F2; background: #C85A2B; border-color: #C85A2B; text-align: center; }
            QPushButton#primaryButton:hover { background: #A94720; border-color: #A94720; }
            QPushButton#secondaryButton { background: #F8F0E5; text-align: center; }
            QPushButton#actionButton { min-height: 18px; }
            QCheckBox { color: #5F584F; spacing: 8px; padding-top: 7px; font-size: 9pt; }
            QCheckBox::indicator { width: 16px; height: 16px; }
            QCheckBox::indicator:unchecked { background: #FFFCF6; border: 1px solid #BBAA98; border-radius: 4px; }
            QCheckBox::indicator:checked { background: #C85A2B; border: 1px solid #C85A2B; border-radius: 4px; }
            QListWidget#activityList { background: transparent; color: #655A4F; border: none; outline: none; padding: 0; }
            QListWidget#activityList::item { padding: 5px 2px; border-bottom: 1px solid #EEE5D9; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { color: #29241F; background: #FFFCF6; selection-background-color: #C85A2B; padding: 8px; border: 1px solid #DCCDBB; border-radius: 7px; }
            QComboBox QAbstractItemView { background: #FFFCF6; color: #29241F; selection-background-color: #F1E5D6; }
            QToolTip { color: #FFF9F2; background: #33271F; border: 1px solid #C85A2B; padding: 5px; }
        """)
        self.quickbar = FloatingHeadingBar()
        self.quickbar.set_callbacks(
            self._insert_heading,
            self._manual_insert,
            self._show_word_preview,
            self._hide_word_preview,
        )
        self._refresh_document()
        self._set_status("Ready. Open Word, then start monitoring.")
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
        self.pause_button.setEnabled(self._auto_enabled)
        self.pause_button.setText("Pause" if self.monitor.running else "Resume")
        self.start_button.setText("Stop monitoring" if self.monitor.running else "Start monitoring  →")
        self.monitor_badge.setText("●  MONITORING" if self.monitor.running else "●  READY")
        self.monitor_badge.setStyleSheet(
            "background:#E0EBD9;color:#42633C;border:1px solid #B8CEAD;border-radius:9px;padding:6px 9px;font-size:8pt;font-weight:900;"
            if self.monitor.running else ""
        )

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
        """Temporarily show the already-open Word document while W is held."""
        # Do not Dispatch a new Word process or create a blank document here.
        # The preview is specifically for the document the user already has open.
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            if word.Documents.Count == 0:
                self._word_preview_app = None
                self._word_preview_hwnd = None
                self._set_status("No Word document is open. Open your document first, then hold W.")
                return

            window = word.ActiveWindow
            if window is None:
                self._set_status("Word is open, but its document window could not be accessed.")
                return

            hwnd = int(window.Hwnd)
            if hwnd <= 0:
                raise RuntimeError("Word returned an invalid window handle.")

            self._word_preview_app = word
            self._word_preview_hwnd = hwnd
            word.Visible = True
            try:
                window.WindowState = 0  # wdWindowStateNormal
            except Exception:
                pass

            user32 = ctypes.windll.user32
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            try:
                window.Activate()
            except Exception:
                pass

            self._refresh_document()
            self._set_status("Word preview is visible. Release W to minimize it again.")
        except Exception:
            self._word_preview_app = None
            self._word_preview_hwnd = None
            self._set_status("Could not access an open Word document. Open Word and your document first.")

    def _hide_word_preview(self):
        """Minimize the same Word window that was restored when W was pressed."""
        hwnd = self._word_preview_hwnd
        word = self._word_preview_app
        # Clear first so repeated release/focus events cannot act on a stale handle.
        self._word_preview_hwnd = None
        self._word_preview_app = None
        if not hwnd:
            return

        try:
            ctypes.windll.user32.ShowWindow(int(hwnd), 6)  # SW_MINIMIZE
        except Exception:
            # COM fallback for systems where the Win32 window call fails.
            if word is not None:
                try:
                    word.WindowState = 2  # wdWindowStateMinimize
                except Exception:
                    pass

    def _refresh_document(self):
        try:
            name = self.writer.active_document_name()
            self.document_label.setText(f"Word document: {name}")
        except WordAutomationError:
            self.document_label.setText("Word document: none detected — open desktop Word and a document.")

    def _set_status(self, message):
        self.status_label.setText(message)
        stamp = __import__("datetime").datetime.now().strftime("%H:%M:%S")
        self._activity_items.insert(0, f"{stamp}   ·   {message}")
        self._activity_items = self._activity_items[:5]
        self.activity_list.clear()
        for item_text in self._activity_items:
            self.activity_list.addItem(QListWidgetItem(item_text))

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
