"""Detect new clipboard images using Windows' clipboard sequence counter."""
import ctypes
import hashlib
import time
from typing import Optional

from PIL import Image
from PySide6.QtCore import QObject, QTimer, Signal
from PySide6.QtGui import QGuiApplication


class ClipboardMonitor(QObject):
    image_detected = Signal(object)  # PIL.Image.Image
    status_changed = Signal(str)

    def __init__(self, interval_ms: int = 300, parent=None):
        super().__init__(parent)
        self._timer = QTimer(self)
        self._timer.setInterval(interval_ms)
        self._timer.timeout.connect(self._poll)
        self._last_sequence = self._sequence_number()
        self._last_hash: Optional[str] = None
        self._last_emitted_at = 0.0
        self._running = False

    @staticmethod
    def _sequence_number() -> int:
        try:
            return int(ctypes.windll.user32.GetClipboardSequenceNumber())
        except Exception:
            return -1

    def start(self) -> None:
        self._last_sequence = self._sequence_number()
        self._running = True
        self._timer.start()
        self.status_changed.emit("Monitoring clipboard for screenshots.")

    def stop(self) -> None:
        self._timer.stop()
        self._running = False
        self.status_changed.emit("Monitoring paused.")

    @property
    def running(self) -> bool:
        return self._running

    def _poll(self) -> None:
        sequence = self._sequence_number()
        if sequence < 0 or sequence == self._last_sequence:
            return
        self._last_sequence = sequence

        # Clipboard formats can be published a moment apart by Snipping Tool.
        clipboard = QGuiApplication.clipboard()
        qimage = clipboard.image()
        if qimage.isNull():
            return

        try:
            buffer = qimage.bits()
            buffer.setsize(qimage.sizeInBytes())
            image = Image.frombytes(
                "RGBA" if qimage.hasAlphaChannel() else "RGB",
                (qimage.width(), qimage.height()),
                bytes(buffer),
                "raw",
                "BGRA" if qimage.hasAlphaChannel() else "BGR",
                qimage.bytesPerLine(),
            ).convert("RGB")
        except Exception:
            # A PNG round-trip is slower but reliable across Qt image formats.
            from PySide6.QtCore import QBuffer, QByteArray, QIODevice
            byte_array = QByteArray()
            qbuffer = QBuffer(byte_array)
            qbuffer.open(QIODevice.OpenModeFlag.WriteOnly)
            qimage.save(qbuffer, "PNG")
            qbuffer.close()
            image = Image.open(__import__("io").BytesIO(bytes(byte_array))).convert("RGB")

        digest = hashlib.sha256(image.tobytes()).hexdigest()
        now = time.monotonic()
        # Windows can update the clipboard sequence for multiple formats for
        # one capture. Ignore the same image if it reappears immediately.
        if digest == self._last_hash and now - self._last_emitted_at < 1.5:
            return

        self._last_hash = digest
        self._last_emitted_at = now
        self.image_detected.emit(image)
