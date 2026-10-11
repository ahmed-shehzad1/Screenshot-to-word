"""Detect new clipboard images using the Windows clipboard sequence counter."""
import ctypes
import hashlib
import io
import time
from typing import Optional

from PIL import Image
from PySide6.QtCore import QObject, QTimer, Signal, QBuffer, QByteArray, QIODevice
from PySide6.QtGui import QGuiApplication


def current_clipboard_image():
    """Return a copy of the current clipboard image as a PIL RGB image, or None."""
    qimage = QGuiApplication.clipboard().image()
    if qimage.isNull():
        return None
    try:
        byte_array = QByteArray()
        qbuffer = QBuffer(byte_array)
        qbuffer.open(QIODevice.OpenModeFlag.WriteOnly)
        qimage.save(qbuffer, "PNG")
        qbuffer.close()
        return Image.open(io.BytesIO(bytes(byte_array))).convert("RGB").copy()
    except Exception:
        return None


class ClipboardMonitor(QObject):
    image_detected = Signal(object)  # PIL.Image.Image
    status_changed = Signal(str)

    def __init__(self, interval_ms: int = 300, parent=None):
        super().__init__(parent)
        self._timer = QTimer(self)
        self._timer.setInterval(int(interval_ms))
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

    def set_interval(self, interval_ms: int) -> None:
        self._timer.setInterval(max(100, int(interval_ms)))

    def start(self) -> None:
        self._last_sequence = self._sequence_number()
        self._running = True
        self._timer.start()
        self.status_changed.emit("Automatic monitoring is running.")

    def stop(self) -> None:
        self._timer.stop()
        self._running = False
        self.status_changed.emit("Automatic monitoring is stopped.")

    @property
    def running(self) -> bool:
        return self._running

    def _poll(self) -> None:
        sequence = self._sequence_number()
        if sequence < 0 or sequence == self._last_sequence:
            return
        self._last_sequence = sequence
        image = current_clipboard_image()
        if image is None:
            return
        digest = hashlib.sha256(image.tobytes()).hexdigest()
        now = time.monotonic()
        # Windows can publish several clipboard formats for a single capture.
        if digest == self._last_hash and now - self._last_emitted_at < 1.5:
            return
        self._last_hash = digest
        self._last_emitted_at = now
        self.image_detected.emit(image)
