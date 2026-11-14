"""Debounced Windows clipboard image monitoring.

Snipping Tool can publish the same capture to the clipboard in several
successive updates. Wait for the clipboard sequence number to settle before
reading it, then suppress repeated delivery of the same image.
"""
import ctypes
import hashlib
import io
import time
from typing import Optional

from PIL import Image
from PySide6.QtCore import QObject, QTimer, Signal, QBuffer, QByteArray, QIODevice
from PySide6.QtGui import QGuiApplication


def _read_clipboard_image():
    """Return the current clipboard image as a detached RGB PIL image."""
    clipboard = QGuiApplication.clipboard()
    qimage = clipboard.image()
    if qimage.isNull():
        return None

    byte_array = QByteArray()
    buffer = QBuffer(byte_array)
    if not buffer.open(QIODevice.OpenModeFlag.WriteOnly):
        return None
    try:
        if not qimage.save(buffer, "PNG"):
            return None
    finally:
        buffer.close()

    try:
        return Image.open(io.BytesIO(bytes(byte_array))).convert("RGB").copy()
    except Exception:
        return None


def current_clipboard_image():
    """Read an image currently on the clipboard (used by manual insertion)."""
    return _read_clipboard_image()


class ClipboardMonitor(QObject):
    image_detected = Signal(object)  # PIL.Image.Image
    status_changed = Signal(str)

    def __init__(self, interval_ms: int = 300, parent=None):
        super().__init__(parent)
        self._timer = QTimer(self)
        self._timer.setInterval(max(100, int(interval_ms)))
        self._timer.timeout.connect(self._poll)

        self._last_sequence = self._sequence_number()
        self._pending_sequence = None
        self._sequence_stable_since = 0.0
        self._last_hash: Optional[str] = None
        self._last_emitted_at = 0.0
        self._running = False

        # Snipping Tool may publish several clipboard formats/updates for one
        # capture. Wait until updates stop, then ignore a repeated image briefly.
        self._settle_seconds = 0.45
        self._duplicate_window_seconds = 8.0

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
        self._pending_sequence = None
        self._sequence_stable_since = 0.0
        self._running = True
        self._timer.start()
        self.status_changed.emit("Monitoring clipboard for screenshots.")

    def stop(self) -> None:
        self._timer.stop()
        self._pending_sequence = None
        self._running = False
        self.status_changed.emit("Monitoring paused.")

    @property
    def running(self) -> bool:
        return self._running

    def _poll(self) -> None:
        if not self._running:
            return

        now = time.monotonic()
        sequence = self._sequence_number()
        if sequence < 0:
            return

        if sequence != self._last_sequence:
            self._last_sequence = sequence
            self._pending_sequence = sequence
            self._sequence_stable_since = now
            return

        if self._pending_sequence is None:
            return

        # Any later clipboard update restarts the settling period. Only read
        # once Windows has stopped changing the clipboard for this interval.
        if now - self._sequence_stable_since < self._settle_seconds:
            return

        self._pending_sequence = None
        image = _read_clipboard_image()
        if image is None:
            return

        digest = hashlib.sha256(image.tobytes()).hexdigest()
        if (
            digest == self._last_hash
            and now - self._last_emitted_at < self._duplicate_window_seconds
        ):
            return

        self._last_hash = digest
        self._last_emitted_at = now
        self.image_detected.emit(image)
