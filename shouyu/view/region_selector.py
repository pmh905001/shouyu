"""Fullscreen drag-to-select overlay used by the OCR capture hotkey.

See docs/screenshot-ocr-design.md §2.2. Captures the whole virtual desktop
once (so multi-monitor setups work and the crop is pixel-exact - no second
grab call, so the overlay window itself can never leak into the result),
shows it dimmed under a "spotlight" that follows the drag, and returns the
cropped region as a PIL Image on release (or None on Esc / too-small drag).

Must be constructed and `.capture()`-ed on the Qt GUI thread (same
constraint as every other dialog in this app) - see
QtApp.request_ocr_capture / _QtBridge._on_ocr_capture for how the hotkey
thread hands off to it.
"""
from __future__ import annotations

import logging
from typing import Optional

from PIL import Image, ImageGrab
from PySide6.QtCore import QRect, QRectF, Qt
from PySide6.QtGui import QColor, QGuiApplication, QImage, QPainter, QPainterPath, QPen, QPixmap
from PySide6.QtWidgets import QDialog

_SELECTION_BORDER_COLOR = QColor(0, 153, 255)
_DIM_COLOR = QColor(0, 0, 0, 120)
_MIN_SELECTION_SIZE = 4


class RegionSelector(QDialog):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlags(
            Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.BypassWindowManagerHint
        )
        self.setCursor(Qt.CrossCursor)

        self._start_point = None
        self._current_point = None
        self.result_image: Optional[Image.Image] = None

        self._screenshot = ImageGrab.grab(all_screens=True)
        self._pixmap = self._pil_to_qpixmap(self._screenshot)

        screen = QGuiApplication.primaryScreen()
        virtual_geometry = screen.virtualGeometry() if screen else QRect(
            0, 0, self._screenshot.width, self._screenshot.height
        )
        self.setGeometry(virtual_geometry)
        # Physical screenshot pixels vs Qt's logical widget coordinates can
        # differ under DPI scaling - scale mouse coordinates up to the
        # screenshot's pixel space when cropping, so the result lines up
        # exactly with what was on screen.
        self._scale_x = self._screenshot.width / max(virtual_geometry.width(), 1)
        self._scale_y = self._screenshot.height / max(virtual_geometry.height(), 1)

    @staticmethod
    def _pil_to_qpixmap(img: Image.Image) -> QPixmap:
        rgb = img.convert("RGB")
        data = rgb.tobytes("raw", "RGB")
        qimage = QImage(data, rgb.width, rgb.height, rgb.width * 3, QImage.Format_RGB888)
        # QImage doesn't copy the buffer by default; `data` must outlive it.
        return QPixmap.fromImage(qimage.copy())

    def paintEvent(self, event) -> None:  # noqa: N802 (Qt override)
        painter = QPainter(self)
        painter.drawPixmap(self.rect(), self._pixmap)

        if self._start_point is not None and self._current_point is not None:
            selection = QRect(self._start_point, self._current_point).normalized()
            full = QPainterPath()
            full.addRect(QRectF(self.rect()))
            hole = QPainterPath()
            hole.addRect(QRectF(selection))
            painter.fillPath(full.subtracted(hole), _DIM_COLOR)
            pen = QPen(_SELECTION_BORDER_COLOR, 2)
            painter.setPen(pen)
            painter.drawRect(selection)
        else:
            painter.fillRect(self.rect(), _DIM_COLOR)

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            self._start_point = event.pos()
            self._current_point = event.pos()
            self.update()

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._start_point is not None:
            self._current_point = event.pos()
            self.update()

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if self._start_point is None or event.button() != Qt.LeftButton:
            return
        selection = QRect(self._start_point, self._current_point).normalized()
        if selection.width() >= _MIN_SELECTION_SIZE and selection.height() >= _MIN_SELECTION_SIZE:
            box = (
                round(selection.left() * self._scale_x),
                round(selection.top() * self._scale_y),
                round(selection.right() * self._scale_x),
                round(selection.bottom() * self._scale_y),
            )
            try:
                self.result_image = self._screenshot.crop(box)
            except Exception:
                logging.exception("failed to crop OCR selection")
                self.result_image = None
        self.accept()

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key_Escape:
            self.result_image = None
            self.reject()
        else:
            super().keyPressEvent(event)

    @classmethod
    def capture(cls) -> Optional[Image.Image]:
        """Show the fullscreen selector and block until the user finishes
        dragging (or cancels with Esc). Must be called on the Qt thread.
        Returns the cropped region as a PIL Image, or None."""
        dialog = cls()
        dialog.exec()
        return dialog.result_image
