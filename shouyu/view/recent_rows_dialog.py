"""alt+/ window: the latest rows of today's sheet, text and pictures.

Rows are listed oldest at the top and newest at the bottom, so the most
recent save sits next to the cursor. Esc closes the window. The entries are
produced by KbExcel.recent_entries(); this module only renders them.
"""
from __future__ import annotations

import re
from html import escape
from typing import List, Optional

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QCursor, QGuiApplication, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from shouyu.view.styles import PANEL_COLOR_HEX, SUBTEXT_COLOR_HEX

PICTURE_MAX_WIDTH = 640
PICTURE_MAX_HEIGHT = 480
WINDOW_WIDTH = 720
WINDOW_HEIGHT = 820
SCREEN_MARGIN = 16  # gap from the screen edges (taskbar is already excluded)
_BREAK_EVERY = 20  # long tokens (paths, ids) get a break opportunity this often


def _breakable(text: str) -> str:
    """Let long unbroken tokens wrap by inserting zero-width spaces."""
    parts = []
    for token in re.split(r'(\s+)', text):
        if len(token) > _BREAK_EVERY and not token.isspace():
            token = '​'.join(token[i:i + _BREAK_EVERY] for i in range(0, len(token), _BREAK_EVERY))
        parts.append(token)
    return ''.join(parts)


class RecentRowsDialog(QDialog):
    _instance: Optional["RecentRowsDialog"] = None

    @classmethod
    def show_entries(cls, title: str, entries: List[dict]) -> None:
        if cls._instance is None:
            cls._instance = RecentRowsDialog()
        cls._instance._render(title, entries)
        cls._instance.show_bottom_right()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("授渔 · 最近内容")
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.setMinimumSize(560, 520)
        self.resize(WINDOW_WIDTH, WINDOW_HEIGHT)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 12)
        layout.setSpacing(10)

        self._header = QLabel()
        self._header.setObjectName("SubtitleLabel")
        layout.addWidget(self._header)

        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setFrameShape(QFrame.NoFrame)
        self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self._content = QWidget()
        self._rows = QVBoxLayout(self._content)
        self._rows.setContentsMargins(0, 0, 0, 0)
        self._rows.setSpacing(8)
        self._rows.addStretch(1)  # keeps the rows packed at the top when there are few
        self._scroll.setWidget(self._content)
        layout.addWidget(self._scroll, 1)

        hint = QLabel("Esc 关闭")
        hint.setObjectName("HintLabel")
        layout.addWidget(hint)

    # ---------- public API ----------

    def show_bottom_right(self) -> None:
        """Place the window in the bottom-right corner of the screen the mouse
        is on (falls back to the primary screen), above the taskbar."""
        screen = QGuiApplication.screenAt(QCursor.pos()) or QGuiApplication.primaryScreen()
        if screen is not None:
            geo = screen.availableGeometry()
            height = min(self.height(), geo.height() - 2 * SCREEN_MARGIN)
            self.resize(min(self.width(), geo.width() - 2 * SCREEN_MARGIN), height)
            self.move(
                geo.right() - self.width() - SCREEN_MARGIN + 1,
                geo.bottom() - self.height() - SCREEN_MARGIN + 1,
            )
        self.show()
        self.raise_()
        self.activateWindow()

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key_Escape:
            self.close()
            return
        super().keyPressEvent(event)

    # ---------- rendering ----------

    def _render(self, title: str, entries: List[dict]) -> None:
        self._header.setText(f"{title} · 最近 {len(entries)} 行")
        while self._rows.count() > 1:  # keep the trailing stretch
            item = self._rows.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for entry in reversed(entries):  # oldest first; the newest ends up at the bottom
            if entry.get("kind") == "image":
                self._rows.insertWidget(self._rows.count() - 1, self._image_row(entry))
            else:
                self._rows.insertWidget(self._rows.count() - 1, self._text_row(entry))

        bar = self._scroll.verticalScrollBar()
        QTimer.singleShot(0, lambda: bar.setValue(bar.maximum()))

    @staticmethod
    def _card() -> QFrame:
        card = QFrame()
        card.setObjectName("RecentRow")
        card.setStyleSheet(f"QFrame#RecentRow {{ background-color: {PANEL_COLOR_HEX}; border-radius: 6px; }}")
        return card

    def _text_row(self, entry: dict) -> QFrame:
        card = self._card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(2)
        label = QLabel(
            f'<span style="color:{SUBTEXT_COLOR_HEX}">{escape(entry.get("label", ""))}</span>'
        )
        label.setTextFormat(Qt.RichText)
        layout.addWidget(label)
        body = QLabel(escape(_breakable(str(entry.get("text", "")))).replace("\n", "<br>"))
        body.setTextFormat(Qt.RichText)
        body.setWordWrap(True)
        body.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(body)
        return card

    def _image_row(self, entry: dict) -> QFrame:
        card = self._card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 8, 12, 8)
        layout.setSpacing(4)
        caption = QLabel(f'<span style="color:{SUBTEXT_COLOR_HEX}">{escape(entry.get("label", ""))} · 图片</span>')
        caption.setTextFormat(Qt.RichText)
        layout.addWidget(caption)

        pixmap = QPixmap()
        png = entry.get("png")
        if png and pixmap.loadFromData(png):
            pixmap = pixmap.scaled(
                PICTURE_MAX_WIDTH, PICTURE_MAX_HEIGHT, Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            picture = QLabel()
            picture.setPixmap(pixmap)
            layout.addWidget(picture)
        else:
            missing = QLabel("（图片无法读取）")
            missing.setObjectName("HintLabel")
            layout.addWidget(missing)
        return card
