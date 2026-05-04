"""Quick duration picker for tasks.

Replaces a free-form QInputDialog with a small grid of preset buttons
(15m / 30m / 1h / 2h / 半天 / 一天) plus a custom spinbox. Lowering the
friction here was an explicit user request: making it easier to estimate
encourages people to actually do it.
"""
from __future__ import annotations

from typing import List, Optional, Tuple

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from shouyu.view.styles import ACCENT_COLOR_HEX, SUBTEXT_COLOR_HEX, TEXT_COLOR_HEX


_PRESETS: List[Tuple[str, int]] = [
    ("不确定", 0),
    ("15 分钟", 15),
    ("30 分钟", 30),
    ("1 小时", 60),
    ("90 分钟", 90),
    ("2 小时", 120),
    ("半天 (4h)", 240),
    ("一天 (8h)", 480),
]


def _format_minutes(minutes: int) -> str:
    if minutes <= 0:
        return "未设定"
    if minutes < 60:
        return f"{minutes} 分钟"
    hours, mins = divmod(minutes, 60)
    if mins == 0:
        return f"{hours} 小时"
    return f"{hours}h {mins}m"


class DurationPickerDialog(QDialog):
    """Modal dialog returning a duration in minutes (or -1 for cancel)."""

    def __init__(
        self,
        current: int = 0,
        task_text: str = "",
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("预计时长")
        self.setModal(True)
        self.value = max(0, int(current or 0))

        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(12)

        if task_text:
            shown = task_text if len(task_text) <= 40 else task_text[:39] + "…"
            header = QLabel(f"为「{shown}」设定预计时长")
            header.setStyleSheet(
                f"color: {TEXT_COLOR_HEX}; font-size: 14px; font-weight: 600;"
            )
            header.setWordWrap(True)
            layout.addWidget(header)

        hint = QLabel(
            "💡 估时不必精确，先选一个最接近的；坚持几天后你会越来越准。"
        )
        hint.setStyleSheet(f"color: {SUBTEXT_COLOR_HEX}; font-size: 12px;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        layout.addWidget(self._section_label("快速选择"))

        grid = QGridLayout()
        grid.setSpacing(8)
        for i, (label, mins) in enumerate(_PRESETS):
            btn = QPushButton(label)
            btn.setMinimumHeight(36)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setAutoDefault(False)
            btn.setDefault(False)
            if mins == self.value:
                btn.setObjectName("PrimaryButton")
            btn.clicked.connect(lambda _checked=False, m=mins: self._pick(m))
            grid.addWidget(btn, i // 4, i % 4)
        layout.addLayout(grid)

        layout.addWidget(self._section_label("自定义（分钟，0 = 不确定）"))

        custom_row = QHBoxLayout()
        custom_row.setSpacing(8)
        self.custom_input = QSpinBox()
        self.custom_input.setRange(0, 600)
        self.custom_input.setSingleStep(5)
        self.custom_input.setValue(self.value)
        self.custom_input.setMinimumHeight(34)
        self.custom_input.setSuffix(" 分钟")
        custom_row.addWidget(self.custom_input, 1)

        confirm_btn = QPushButton("确定")
        confirm_btn.setObjectName("PrimaryButton")
        confirm_btn.setMinimumHeight(34)
        confirm_btn.setCursor(Qt.PointingHandCursor)
        confirm_btn.clicked.connect(lambda: self._pick(self.custom_input.value()))
        custom_row.addWidget(confirm_btn)
        layout.addLayout(custom_row)

        bottom = QHBoxLayout()
        bottom.addStretch(1)
        cancel_btn = QPushButton("取消")
        cancel_btn.setCursor(Qt.PointingHandCursor)
        cancel_btn.setAutoDefault(False)
        cancel_btn.setDefault(False)
        cancel_btn.clicked.connect(self.reject)
        bottom.addWidget(cancel_btn)
        layout.addLayout(bottom)

    @staticmethod
    def _section_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {ACCENT_COLOR_HEX}; font-size: 12px; font-weight: 600;")
        return lbl

    def _pick(self, minutes: int) -> None:
        self.value = max(0, int(minutes or 0))
        self.accept()

    @classmethod
    def get_duration(
        cls,
        current: int,
        task_text: str = "",
        parent: Optional[QWidget] = None,
    ) -> int:
        """Run the dialog modally. Returns minutes, or -1 if cancelled."""
        dlg = cls(current=current, task_text=task_text, parent=parent)
        if dlg.exec() == QDialog.Accepted:
            return dlg.value
        return -1


__all__ = ["DurationPickerDialog", "_format_minutes"]
