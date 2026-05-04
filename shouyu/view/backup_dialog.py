"""Backup restore dialog.

Shows the list of `*_backup_*.xlsx` files in the same folder as the
canonical Excel, lets the user pick one, and copies it back over the
canonical file (with a pre-restore safety copy taken first so the
operation is reversible).

Triggered by the `restore_backup` hotkey (configurable in kb.ini) or
automatically when the canonical Excel is corrupt and we managed to
auto-recover from a backup at startup.
"""
from __future__ import annotations

import logging
import os
import time
from typing import List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from shouyu.config import Config
from shouyu.service.excel import KbExcel
from shouyu.util.process import ProcessManager
from shouyu.view.styles import (
    ACCENT_COLOR_HEX,
    SUBTEXT_COLOR_HEX,
    TEXT_COLOR_HEX,
)


def _human_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.2f} MB"


def _human_age(seconds: float) -> str:
    seconds = max(0, int(seconds))
    if seconds < 60:
        return f"{seconds}s 前"
    if seconds < 3600:
        return f"{seconds // 60}m 前"
    if seconds < 86400:
        return f"{seconds // 3600}h 前"
    return f"{seconds // 86400}d 前"


class BackupRestoreDialog(QDialog):
    """Modal dialog for picking a backup to restore."""

    _instance: Optional["BackupRestoreDialog"] = None

    @classmethod
    def get_or_create(cls) -> "BackupRestoreDialog":
        if cls._instance is None:
            cls._instance = BackupRestoreDialog()
        return cls._instance

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("从备份恢复 Excel")
        self.setModal(True)
        self.setMinimumSize(720, 480)

        self._excel_path: str = Config.excel_path()
        self._backups: List[str] = []

        self._build_ui()

    # ---------- public API ----------

    def show_centered(self, *, recovered_from: Optional[str] = None) -> None:
        self._excel_path = Config.excel_path()
        self._reload_backups(highlight=recovered_from)
        if recovered_from:
            self.banner_label.setText(
                f"⚠ 检测到主文件 {os.path.basename(self._excel_path)} 已损坏，"
                f"已自动加载备份 {os.path.basename(recovered_from)}。"
                "保存任何改动会把这份恢复后的内容写回主文件；如不满意请在下面挑一个更早的版本。"
            )
            self.banner_label.setVisible(True)
        else:
            self.banner_label.setVisible(False)
        self.show()
        self.raise_()
        self.activateWindow()
        screen = self.screen() or self.window().screen()
        if screen is not None:
            geo = screen.availableGeometry()
            self.move(
                geo.center().x() - self.width() // 2,
                geo.center().y() - self.height() // 2,
            )

    # ---------- ui ----------

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 16)
        layout.setSpacing(12)

        title = QLabel("从备份恢复 Excel")
        title.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {TEXT_COLOR_HEX};")
        layout.addWidget(title)

        path_row = QHBoxLayout()
        path_row.addWidget(QLabel("当前主文件："))
        self.path_label = QLabel("")
        self.path_label.setStyleSheet(f"color: {ACCENT_COLOR_HEX};")
        path_row.addWidget(self.path_label, 1)
        layout.addLayout(path_row)

        self.banner_label = QLabel("")
        self.banner_label.setWordWrap(True)
        self.banner_label.setStyleSheet(
            "padding: 10px 12px; "
            "background-color: rgba(255, 180, 84, 0.12); "
            "border: 1px solid rgba(255, 180, 84, 0.45); "
            "border-radius: 6px; "
            "color: #FFB454; font-size: 12px;"
        )
        self.banner_label.setVisible(False)
        layout.addWidget(self.banner_label)

        hint = QLabel(
            "👇 选择一个时间点的备份还原。还原前会把当前文件复制为 "
            "`<name>.pre_restore_<时间>.xlsx`，所以这一步是可撤销的。"
        )
        hint.setStyleSheet(f"color: {SUBTEXT_COLOR_HEX}; font-size: 12px;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(lambda _it: self._on_preview_clicked())
        layout.addWidget(self.list_widget, stretch=1)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)

        refresh_btn = QPushButton("🔄 刷新")
        refresh_btn.setAutoDefault(False)
        refresh_btn.clicked.connect(lambda: self._reload_backups())
        button_row.addWidget(refresh_btn)

        preview_btn = QPushButton("👀 在 Excel 中预览所选")
        preview_btn.setAutoDefault(False)
        preview_btn.clicked.connect(self._on_preview_clicked)
        button_row.addWidget(preview_btn)

        button_row.addStretch(1)

        cancel_btn = QPushButton("取消")
        cancel_btn.setAutoDefault(False)
        cancel_btn.clicked.connect(self.reject)
        button_row.addWidget(cancel_btn)

        restore_btn = QPushButton("✅ 还原所选")
        restore_btn.setObjectName("PrimaryButton")
        restore_btn.setAutoDefault(False)
        restore_btn.clicked.connect(self._on_restore_clicked)
        button_row.addWidget(restore_btn)

        layout.addLayout(button_row)

    # ---------- data ----------

    def _reload_backups(self, *, highlight: Optional[str] = None) -> None:
        self.path_label.setText(self._excel_path)
        self._backups = KbExcel.list_backups(self._excel_path)
        self.list_widget.clear()
        if not self._backups:
            empty = QListWidgetItem("（没有找到备份）")
            empty.setFlags(empty.flags() & ~Qt.ItemIsEnabled & ~Qt.ItemIsSelectable)
            self.list_widget.addItem(empty)
            return
        now = time.time()
        highlight_idx = -1
        for i, path in enumerate(self._backups):
            try:
                mtime = os.path.getmtime(path)
                size = os.path.getsize(path)
            except OSError:
                continue
            label = (
                f"{time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(mtime))}"
                f"   ·   {_human_size(size):>10}   ·   {_human_age(now - mtime):>8}"
                f"\n{os.path.basename(path)}"
            )
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, path)
            self.list_widget.addItem(item)
            if highlight and os.path.normpath(path) == os.path.normpath(highlight):
                highlight_idx = i
        if self.list_widget.count() > 0:
            self.list_widget.setCurrentRow(max(0, highlight_idx))

    def _selected_backup(self) -> Optional[str]:
        item = self.list_widget.currentItem()
        if item is None:
            return None
        path = item.data(Qt.UserRole)
        return path if isinstance(path, str) else None

    # ---------- actions ----------

    def _on_preview_clicked(self) -> None:
        path = self._selected_backup()
        if not path:
            return
        try:
            ProcessManager.open_file(path)
        except Exception:
            logging.exception("failed to open backup preview")
            QMessageBox.warning(self, "无法打开", f"打开预览失败：\n{path}")

    def _on_restore_clicked(self) -> None:
        path = self._selected_backup()
        if not path:
            QMessageBox.information(self, "请选择", "请先选择一个备份。")
            return
        confirm = QMessageBox.question(
            self,
            "确认还原",
            (
                f"将用以下备份覆盖当前主文件：\n\n"
                f"{path}\n\n"
                f"主文件: {self._excel_path}\n\n"
                "（当前文件会被另存为 `*.pre_restore_*.xlsx`，这一步是可撤销的）"
            ),
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return
        try:
            safety = KbExcel.restore_from_backup(path, self._excel_path)
        except Exception as e:
            logging.exception("restore failed")
            QMessageBox.critical(self, "还原失败", f"还原失败：\n{e}")
            return
        msg = (
            f"已从备份还原主文件。\n\n"
            f"原主文件已另存为：\n{safety or '（无）'}\n\n"
            "建议重启 shouyu 让所有缓存重新加载。"
        )
        QMessageBox.information(self, "还原成功", msg)
        self.accept()
