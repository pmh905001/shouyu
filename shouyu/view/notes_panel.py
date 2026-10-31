"""Notes panel opened by `alt+/`: recent clipboard / screenshot / OCR records
as chat bubbles, with instant search.

Design reference: docs/design/retrieval/20261031-01-notes-panel-and-day-view.md
(§4.2). Requirements REC-01 .. REC-17.

- Opens on the screen under the mouse, focus in the search box (at the
  bottom, like a chat input); the newest record sits right above it.
- Empty search box: browse, newest page first, older pages load when
  scrolling to the top.
- Typing: every keystroke restarts a 120 ms timer, then `NoteSearcher` runs
  the query on its own thread (near days first). Results of an outdated
  generation are dropped.
- Keyboard only: arrows move, Enter copies, Esc hides, Tab / alt+1..5 switch
  the range. The panel hides when it loses focus, except while one of its
  own dialogs / menus is open.
"""
from __future__ import annotations

import logging
import time
from datetime import datetime, timedelta
from typing import List, Optional

from PySide6.QtCore import QEvent, QObject, Qt, QTimer, Signal
from PySide6.QtGui import QCursor, QGuiApplication, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QButtonGroup,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from shouyu.service import note_store
from shouyu.service.note_search import NoteSearcher
from shouyu.service.note_store import Note, NoteQuery
from shouyu.view.note_list import NoteListView, _clip_text
from shouyu.view.styles import ACCENT_COLOR_HEX, BG_COLOR_HEX, SUBTEXT_COLOR_HEX

SEARCH_DEBOUNCE_MS = 120
PANEL_SIZE = (580, 740)

SCOPE_TODAY = 'today'
SCOPE_YESTERDAY = 'yesterday'
SCOPE_DAY_BEFORE = 'day_before'
SCOPE_WEEK = 'week'
SCOPE_ALL = 'all'
SCOPES = [
    (SCOPE_TODAY, '今天'),
    (SCOPE_YESTERDAY, '昨天'),
    (SCOPE_DAY_BEFORE, '前天'),
    (SCOPE_WEEK, '7天'),
    (SCOPE_ALL, '全部'),
]


def _shift(date_str: str, days: int) -> str:
    return (datetime.strptime(date_str, '%Y-%m-%d') + timedelta(days=days)).strftime('%Y-%m-%d')


def scope_range(scope: str, today: str):
    """(date_from, date_to) for a scope; None means unbounded."""
    if scope == SCOPE_TODAY:
        return today, today
    if scope == SCOPE_YESTERDAY:
        d = _shift(today, -1)
        return d, d
    if scope == SCOPE_DAY_BEFORE:
        d = _shift(today, -2)
        return d, d
    if scope == SCOPE_WEEK:
        return _shift(today, -6), today
    return None, None


class NotesPanel(QWidget):
    _instance: Optional["NotesPanel"] = None

    # Marshals NoteSearcher callbacks (worker thread) onto the Qt thread.
    _results = Signal(int, list, bool, bool)

    @classmethod
    def get_or_create(cls) -> "NotesPanel":
        if cls._instance is None:
            cls._instance = NotesPanel()
        return cls._instance

    def __init__(self) -> None:
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowTitle('授渔 · 记录')
        self.setObjectName('NotesPanel')
        self.setStyleSheet(
            f'#NotesPanel {{ background-color: {BG_COLOR_HEX}; border: 1px solid #3A3A3A; border-radius: 10px; }}'
            f'QPushButton[scope="true"] {{ padding: 4px 12px; border-radius: 12px; }}'
            f'QPushButton[scope="true"]:checked {{ background-color: {ACCENT_COLOR_HEX}; color: white; border: none; }}'
        )
        self.resize(*PANEL_SIZE)
        self._today = time.strftime('%Y-%m-%d')
        self._scope = SCOPE_ALL
        self._generation = 0
        self._searching = False
        self._result_count = 0
        self._oldest: Optional[tuple] = None  # browse paging cursor (plan_date, id)
        self._no_more_older = False
        self._dialogs_open = 0
        self._searcher = NoteSearcher(today_fn=lambda: self._today)
        self._results.connect(self._on_results, Qt.QueuedConnection)

        self._debounce = QTimer(self)
        self._debounce.setSingleShot(True)
        self._debounce.setInterval(SEARCH_DEBOUNCE_MS)
        self._debounce.timeout.connect(self._run_search)

        self._build_ui()
        self._install_shortcuts()

    # ---------- ui ----------

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        top = QHBoxLayout()
        self._scope_group = QButtonGroup(self)
        self._scope_group.setExclusive(True)
        self._scope_buttons = {}
        for i, (key, label) in enumerate(SCOPES):
            btn = QPushButton(label)
            btn.setProperty('scope', True)
            btn.setCheckable(True)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setToolTip(f'alt+{i + 1}')
            btn.clicked.connect(lambda _checked=False, k=key: self.set_scope(k))
            self._scope_group.addButton(btn)
            self._scope_buttons[key] = btn
            top.addWidget(btn)
        top.addStretch(1)
        self.count_label = QLabel('')
        self.count_label.setObjectName('HintLabel')
        top.addWidget(self.count_label)
        root.addLayout(top)

        self.list_view = NoteListView(self, today_fn=lambda: self._today)
        self.list_view.setFocusPolicy(Qt.NoFocus)
        self.list_view.copied.connect(self._on_copied)
        self.list_view.older_requested.connect(self._load_older)
        self.list_view.open_day_requested.connect(self._open_day_view)
        self.list_view.notes_changed.connect(self._on_notes_changed)
        self.list_view.dialog_open.connect(self._on_dialog_open)
        root.addWidget(self.list_view, 1)

        self.empty_label = QLabel('')
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet(f'color: {SUBTEXT_COLOR_HEX}; padding: 6px;')
        self.empty_label.hide()
        root.addWidget(self.empty_label)

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText('搜索记录：输入一两个字即可；空格分开多个词')
        self.search_box.setClearButtonEnabled(True)
        self.search_box.textChanged.connect(self._on_text_changed)
        self.search_box.installEventFilter(self)
        root.addWidget(self.search_box)

        hint = QLabel('↑↓ 选择 · Enter 复制 · Space 预览 · Tab 切换范围 · Ctrl+D 日视图 · Del 删除 · Esc 关闭')
        hint.setObjectName('HintLabel')
        root.addWidget(hint)

    def _install_shortcuts(self) -> None:
        for i, (key, _label) in enumerate(SCOPES):
            sc = QShortcut(QKeySequence(f'Alt+{i + 1}'), self)
            sc.activated.connect(lambda k=key: self.set_scope(k))
        QShortcut(QKeySequence('Ctrl+D'), self).activated.connect(self._open_current_day)
        QShortcut(QKeySequence('Ctrl+E'), self).activated.connect(self.list_view.edit_current)
        QShortcut(QKeySequence('Ctrl+Space'), self).activated.connect(self.list_view.preview_current)

    # ---------- open / close ----------

    def toggle(self) -> None:
        if self.isVisible() and self.isActiveWindow():
            self.hide()
            return
        self.open()

    def open(self) -> None:
        self._today = time.strftime('%Y-%m-%d')
        self._scope = SCOPE_ALL
        self._scope_buttons[SCOPE_ALL].setChecked(True)
        self.search_box.blockSignals(True)
        self.search_box.clear()
        self.search_box.blockSignals(False)
        self._place_on_cursor_screen()
        self.show()
        self.raise_()
        self.activateWindow()
        self.search_box.setFocus()
        self._reload()

    def _place_on_cursor_screen(self) -> None:
        screen = QGuiApplication.screenAt(QCursor.pos()) or QGuiApplication.primaryScreen()
        if screen is None:
            return
        avail = screen.availableGeometry()
        w = min(PANEL_SIZE[0], avail.width() - 40)
        h = min(PANEL_SIZE[1], avail.height() - 40)
        self.setGeometry(avail.left() + (avail.width() - w) // 2, avail.top() + (avail.height() - h) // 2, w, h)

    def changeEvent(self, event) -> None:
        super().changeEvent(event)
        if event.type() == QEvent.ActivationChange and not self.isActiveWindow():
            # Deferred: activation briefly moves to our own dialogs / menus.
            QTimer.singleShot(0, self._hide_if_inactive)

    def _hide_if_inactive(self) -> None:
        if self._dialogs_open or not self.isVisible():
            return
        active = QGuiApplication.focusWindow()
        if active is not None and active is not self.windowHandle():
            # One of our own child dialogs got focus without telling us.
            transient = active.transientParent()
            if transient is self.windowHandle():
                return
        if not self.isActiveWindow():
            self.hide()

    def _on_dialog_open(self, is_open: bool) -> None:
        self._dialogs_open = max(0, self._dialogs_open + (1 if is_open else -1))
        if not is_open:
            self.raise_()
            self.activateWindow()
            self.search_box.setFocus()

    # ---------- keyboard ----------

    def eventFilter(self, obj: QObject, event) -> bool:
        if obj is self.search_box and event.type() == QEvent.KeyPress:
            key, mods = event.key(), event.modifiers()
            empty = not self.search_box.text()
            if key == Qt.Key_Escape:
                self.hide()
                return True
            if key in (Qt.Key_Return, Qt.Key_Enter):
                self.list_view.copy_current()
                return True
            if key == Qt.Key_Up:
                self.list_view.move_selection(-1)
                return True
            if key == Qt.Key_Down:
                self.list_view.move_selection(1)
                return True
            if key == Qt.Key_PageUp:
                self.list_view.move_selection(-8)
                return True
            if key == Qt.Key_PageDown:
                self.list_view.move_selection(8)
                return True
            if key == Qt.Key_Tab:
                self._cycle_scope(1)
                return True
            if key == Qt.Key_Backtab:
                self._cycle_scope(-1)
                return True
            if key == Qt.Key_Space and empty and not mods:
                self.list_view.preview_current()
                return True
            if key == Qt.Key_Delete and empty:
                self.list_view.delete_current()
                return True
        return super().eventFilter(obj, event)

    def keyPressEvent(self, event) -> None:
        if event.key() == Qt.Key_Escape:
            self.hide()
            return
        super().keyPressEvent(event)

    # ---------- scope / search ----------

    def set_scope(self, scope: str) -> None:
        self._scope = scope
        self._scope_buttons[scope].setChecked(True)
        self.search_box.setFocus()
        self._reload()

    def _cycle_scope(self, step: int) -> None:
        keys = [k for k, _ in SCOPES]
        self.set_scope(keys[(keys.index(self._scope) + step) % len(keys)])

    def _terms(self) -> List[str]:
        return note_store.parse_terms(self.search_box.text())

    def _on_text_changed(self, _text: str) -> None:
        self._debounce.start()

    def _run_search(self) -> None:
        self._reload()

    def _reload(self) -> None:
        """Re-query for the current scope + search text."""
        self._debounce.stop()
        self._generation += 1
        terms = self._terms()
        date_from, date_to = scope_range(self._scope, self._today)
        self._oldest = None
        self._no_more_older = False
        if not terms:
            self._searching = False
            try:
                if self._scope == SCOPE_ALL:
                    notes = note_store.page_before(_shift(self._today, 1), 0)
                    self._remember_oldest(notes)
                else:
                    notes = note_store.search(NoteQuery(date_from=date_from, date_to=date_to))
                    self._no_more_older = True
            except Exception:
                logging.exception('failed to load notes')
                notes = []
            self.list_view.set_notes(notes, [])
            self._update_labels(len(notes))
            return
        self._searching = True
        self._result_count = 0
        generation = self._generation
        self._searcher.submit(
            generation,
            NoteQuery(terms=terms, date_from=date_from, date_to=date_to),
            lambda g, found, appended, finished: self._results.emit(g, found, appended, finished),
        )
        self._update_labels(None)

    def _on_results(self, generation: int, found: list, appended: bool, finished: bool) -> None:
        if generation != self._generation:
            return
        terms = self._terms()
        if not appended:
            self.list_view.set_notes(found, terms)
            self._result_count = len(found)
        else:
            self.list_view.note_model.terms = terms
            self.list_view.prepend_notes(found)
            self._result_count += len(found)
            if self.list_view.current_note() is None:
                self.list_view.select_last()
        if finished:
            self._searching = False
        self._update_labels(self._result_count)

    def _remember_oldest(self, notes: List[Note]) -> None:
        if notes:
            self._oldest = (notes[0].plan_date, notes[0].id)
        if len(notes) < note_store.PAGE_SIZE:
            self._no_more_older = True

    def _load_older(self) -> None:
        """Browse mode in the "全部" scope: fetch the page before the oldest shown."""
        if self._terms() or self._scope != SCOPE_ALL or self._no_more_older or self._oldest is None:
            return
        try:
            older = note_store.page_before(*self._oldest)
        except Exception:
            logging.exception('failed to load older notes')
            return
        self._remember_oldest(older)
        self.list_view.prepend_notes(older)
        self._update_labels(len(self.list_view.note_model.notes()))

    def _update_labels(self, count: Optional[int]) -> None:
        terms = self._terms()
        if count is None or self._searching:
            self.count_label.setText('搜索中…' if count is None else f'{count} 条 · 搜索中…')
        elif terms and count >= note_store.SEARCH_LIMIT:
            self.count_label.setText(f'最近 {count} 条（更多请缩小范围）')
        elif not terms and not self._no_more_older:
            self.count_label.setText(f'已加载 {count} 条，往上滚动加载更多')
        else:
            self.count_label.setText(f'{count} 条')
        has_today = any(n.plan_date == self._today for n in self.list_view.note_model.notes())
        if terms:
            text = '' if (count or self._searching) else '没有找到匹配的记录'
        elif self._scope in (SCOPE_ALL, SCOPE_TODAY) and not has_today:
            text = '今天还没有记录' + ('，往上滚动查看以前的' if self.list_view.note_model.notes() else '')
        elif not count:
            text = '这一天没有记录' if self._scope in (SCOPE_YESTERDAY, SCOPE_DAY_BEFORE) else '没有记录'
        else:
            text = ''
        self.empty_label.setText(text)
        self.empty_label.setVisible(bool(text))

    # ---------- actions ----------

    def _on_copied(self, note: Note) -> None:
        self.hide()
        try:
            from shouyu.view.msgbox import MessageBox

            preview = '截图' if note.kind == note_store.KIND_IMAGE else _clip_text(note.text)[0][:60]
            MessageBox.pop_up_message(title='已复制', msg=preview)
        except Exception:
            logging.exception('failed to show copied toast')

    def _on_notes_changed(self) -> None:
        self._update_labels(len(self.list_view.note_model.notes()))

    def _open_current_day(self) -> None:
        note = self.list_view.current_note()
        self._open_day_view(note.plan_date if note is not None else self._today)

    def _open_day_view(self, plan_date: str) -> None:
        self.hide()
        try:
            from shouyu.view.day_view import DayView

            DayView.get_or_create().open_date(plan_date)
        except Exception:
            logging.exception('failed to open the day view')
