"""Chat-style bubble list of notes, shared by the notes panel and the day view.

Design reference: docs/design/retrieval/20261031-01-notes-panel-and-day-view.md
(§4.3). Rows are either a day header or one note, oldest at the top and the
newest at the bottom (closest to the panel's search box). A `QListView` +
delegate only paints visible rows, so a day with tens of thousands of notes
does not create tens of thousands of widgets. Image thumbnails are decoded
the first time their row is painted and cached by attachment id.
"""
from __future__ import annotations

import html
import logging
import time
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Callable, Dict, List, Optional, Tuple

from PySide6.QtCore import QAbstractListModel, QModelIndex, QRect, QRectF, QSize, Qt, QTimer, Signal
from PySide6.QtGui import (
    QAbstractTextDocumentLayout,
    QAction,
    QColor,
    QFont,
    QFontMetrics,
    QGuiApplication,
    QImage,
    QPainter,
    QPainterPath,
    QPalette,
    QPixmap,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QLabel,
    QListView,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QScrollArea,
    QStyle,
    QStyledItemDelegate,
    QStyleOptionViewItem,
    QVBoxLayout,
    QWidget,
)

from shouyu.service import note_store
from shouyu.service.note_store import Note
from shouyu.view.styles import ACCENT_COLOR_HEX, PANEL_COLOR_HEX, SUBTEXT_COLOR_HEX, TEXT_COLOR_HEX

SOURCE_LABELS = {
    note_store.SOURCE_CLIPBOARD: '剪贴板',
    note_store.SOURCE_OCR: 'OCR',
    note_store.SOURCE_MANUAL: '手动',
    note_store.SOURCE_EXCEL_IMPORT: 'Excel 导入',
}
_WEEKDAYS = '一二三四五六日'

MAX_TEXT_LINES = 6
MAX_TEXT_CHARS = 600
THUMB_MAX_W = 240
THUMB_MAX_H = 160
THUMB_CACHE_SIZE = 300

HEADER_ROLE = Qt.UserRole + 1
NOTE_ROLE = Qt.UserRole + 2


# ---------- formatting ----------

def day_label(plan_date: str, today: Optional[str] = None) -> str:
    today = today or time.strftime('%Y-%m-%d')
    try:
        d = datetime.strptime(plan_date, '%Y-%m-%d')
        t = datetime.strptime(today, '%Y-%m-%d')
    except ValueError:
        return plan_date
    delta = (t - d).days
    if delta == 0:
        return '今天'
    if delta == 1:
        return f'昨天 {d:%m-%d}'
    if delta == 2:
        return f'前天 {d:%m-%d}'
    weekday = f'星期{_WEEKDAYS[d.weekday()]}'
    if d.year != t.year:
        return f'{d:%Y-%m-%d} {weekday}'
    return f'{d:%m-%d} {weekday}'


def note_time(note: Note) -> str:
    """Local HH:MM of `created_at` (stored as UTC); imported notes have none."""
    if note.source == note_store.SOURCE_EXCEL_IMPORT:
        return '—'
    try:
        utc = datetime.strptime(note.created_at.replace('Z', ''), '%Y-%m-%dT%H:%M:%S.%f')
    except (ValueError, AttributeError):
        return ''
    return utc.replace(tzinfo=timezone.utc).astimezone().strftime('%H:%M')


def note_meta(note: Note) -> str:
    source = '截图' if note.kind == note_store.KIND_IMAGE and note.source == note_store.SOURCE_CLIPBOARD else (
        SOURCE_LABELS.get(note.source, note.source)
    )
    edited = ' · 已修改' if note.updated_at else ''
    return f'{note_time(note)}  {source}{edited}'


def _clip_text(text: str) -> Tuple[str, bool]:
    lines = (text or '').splitlines() or ['']
    clipped = '\n'.join(lines[:MAX_TEXT_LINES])
    truncated = len(lines) > MAX_TEXT_LINES
    if len(clipped) > MAX_TEXT_CHARS:
        clipped = clipped[:MAX_TEXT_CHARS]
        truncated = True
    return clipped, truncated


def highlight_html(text: str, terms: List[str]) -> str:
    """HTML for `text` with every (case-insensitive) occurrence of any term
    highlighted. The text is escaped first, so notes can't inject markup."""
    spans: List[Tuple[int, int]] = []
    lowered = text.lower()
    for term in terms:
        needle = term.lower()
        if not needle:
            continue
        start = lowered.find(needle)
        while start != -1:
            spans.append((start, start + len(needle)))
            start = lowered.find(needle, start + len(needle))
    spans.sort()
    merged: List[Tuple[int, int]] = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    out, pos = [], 0
    for s, e in merged:
        out.append(html.escape(text[pos:s]))
        out.append(f'<span style="background-color:#7A5C00;color:#FFFFFF;">{html.escape(text[s:e])}</span>')
        pos = e
    out.append(html.escape(text[pos:]))
    return ''.join(out).replace('\n', '<br>')


# ---------- model ----------

@dataclass
class _Row:
    date: str
    note: Optional[Note] = None  # None: day header


class NoteListModel(QAbstractListModel):
    def __init__(self, parent=None, today_fn: Callable[[], str] = lambda: time.strftime('%Y-%m-%d')) -> None:
        super().__init__(parent)
        self._today_fn = today_fn
        self._notes: List[Note] = []
        self._rows: List[_Row] = []
        self.terms: List[str] = []

    # -- Qt --
    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._rows)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid() or not (0 <= index.row() < len(self._rows)):
            return None
        row = self._rows[index.row()]
        if role == HEADER_ROLE:
            return day_label(row.date, self._today_fn()) if row.note is None else None
        if role == NOTE_ROLE:
            return row.note
        if role == Qt.DisplayRole:
            return row.note.text if row.note is not None else day_label(row.date, self._today_fn())
        return None

    def flags(self, index: QModelIndex):
        if not index.isValid():
            return Qt.NoItemFlags
        if self._rows[index.row()].note is None:
            return Qt.ItemIsEnabled
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable

    # -- content --
    def _rebuild(self) -> None:
        rows: List[_Row] = []
        last_date = None
        for note in self._notes:
            if note.plan_date != last_date:
                rows.append(_Row(date=note.plan_date))
                last_date = note.plan_date
            rows.append(_Row(date=note.plan_date, note=note))
        self._rows = rows

    def set_notes(self, notes: List[Note], terms: Optional[List[str]] = None) -> None:
        """`notes` in any order; shown oldest first."""
        self.beginResetModel()
        self.terms = list(terms or [])
        self._notes = sorted(notes, key=lambda n: (n.plan_date, n.id))
        self._rebuild()
        self.endResetModel()

    def prepend_notes(self, older: List[Note]) -> None:
        if not older:
            return
        known = {n.id for n in self._notes}
        merged = [n for n in older if n.id not in known] + self._notes
        self.beginResetModel()
        self._notes = sorted(merged, key=lambda n: (n.plan_date, n.id))
        self._rebuild()
        self.endResetModel()

    def append_note(self, note: Note) -> None:
        self.prepend_notes([note])

    def remove_note(self, note_id: int) -> None:
        self.beginResetModel()
        self._notes = [n for n in self._notes if n.id != note_id]
        self._rebuild()
        self.endResetModel()

    def replace_note(self, note: Note) -> None:
        for i, row in enumerate(self._rows):
            if row.note is not None and row.note.id == note.id:
                row.note = note
                self._notes = [note if n.id == note.id else n for n in self._notes]
                idx = self.index(i)
                self.dataChanged.emit(idx, idx)
                return

    def notes(self) -> List[Note]:
        return list(self._notes)

    def row_of_note(self, note_id: int) -> int:
        for i, row in enumerate(self._rows):
            if row.note is not None and row.note.id == note_id:
                return i
        return -1

    def note_rows(self) -> List[int]:
        return [i for i, row in enumerate(self._rows) if row.note is not None]


# ---------- delegate ----------

def _png_size(data: bytes) -> Optional[Tuple[int, int]]:
    """Width / height from the PNG header, without decoding the image."""
    if len(data) >= 24 and data[:8] == b'\x89PNG\r\n\x1a\n' and data[12:16] == b'IHDR':
        return int.from_bytes(data[16:20], 'big'), int.from_bytes(data[20:24], 'big')
    return None


def thumb_size(width: int, height: int) -> Tuple[int, int]:
    if width <= 0 or height <= 0:
        return THUMB_MAX_W, THUMB_MAX_H
    scale = min(1.0, THUMB_MAX_W / width, THUMB_MAX_H / height)
    return max(1, round(width * scale)), max(1, round(height * scale))


class _ThumbCache:
    def __init__(self, size: int = THUMB_CACHE_SIZE) -> None:
        self._size = size
        self._items: "OrderedDict[int, Optional[QPixmap]]" = OrderedDict()
        self._dims: Dict[int, Tuple[int, int]] = {}

    def dims(self, note: Note) -> Tuple[int, int]:
        """Thumbnail size for layout; reads only the PNG header."""
        key = note.attachment_id
        if key is None:
            return THUMB_MAX_W, 40
        if key not in self._dims:
            size = None
            try:
                data = note_store.image_bytes(note)
                size = _png_size(data or b'')
                if size is None and data:
                    image = QImage.fromData(data)
                    size = (image.width(), image.height()) if not image.isNull() else None
            except Exception:
                logging.exception(f'failed to read image size for note {note.id}')
            self._dims[key] = thumb_size(*size) if size else (THUMB_MAX_W, 40)
        return self._dims[key]

    def get(self, note: Note) -> Optional[QPixmap]:
        key = note.attachment_id
        if key is None:
            return None
        if key in self._items:
            self._items.move_to_end(key)
            return self._items[key]
        pixmap = None
        try:
            data = note_store.image_bytes(note)
            image = QImage.fromData(data) if data else QImage()
            if not image.isNull():
                pixmap = QPixmap.fromImage(
                    image.scaled(THUMB_MAX_W, THUMB_MAX_H, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                )
        except Exception:
            logging.exception(f'failed to load thumbnail for note {note.id}')
        self._items[key] = pixmap
        while len(self._items) > self._size:
            self._items.popitem(last=False)
        return pixmap


class NoteBubbleDelegate(QStyledItemDelegate):
    MARGIN_X = 14
    BUBBLE_PAD = 10
    META_H = 18
    HEADER_H = 30
    ROW_GAP = 8

    def __init__(self, model: NoteListModel, parent=None) -> None:
        super().__init__(parent)
        self._model = model
        self._thumbs = _ThumbCache()
        self._size_cache: Dict[Tuple[int, int, Optional[str], Tuple[str, ...]], int] = {}

    def _bubble_width(self, option: QStyleOptionViewItem) -> int:
        return max(120, option.rect.width() - 2 * self.MARGIN_X)

    def _document(self, note: Note, width: int) -> QTextDocument:
        text, truncated = _clip_text(note.text)
        doc = QTextDocument()
        doc.setDefaultStyleSheet(f'body {{ color: {TEXT_COLOR_HEX}; }}')
        doc.setDocumentMargin(0)
        body = highlight_html(text, self._model.terms)
        if truncated:
            body += f'<span style="color:{SUBTEXT_COLOR_HEX};"> …</span>'
        doc.setHtml(f'<body>{body}</body>')
        doc.setTextWidth(width - 2 * self.BUBBLE_PAD)
        return doc

    def sizeHint(self, option: QStyleOptionViewItem, index: QModelIndex) -> QSize:
        note: Optional[Note] = index.data(NOTE_ROLE)
        width = option.rect.width() or 480
        if note is None:
            return QSize(width, self.HEADER_H)
        if note.kind == note_store.KIND_IMAGE:
            _w, h = self._thumbs.dims(note)
            return QSize(width, self.META_H + 2 + h + 2 * self.BUBBLE_PAD + self.ROW_GAP)
        key = (note.id, width, note.updated_at, tuple(self._model.terms))
        height = self._size_cache.get(key)
        if height is None:
            doc = self._document(note, self._bubble_width(option))
            height = int(doc.size().height()) + 2 * self.BUBBLE_PAD + self.META_H + 2 + self.ROW_GAP
            if len(self._size_cache) > 5000:
                self._size_cache.clear()
            self._size_cache[key] = height
        return QSize(width, height)

    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex) -> None:
        painter.save()
        try:
            note: Optional[Note] = index.data(NOTE_ROLE)
            rect = option.rect
            if note is None:
                self._paint_header(painter, rect, index.data(HEADER_ROLE) or '')
                return
            selected = bool(option.state & QStyle.State_Selected)
            meta_rect = QRect(rect.left() + self.MARGIN_X, rect.top() + 2, rect.width() - 2 * self.MARGIN_X, self.META_H)
            painter.setPen(QColor(SUBTEXT_COLOR_HEX))
            small = QFont(option.font)
            small.setPointSizeF(max(7.0, option.font.pointSizeF() - 2))
            painter.setFont(small)
            painter.drawText(meta_rect, Qt.AlignLeft | Qt.AlignVCenter, note_meta(note))

            top = rect.top() + self.META_H + 2
            width = self._bubble_width(option)
            if note.kind == note_store.KIND_IMAGE:
                pixmap = self._thumbs.get(note)
                content_w, content_h = self._thumbs.dims(note)
                bubble = QRectF(rect.left() + self.MARGIN_X, top,
                                content_w + 2 * self.BUBBLE_PAD, content_h + 2 * self.BUBBLE_PAD)
                self._paint_bubble(painter, bubble, selected)
                if pixmap:
                    painter.drawPixmap(
                        QRect(int(bubble.left()) + self.BUBBLE_PAD, int(bubble.top()) + self.BUBBLE_PAD, content_w, content_h),
                        pixmap,
                    )
                else:
                    painter.setPen(QColor(SUBTEXT_COLOR_HEX))
                    painter.drawText(bubble, Qt.AlignCenter, '图片无法显示')
                return

            doc = self._document(note, width)
            ideal = min(doc.idealWidth(), width - 2 * self.BUBBLE_PAD)
            doc.setTextWidth(max(ideal, 40))
            bubble = QRectF(rect.left() + self.MARGIN_X, top,
                            doc.size().width() + 2 * self.BUBBLE_PAD,
                            doc.size().height() + 2 * self.BUBBLE_PAD)
            self._paint_bubble(painter, bubble, selected)
            painter.translate(bubble.left() + self.BUBBLE_PAD, bubble.top() + self.BUBBLE_PAD)
            ctx = QAbstractTextDocumentLayout.PaintContext()
            ctx.palette.setColor(QPalette.Text, QColor(TEXT_COLOR_HEX))
            doc.documentLayout().draw(painter, ctx)
        finally:
            painter.restore()

    def _paint_header(self, painter: QPainter, rect: QRect, label: str) -> None:
        painter.setPen(QColor(SUBTEXT_COLOR_HEX))
        fm = QFontMetrics(painter.font())
        text_w = fm.horizontalAdvance(label) + 24
        mid_y = rect.center().y()
        left = rect.left() + self.MARGIN_X
        right = rect.right() - self.MARGIN_X
        text_left = rect.center().x() - text_w // 2
        painter.setPen(QColor('#3A3A3A'))
        painter.drawLine(left, mid_y, max(left, text_left), mid_y)
        painter.drawLine(min(right, text_left + text_w), mid_y, right, mid_y)
        painter.setPen(QColor(SUBTEXT_COLOR_HEX))
        painter.drawText(QRect(text_left, rect.top(), text_w, rect.height()), Qt.AlignCenter, label)

    @staticmethod
    def _paint_bubble(painter: QPainter, bubble: QRectF, selected: bool) -> None:
        painter.setRenderHint(QPainter.Antialiasing, True)
        path = QPainterPath()
        path.addRoundedRect(bubble, 8, 8)
        painter.fillPath(path, QColor('#1B3A6B') if selected else QColor(PANEL_COLOR_HEX))
        if selected:
            painter.setPen(QColor(ACCENT_COLOR_HEX))
            painter.drawPath(path)


# ---------- actions shared by the panel and the day view ----------

def copy_to_clipboard(note: Note) -> bool:
    """Put the note's text or image back on the system clipboard."""
    clipboard = QGuiApplication.clipboard()
    if note.kind == note_store.KIND_IMAGE:
        data = note_store.image_bytes(note)
        image = QImage.fromData(data) if data else QImage()
        if image.isNull():
            return False
        clipboard.setImage(image)
        return True
    clipboard.setText(note.text)
    return True


class NotePreviewDialog(QDialog):
    def __init__(self, note: Note, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f'{day_label(note.plan_date)}  {note_meta(note)}')
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        layout = QVBoxLayout(self)
        screen = (parent.screen() if parent is not None else None) or QGuiApplication.primaryScreen()
        avail = screen.availableGeometry()
        if note.kind == note_store.KIND_IMAGE:
            data = note_store.image_bytes(note)
            pixmap = QPixmap()
            if data:
                pixmap.loadFromData(data)
            label = QLabel()
            label.setAlignment(Qt.AlignCenter)
            max_w, max_h = int(avail.width() * 0.9), int(avail.height() * 0.85)
            if not pixmap.isNull() and (pixmap.width() > max_w or pixmap.height() > max_h):
                pixmap = pixmap.scaled(max_w, max_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            label.setPixmap(pixmap)
            scroll = QScrollArea()
            scroll.setWidget(label)
            scroll.setWidgetResizable(True)
            layout.addWidget(scroll)
            self.resize(min(max_w, pixmap.width() + 40), min(max_h, pixmap.height() + 80))
        else:
            edit = QPlainTextEdit(note.text)
            edit.setReadOnly(True)
            layout.addWidget(edit)
            self.resize(min(900, int(avail.width() * 0.6)), min(640, int(avail.height() * 0.6)))
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class NoteEditDialog(QDialog):
    def __init__(self, note: Note, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle('修改文字')
        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        layout = QVBoxLayout(self)
        self.editor = QPlainTextEdit(note.text)
        layout.addWidget(self.editor)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(640, 400)

    def text(self) -> str:
        return self.editor.toPlainText()


# ---------- view ----------

class NoteListView(QListView):
    """List + keyboard / context-menu actions. Emits:
    - `copied(note)` after a note went to the clipboard (Enter / double click)
    - `older_requested()` when scrolled to the top (load an older page)
    - `open_day_requested(plan_date)`
    - `notes_changed()` after an edit or delete
    - `dialog_open(bool)` around modal dialogs (the panel must not auto-hide)."""

    copied = Signal(object)
    older_requested = Signal()
    open_day_requested = Signal(str)
    notes_changed = Signal()
    dialog_open = Signal(bool)

    def __init__(self, parent: Optional[QWidget] = None, today_fn: Callable[[], str] = lambda: time.strftime('%Y-%m-%d')) -> None:
        super().__init__(parent)
        self.note_model = NoteListModel(self, today_fn=today_fn)
        self.setModel(self.note_model)
        self.setItemDelegate(NoteBubbleDelegate(self.note_model, self))
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setResizeMode(QListView.Adjust)
        self.setLayoutMode(QListView.Batched)
        self.setBatchSize(100)
        self.setUniformItemSizes(False)
        self.setWordWrap(True)
        self.setMouseTracking(False)
        self.setStyleSheet('QListView { border: none; padding: 0; }')
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_menu)
        self.doubleClicked.connect(lambda _idx: self.copy_current())
        self.verticalScrollBar().valueChanged.connect(self._on_scroll)
        self.allow_open_day = True
        self._suppress_older = False
        self._top_gap = 0
        for sig in (self.note_model.modelReset, self.note_model.rowsInserted, self.note_model.rowsRemoved,
                    self.note_model.dataChanged):
            sig.connect(self._schedule_align)

    # -- bottom alignment: few notes sit next to the search box, like a chat --
    def _schedule_align(self, *_args) -> None:
        QTimer.singleShot(0, self._align_bottom)

    def content_height(self, limit: int) -> int:
        option = QStyleOptionViewItem()
        option.rect = QRect(0, 0, self.viewport().width(), 0)
        delegate = self.itemDelegate()
        total = 0
        for row in range(self.note_model.rowCount() - 1, -1, -1):
            total += delegate.sizeHint(option, self.note_model.index(row)).height()
            if total >= limit:
                break
        return total

    def _align_bottom(self) -> None:
        available = self.height() - self.frameWidth() * 2
        gap = max(0, available - self.content_height(available))
        if gap != self._top_gap:
            self._top_gap = gap
            self.setViewportMargins(0, gap, 0, 0)
            if self.currentIndex().isValid():
                self.scrollTo(self.currentIndex(), QAbstractItemView.EnsureVisible)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._schedule_align()

    # -- selection --
    def current_note(self) -> Optional[Note]:
        index = self.currentIndex()
        return index.data(NOTE_ROLE) if index.isValid() else None

    def select_row(self, row: int) -> None:
        if 0 <= row < self.note_model.rowCount():
            index = self.note_model.index(row)
            self.setCurrentIndex(index)
            self.scrollTo(index, QAbstractItemView.EnsureVisible)

    def select_last(self) -> None:
        rows = self.note_model.note_rows()
        if rows:
            self.select_row(rows[-1])
        self.scrollToBottom()

    def select_note(self, note_id: int) -> None:
        self.select_row(self.note_model.row_of_note(note_id))

    def move_selection(self, delta: int) -> None:
        rows = self.note_model.note_rows()
        if not rows:
            return
        current = self.currentIndex().row()
        if current not in rows:
            target = rows[-1] if delta < 0 else rows[0]
        else:
            pos = max(0, min(len(rows) - 1, rows.index(current) + delta))
            target = rows[pos]
        self.select_row(target)
        if target == rows[0] and delta < 0:
            self.older_requested.emit()

    # -- loading --
    def set_notes(self, notes: List[Note], terms: Optional[List[str]] = None, select_last: bool = True) -> None:
        self._suppress_older = True
        self.note_model.set_notes(notes, terms)
        if select_last:
            self.select_last()
        self._suppress_older = False

    def prepend_notes(self, older: List[Note]) -> None:
        """Insert older notes above without moving what the user is looking at."""
        if not older:
            return
        self._suppress_older = True
        anchor = self.indexAt(self.viewport().rect().topLeft())
        anchor_note = anchor.data(NOTE_ROLE) if anchor.isValid() else None
        anchor_offset = self.visualRect(anchor).top() if anchor.isValid() else 0
        current = self.current_note()
        self.note_model.prepend_notes(older)
        if current is not None:
            self.setCurrentIndex(self.note_model.index(self.note_model.row_of_note(current.id)))
        if anchor_note is not None:
            row = self.note_model.row_of_note(anchor_note.id)
            if row >= 0:
                self.scrollTo(self.note_model.index(row), QAbstractItemView.PositionAtTop)
                self.verticalScrollBar().setValue(self.verticalScrollBar().value() - anchor_offset)
        self._suppress_older = False

    def _on_scroll(self, value: int) -> None:
        if not self._suppress_older and value <= self.verticalScrollBar().minimum() + 4:
            self.older_requested.emit()

    # -- actions --
    def copy_current(self) -> None:
        note = self.current_note()
        if note is None:
            return
        try:
            if copy_to_clipboard(note):
                self.copied.emit(note)
        except Exception:
            logging.exception('failed to copy note to clipboard')

    def preview_current(self) -> None:
        note = self.current_note()
        if note is None:
            return
        self.dialog_open.emit(True)
        try:
            NotePreviewDialog(note, self.window()).exec()
        finally:
            self.dialog_open.emit(False)

    def edit_current(self) -> None:
        note = self.current_note()
        if note is None or note.kind != note_store.KIND_TEXT:
            return
        self.dialog_open.emit(True)
        try:
            dialog = NoteEditDialog(note, self.window())
            if dialog.exec() != QDialog.Accepted:
                return
            text = dialog.text().strip()
            if not text or text == note.text:
                return
            note_store.update_text(note.id, text)
            updated = note_store.get(note.id)
            if updated is not None:
                self.note_model.replace_note(updated)
            self.notes_changed.emit()
        except Exception:
            logging.exception('failed to update note')
        finally:
            self.dialog_open.emit(False)

    def delete_current(self) -> None:
        note = self.current_note()
        if note is None:
            return
        self.dialog_open.emit(True)
        try:
            preview = '截图' if note.kind == note_store.KIND_IMAGE else _clip_text(note.text)[0][:80]
            answer = QMessageBox.question(
                self.window(), '删除记录', f'删除这条记录？\n\n{preview}',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if answer != QMessageBox.Yes:
                return
            row = self.currentIndex().row()
            note_store.delete(note.id)
            self.note_model.remove_note(note.id)
            rows = self.note_model.note_rows()
            if rows:
                self.select_row(min(rows, key=lambda r: abs(r - row)))
            self.notes_changed.emit()
        except Exception:
            logging.exception('failed to delete note')
        finally:
            self.dialog_open.emit(False)

    def _show_menu(self, pos) -> None:
        index = self.indexAt(pos)
        note = index.data(NOTE_ROLE) if index.isValid() else None
        if note is None:
            return
        self.setCurrentIndex(index)
        menu = QMenu(self)
        actions = [
            ('复制', self.copy_current, True),
            ('预览', self.preview_current, True),
            ('修改文字', self.edit_current, note.kind == note_store.KIND_TEXT),
            ('删除', self.delete_current, True),
        ]
        for label, slot, enabled in actions:
            action = QAction(label, menu)
            action.setEnabled(enabled)
            action.triggered.connect(slot)
            menu.addAction(action)
        if self.allow_open_day:
            menu.addSeparator()
            action = QAction('在日视图中查看', menu)
            action.triggered.connect(lambda: self.open_day_requested.emit(note.plan_date))
            menu.addAction(action)
        self.dialog_open.emit(True)
        try:
            menu.exec(self.viewport().mapToGlobal(pos))
        finally:
            self.dialog_open.emit(False)
