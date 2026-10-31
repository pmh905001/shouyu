"""The alt+/ notes panel, driven offscreen against a temp database
(REC-01 .. REC-16)."""
import io
import os
import time

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

import pytest
from PIL import Image
from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QMessageBox

from shouyu.service import note_store
from shouyu.view import note_list
from shouyu.view.msgbox import MessageBox


def _day(offset):
    return time.strftime('%Y-%m-%d', time.localtime(time.time() + offset * 86400))


TODAY, YESTERDAY, OLDER = _day(0), _day(-1), _day(-10)


def _png():
    buf = io.BytesIO()
    Image.new('RGB', (300, 120), (20, 120, 200)).save(buf, format='PNG')
    return buf.getvalue()


def _add(conn, plan_date, text):
    return conn.execute(
        "INSERT INTO note (plan_date, kind, text, source, created_at) VALUES (?, 'text', ?, 'clipboard', ?)",
        (plan_date, text, '2026-10-31T02:00:00.000Z'),
    ).lastrowid


@pytest.fixture
def qt(db, v2env, monkeypatch):
    app = QApplication.instance() or QApplication([])
    toasts = []
    monkeypatch.setattr(MessageBox, 'pop_up_message', classmethod(lambda cls, *a, **kw: toasts.append(kw or a)))
    yield toasts
    app.processEvents()


@pytest.fixture
def panel(qt):
    from shouyu.view.notes_panel import NotesPanel

    p = NotesPanel()
    p._hide_if_inactive = lambda: None  # offscreen windows never become "active"
    yield p
    p._searcher.close()
    p.deleteLater()


def _wait_search(panel, timeout=10):
    deadline = time.time() + timeout
    while panel._searching and time.time() < deadline:
        QApplication.processEvents()
        time.sleep(0.01)
    QApplication.processEvents()
    assert not panel._searching


def _shown(panel):
    return [n.text for n in panel.list_view.note_model.notes()]


def _type(panel, text):
    panel.search_box.setText(text)
    panel._run_search()  # skip the debounce timer
    _wait_search(panel)


def test_open_shows_newest_last_and_selected(panel, db):
    _add(db, OLDER, '十天前')
    _add(db, YESTERDAY, '昨天的数据')
    _add(db, TODAY, '今天第一条')
    _add(db, TODAY, '今天最后一条')
    db.commit()
    panel.open()
    assert _shown(panel) == ['十天前', '昨天的数据', '今天第一条', '今天最后一条']
    assert panel.list_view.current_note().text == '今天最后一条'
    assert panel.search_box.hasFocus() or not panel.isActiveWindow()
    assert panel.empty_label.isHidden()
    assert panel.count_label.text() == '4 条'
    # Day headers are rows of their own, not selectable.
    model = panel.list_view.note_model
    headers = [model.index(r).data(note_list.HEADER_ROLE) for r in range(model.rowCount())]
    assert [h for h in headers if h] == [note_list.day_label(OLDER), note_list.day_label(YESTERDAY), '今天']


def test_empty_today_hint(panel, db):
    _add(db, YESTERDAY, '昨天的')
    db.commit()
    panel.open()
    assert panel.empty_label.text() == '今天还没有记录，往上滚动查看以前的'


def test_instant_search_scopes_and_highlight(panel, db):
    _add(db, OLDER, '数据迁移 旧')
    _add(db, YESTERDAY, '昨天的数据')
    _add(db, TODAY, '今天的数据')
    _add(db, TODAY, '无关')
    db.commit()
    panel.open()

    _type(panel, '数')
    assert _shown(panel) == ['数据迁移 旧', '昨天的数据', '今天的数据']
    assert panel.list_view.current_note().text == '今天的数据'  # closest to the search box
    assert panel.list_view.note_model.terms == ['数']
    assert panel.count_label.text() == '3 条'

    panel.set_scope('yesterday')
    _wait_search(panel)
    assert _shown(panel) == ['昨天的数据']
    panel.set_scope('today')
    _wait_search(panel)
    assert _shown(panel) == ['今天的数据']
    panel._cycle_scope(1)  # today -> yesterday
    _wait_search(panel)
    assert panel._scope == 'yesterday'
    panel.set_scope('all')
    _wait_search(panel)

    _type(panel, '数据 旧')
    assert _shown(panel) == ['数据迁移 旧']
    _type(panel, '没有这个词')
    assert _shown(panel) == []
    assert panel.empty_label.text() == '没有找到匹配的记录'

    # Clearing the box goes back to browsing.
    _type(panel, '')
    assert len(_shown(panel)) == 4


def test_keyboard_select_copy_and_hide(panel, db, qt):
    _add(db, TODAY, '第一条')
    _add(db, TODAY, '第二条')
    db.commit()
    panel.open()
    QTest.keyClick(panel.search_box, Qt.Key_Up)
    assert panel.list_view.current_note().text == '第一条'
    QTest.keyClick(panel.search_box, Qt.Key_Down)
    assert panel.list_view.current_note().text == '第二条'
    QTest.keyClick(panel.search_box, Qt.Key_Return)
    assert QGuiApplication.clipboard().text() == '第二条'
    assert not panel.isVisible()
    assert qt[-1] == {'title': '已复制', 'msg': '第二条'}

    panel.open()
    QTest.keyClick(panel.search_box, Qt.Key_Escape)
    assert not panel.isVisible()


def test_copy_image_note(panel, db):
    note_store.save(image_png=_png())
    panel.open()
    panel.list_view.copy_current()
    image = QGuiApplication.clipboard().image()
    assert (image.width(), image.height()) == (300, 120)


def test_paging_back_through_older_notes(panel, db):
    db.execute('BEGIN')
    for i in range(450):
        _add(db, _day(-1 - i // 100), f'旧记录 {i}')
    db.execute('COMMIT')
    panel.open()
    assert len(_shown(panel)) == note_store.PAGE_SIZE
    assert panel.count_label.text() == '已加载 200 条，往上滚动加载更多'
    panel._load_older()
    assert len(_shown(panel)) == 400
    panel._load_older()
    assert len(_shown(panel)) == 450
    assert panel._no_more_older
    assert panel.count_label.text() == '450 条'
    panel._load_older()  # nothing more, no error
    assert len(_shown(panel)) == 450


def test_delete_and_edit_from_the_panel(panel, db, monkeypatch):
    keep = _add(db, TODAY, '留着')
    _add(db, TODAY, '删掉我')
    db.commit()
    panel.open()
    monkeypatch.setattr(QMessageBox, 'question', staticmethod(lambda *a, **kw: QMessageBox.Yes))
    QTest.keyClick(panel.search_box, Qt.Key_Delete)  # empty box: Delete removes the selected note
    assert _shown(panel) == ['留着']
    assert [n.text for n in note_store.list_day(TODAY)] == ['留着']

    class _FakeEdit:
        def __init__(self, *a, **kw):
            pass

        def exec(self):
            return note_list.QDialog.Accepted

        def text(self):
            return '改过了'

    monkeypatch.setattr(note_list, 'NoteEditDialog', _FakeEdit)
    panel.list_view.select_note(keep)
    panel.list_view.edit_current()
    assert _shown(panel) == ['改过了']
    assert note_store.get(keep).text == '改过了'
    assert panel._dialogs_open == 0


def test_rendering_text_and_image_bubbles(panel, db):
    """Paint every row (text, long text, image, headers) without errors."""
    _add(db, YESTERDAY, '短')
    _add(db, TODAY, '\n'.join(f'第 {i} 行 <b>不是标签</b>' for i in range(20)))
    note_store.save(image_png=_png())
    panel.open()
    _type(panel, '行')
    panel.resize(580, 740)
    pixmap = panel.grab()
    assert not pixmap.isNull()


def test_formatting_helpers():
    assert note_list.day_label('2026-10-31', '2026-10-31') == '今天'
    assert note_list.day_label('2026-10-30', '2026-10-31') == '昨天 10-30'
    assert note_list.day_label('2026-10-29', '2026-10-31') == '前天 10-29'
    assert note_list.day_label('2026-10-27', '2026-10-31') == '10-27 星期二'
    assert note_list.day_label('2025-12-31', '2026-10-31') == '2025-12-31 星期三'
    html = note_list.highlight_html('Hello <x> hello', ['hello'])
    assert html.count('<span') == 2 and '&lt;x&gt;' in html
    assert note_list.highlight_html('a\nb', []) == 'a<br>b'


def test_scope_range():
    from shouyu.view.notes_panel import scope_range

    assert scope_range('today', '2026-10-31') == ('2026-10-31', '2026-10-31')
    assert scope_range('yesterday', '2026-10-31') == ('2026-10-30', '2026-10-30')
    assert scope_range('day_before', '2026-10-31') == ('2026-10-29', '2026-10-29')
    assert scope_range('week', '2026-10-31') == ('2026-10-25', '2026-10-31')
    assert scope_range('all', '2026-10-31') == (None, None)
