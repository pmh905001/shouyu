import io

from PIL import Image

from conftest import queue_rows
from shouyu.service import message_queue, note_store


def _png() -> bytes:
    buf = io.BytesIO()
    Image.new('RGB', (40, 20), (200, 30, 30)).save(buf, format='PNG')
    return buf.getvalue()


def test_save_text_and_image_then_last(db):
    assert note_store.last() is None

    text_note = note_store.save(text='第一段文字', excel_column='B')
    assert note_store.last().id == text_note.id
    assert note_store.last().text == '第一段文字'

    png = _png()
    image_note = note_store.save(image_png=png, excel_column='A', source=note_store.SOURCE_OCR)
    last = note_store.last()
    assert last.id == image_note.id
    assert (last.kind, last.excel_column, last.source, last.text) == ('image', 'A', 'ocr', '')
    assert note_store.image_bytes(last) == png

    msgs = queue_rows(db, 'clipboard_append')
    assert msgs[0]['payload'] == {'column': 'B', 'note_id': text_note.id, 'text': '第一段文字'}
    assert msgs[1]['payload'] == {'column': 'A', 'note_id': image_note.id}
    # The note and its mirror message share one attachment row (no copy).
    assert msgs[1]['attachment_id'] == image_note.attachment_id
    assert db.execute("SELECT COUNT(*) FROM attachments").fetchone()[0] == 1


def test_no_mirror_when_disabled(db, v2env):
    v2env['settings']['excel_mirror'] = False
    note_store.save(text='only sqlite')
    assert queue_rows(db, 'clipboard_append') == []
    assert note_store.last().text == 'only sqlite'


def test_purge_keeps_attachments_still_used_by_notes(db):
    note = note_store.save(image_png=_png())
    msg = queue_rows(db, 'clipboard_append')[0]
    message_queue.mark_done([msg['id']])
    with message_queue._connect() as conn:
        conn.execute("UPDATE queue SET created_at = '2000-01-01T00:00:00.000Z'")

    assert message_queue.purge_resolved(older_than_days=1) == 1
    assert queue_rows(db) == []
    assert note_store.image_bytes(note_store.last()) is not None
    assert db.execute("SELECT id FROM attachments").fetchone()[0] == note.attachment_id


def test_show_status_reads_last_note(db, monkeypatch):
    from shouyu.action.shortcut import Shortcut
    from shouyu.view.msgbox import MessageBox

    shown = []
    monkeypatch.setattr(MessageBox, 'pop_up_message', classmethod(lambda cls, **kw: shown.append(kw)))

    Shortcut.show_status()
    assert shown[-1]['msg'] == '还没有记录'

    note_store.save(text='x' * 300)
    Shortcut.show_status()
    assert shown[-1]['title'].startswith('上一条记录 · ')
    assert shown[-1]['msg'] == 'x' * 100

    png = _png()
    note_store.save(image_png=png)
    Shortcut.show_status()
    assert shown[-1]['msg'] == '截图'
    with open(shown[-1]['image_path'], 'rb') as f:
        assert f.read() == png


def test_clipboard_hotkey_saves_note(db, monkeypatch):
    from shouyu.action import shortcut as shortcut_module
    from shouyu.service import dispatch

    monkeypatch.setattr(shortcut_module.ImageGrab, 'grabclipboard', lambda: None)
    monkeypatch.setattr(shortcut_module.pyperclip, 'paste', lambda: '剪贴板内容')
    toasts = []
    monkeypatch.setattr(dispatch, 'notify_recorded', lambda title, msg: toasts.append((title, msg)))

    shortcut_module.Shortcut.save_clipboard_to_column('A')

    last = note_store.last()
    assert (last.text, last.excel_column, last.source) == ('剪贴板内容', 'A', 'clipboard')
    assert toasts == [('已记录', '剪贴板内容')]


def test_format_note_time():
    from shouyu.action.shortcut import Shortcut

    assert Shortcut._format_note_time('2020-01-02T03:04:05.000Z').startswith('2020-01-02 ')
    assert Shortcut._format_note_time('garbage') == 'garbage'
