"""note_store paging / search / edit / delete and the v1 -> v2 note migration
(docs/design/retrieval/20261031-01-notes-panel-and-day-view.md §3)."""
import sqlite3
import threading

import pytest

from conftest import queue_rows
from shouyu.service import datastore, message_queue, note_store
from shouyu.service.note_search import NoteSearcher, split_range
from shouyu.service.note_store import NoteQuery


def _add(conn, plan_date, text, kind='text', source='clipboard'):
    return conn.execute(
        "INSERT INTO note (plan_date, kind, text, source, created_at) VALUES (?, ?, ?, ?, ?)",
        (plan_date, kind, text, source, '2026-10-31T00:00:00.000Z'),
    ).lastrowid


@pytest.fixture
def notes(db):
    ids = {
        'old': _add(db, '2026-10-01', 'medis 账号 old'),
        'a': _add(db, '2026-10-29', '数据迁移 HandleChat'),
        'b': _add(db, '2026-10-29', 'medis 账号申请 xiaokang'),
        'img': _add(db, '2026-10-30', '', kind='image'),
        'c': _add(db, '2026-10-30', '100% done_ok'),
        'd': _add(db, '2026-10-31', '今天的数据'),
    }
    db.commit()
    return ids


def _texts(notes):
    return [n.text for n in notes]


def test_migration_keeps_v1_notes_and_their_ids(v2env):
    """A database created with schema v1 (the previous release) upgrades in place."""
    message_queue.init_db()
    conn = sqlite3.connect(v2env['db_path'])
    conn.execute("CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
    for statement in datastore._MIGRATIONS[1]:
        conn.execute(statement)
    conn.execute("INSERT INTO meta VALUES ('schema_version', '1')")
    conn.execute(
        "INSERT INTO note (id, plan_date, kind, text, source, excel_column, created_at) "
        "VALUES (7, '2026-10-31', 'text', 'v1 note', 'clipboard', 'B', '2026-10-31T01:00:00.000Z')"
    )
    conn.commit()
    conn.close()

    datastore.init_schema()
    note = note_store.get(7)
    assert (note.text, note.excel_column, note.updated_at) == ('v1 note', 'B', None)
    with datastore.reader() as c:
        assert datastore.get_meta(c, datastore.META_SCHEMA_VERSION) == str(datastore.SCHEMA_VERSION)
        cols = {r['name'] for r in c.execute("PRAGMA table_info(note)")}
    assert {'source_ref', 'updated_at', 'deleted_at'} <= cols
    note_store.add_manual('2026-10-31', 'manual source is allowed now')


def test_page_before_walks_back_in_pages(notes):
    first = note_store.page_before('2026-11-01', 0, limit=2)
    assert _texts(first) == ['100% done_ok', '今天的数据']  # oldest first within the page
    second = note_store.page_before(first[0].plan_date, first[0].id, limit=3)
    assert [n.id for n in second] == [notes['a'], notes['b'], notes['img']]
    third = note_store.page_before(second[0].plan_date, second[0].id, limit=3)
    assert [n.id for n in third] == [notes['old']]
    assert note_store.page_before(third[0].plan_date, third[0].id) == []


def test_search_terms_case_and_ranges(notes):
    search = lambda *terms, **kw: _texts(note_store.search(NoteQuery(terms=list(terms), **kw)))
    assert search('数') == ['今天的数据', '数据迁移 HandleChat']  # newest first
    assert search('handlechat') == ['数据迁移 HandleChat']  # ASCII is case-insensitive
    assert search('medis', '账号') == ['medis 账号申请 xiaokang', 'medis 账号 old']
    assert search('medis', '申请') == ['medis 账号申请 xiaokang']
    assert search('medis', date_from='2026-10-29', date_to='2026-10-29') == ['medis 账号申请 xiaokang']
    # LIKE wildcards in the user's input are literal characters.
    assert search('%') == ['100% done_ok']
    assert search('_') == ['100% done_ok']
    assert search('d_ne') == []  # as a wildcard it would match "done"
    # No terms: browse, images included.
    assert len(note_store.search(NoteQuery(date_from='2026-10-30', date_to='2026-10-30'))) == 2
    assert len(note_store.search(NoteQuery(terms=['数'], limit=1))) == 1


def test_parse_terms():
    assert note_store.parse_terms('  medis   账号 ') == ['medis', '账号']
    assert note_store.parse_terms('') == []


def test_delete_and_update(notes, db):
    assert note_store.update_text(notes['a'], '数据迁移 改过')
    assert note_store.get(notes['a']).updated_at
    assert not note_store.update_text(notes['img'], 'images have no text')
    with pytest.raises(ValueError):
        note_store.update_text(notes['a'], '  ')

    assert note_store.delete(notes['d'])
    assert not note_store.delete(notes['d'])
    assert note_store.get(notes['d']) is None
    assert note_store.last().id == notes['c']
    assert '今天的数据' not in _texts(note_store.search(NoteQuery(terms=['数'])))
    assert notes['d'] not in [n.id for n in note_store.list_day('2026-10-31')]
    assert notes['d'] not in [n.id for n in note_store.page_before('2026-11-01', 0)]


def test_add_manual_is_not_mirrored(db):
    note = note_store.add_manual('2026-10-29', '  补一条  ')
    assert (note.text, note.source, note.plan_date) == ('补一条', 'manual', '2026-10-29')
    assert [n.text for n in note_store.list_day('2026-10-29')] == ['补一条']
    assert queue_rows(db) == []
    with pytest.raises(ValueError):
        note_store.add_manual('2026-10-29', ' ')


def test_split_range():
    q = lambda lo=None, hi=None: NoteQuery(terms=['x'], date_from=lo, date_to=hi)
    parts = lambda *a: [(p.date_from, p.date_to) for p in split_range(q(*a), '2026-10-31')]
    assert parts() == [('2026-10-29', None), (None, '2026-10-28')]
    assert parts('2026-10-31', '2026-10-31') == [('2026-10-31', '2026-10-31')]
    assert parts('2026-10-30', '2026-10-30') == [('2026-10-30', '2026-10-30')]
    assert parts('2026-10-25', '2026-10-31') == [('2026-10-29', '2026-10-31'), ('2026-10-25', '2026-10-28')]
    assert parts('2026-10-01', '2026-10-05') == [('2026-10-01', '2026-10-05')]


def _collect():
    results, done = [], threading.Event()

    def cb(generation, found, appended, finished):
        results.append((generation, [n.text for n in found], appended, finished))
        if finished:
            done.set()

    return results, done, cb


def test_searcher_reports_near_days_first(notes):
    searcher = NoteSearcher(today_fn=lambda: '2026-10-31')
    results, done, cb = _collect()
    searcher.submit(1, NoteQuery(terms=['medis']), cb)
    assert done.wait(10)
    searcher.close()
    assert results == [
        (1, ['medis 账号申请 xiaokang'], False, False),  # 10-29..10-31
        (1, ['medis 账号 old'], True, True),
    ]


def test_searcher_only_runs_the_newest_request(notes, monkeypatch):
    """A submit while a query runs interrupts it; the stale generation never finishes."""
    gate, entered = threading.Event(), threading.Event()
    original = note_store.search

    def slow_search(query, conn=None):
        if query.terms == ['slow']:
            entered.set()
            gate.wait(10)
            # The interrupt only lands on a running statement; emulate one.
            raise sqlite3.OperationalError('interrupted')
        return original(query, conn=conn)

    monkeypatch.setattr(note_store, 'search', slow_search)
    searcher = NoteSearcher(today_fn=lambda: '2026-10-31')
    results, done, cb = _collect()
    searcher.submit(1, NoteQuery(terms=['slow']), cb)
    assert entered.wait(10)
    searcher.submit(2, NoteQuery(terms=['数'], date_from='2026-10-31', date_to='2026-10-31'), cb)
    gate.set()
    assert done.wait(10)
    searcher.close()
    assert results == [(2, ['今天的数据'], False, True)]


def test_searcher_interrupts_a_real_running_statement(notes, monkeypatch):
    """conn.interrupt() really stops a statement running on the worker's own
    connection. The stand-in query never ends on its own."""
    import time

    entered = threading.Event()
    original = note_store.search

    def endless(query, conn=None):
        if query.terms == ['forever']:
            entered.set()
            conn.execute(
                "WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM c) SELECT COUNT(*) FROM c"
            ).fetchall()
        return original(query, conn=conn)

    monkeypatch.setattr(note_store, 'search', endless)
    searcher = NoteSearcher(today_fn=lambda: '2026-10-31')
    results, done, cb = _collect()
    searcher.submit(1, NoteQuery(terms=['forever']), cb)
    assert entered.wait(10)
    time.sleep(0.3)  # let the statement start running
    started = time.perf_counter()
    searcher.submit(2, NoteQuery(terms=['数'], date_from='2026-10-31', date_to='2026-10-31'), cb)
    assert done.wait(10)
    searcher.close()
    assert time.perf_counter() - started < 2
    assert results == [(2, ['今天的数据'], False, True)]
