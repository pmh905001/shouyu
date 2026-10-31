import os
import sys

import pytest

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from shouyu.config import Config  # noqa: E402
from shouyu.util.state import AppState  # noqa: E402


@pytest.fixture
def v2env(tmp_path, monkeypatch):
    """Point the database, the Excel file and the state file at a temp dir,
    so tests never touch real data. Mirroring to Excel is on by default."""
    db_path = str(tmp_path / 'shouyu_v2.db')
    excel_path = str(tmp_path / 'kb.xlsx')
    settings = {'excel_mirror': True}
    monkeypatch.setattr(Config, 'sqlite_db_path', classmethod(lambda cls: db_path))
    monkeypatch.setattr(Config, 'excel_path', classmethod(lambda cls: excel_path))
    monkeypatch.setattr(Config, 'excel_mirror', classmethod(lambda cls: settings['excel_mirror']))
    monkeypatch.setattr(AppState, 'FILE_NAME', str(tmp_path / 'shouyu_state.json'))
    return {'db_path': db_path, 'excel_path': excel_path, 'settings': settings, 'dir': tmp_path}


@pytest.fixture
def db(v2env):
    """An open read connection to the temp database (schema initialized)."""
    from shouyu.service import datastore

    datastore.init_schema()
    conn = datastore._open()
    yield conn
    conn.close()


def queue_rows(conn, kind=None):
    import json

    sql = "SELECT id, kind, payload, attachment_id, status FROM queue"
    params = ()
    if kind:
        sql += " WHERE kind = ?"
        params = (kind,)
    rows = conn.execute(sql + " ORDER BY id", params).fetchall()
    return [dict(r, payload=json.loads(r['payload'])) for r in rows]
