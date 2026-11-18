"""Disk-backed browsing of immutable report results, including legacy JSON runs."""

import json
import os
import sqlite3
import threading
import uuid
from contextlib import closing
from pathlib import Path


_legacy_lock = threading.Lock()


def write_result_store(directory: Path, rows: list[dict], columns: list[str] | None = None) -> None:
    columns = columns if columns is not None else list(rows[0]) if rows else []
    target = directory / 'data.sqlite'
    temporary = directory / f'data-{uuid.uuid4().hex}.tmp'
    try:
        with closing(sqlite3.connect(temporary)) as conn, conn:
            conn.execute('CREATE TABLE metadata (columns_json TEXT, total INTEGER)')
            conn.execute('INSERT INTO metadata VALUES (?, ?)', (json.dumps(columns), len(rows)))
            conn.execute('CREATE TABLE result_rows (row_number INTEGER PRIMARY KEY, data TEXT, search_text TEXT)')
            conn.executemany('INSERT INTO result_rows VALUES (?, ?, ?)', (
                (index, json.dumps([row.get(c) for c in columns], ensure_ascii=False, default=str),
                 '\n'.join(str(row.get(c, '')) for c in columns).casefold())
                for index, row in enumerate(rows)
            ))
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)


def page_options(query: dict) -> tuple[int, int]:
    try:
        page = int(query.get('page', ['1'])[0])
        size = int(query.get('page_size', ['50'])[0])
    except (ValueError, TypeError) as exc:
        raise ValueError('页码和每页条数必须是整数') from exc
    if page < 1 or not 1 <= size <= 500:
        raise ValueError('页码至少为 1，每页条数为 1 至 500')
    return page, size


def result_page(directory: Path, query: dict) -> dict:
    page, size = page_options(query)
    target = directory / 'data.sqlite'
    if not target.exists():
        with _legacy_lock:
            if not target.exists():
                source = directory / 'data.json'
                if not source.exists():
                    raise ValueError('此运行尚无可查询的结果')
                write_result_store(directory, json.loads(source.read_text(encoding='utf-8')))
    with closing(sqlite3.connect(target.as_uri() + '?mode=ro', uri=True)) as conn:
        metadata = conn.execute('SELECT columns_json, total FROM metadata').fetchone()
        columns, total = json.loads(metadata[0]), metadata[1]
        search = str(query.get('q', [''])[0]).strip().casefold()
        sort = query.get('sort', [''])[0]
        direction = query.get('direction', ['asc'])[0]
        if direction not in {'asc', 'desc'} or (sort and sort not in columns):
            raise ValueError('无效的排序字段或方向')
        where = 'WHERE instr(search_text, ?) > 0' if search else ''
        params = [search] if search else []
        count = conn.execute(f'SELECT COUNT(*) FROM result_rows {where}', params).fetchone()[0] if search else total
        page = min(page, max(1, (count + size - 1) // size))
        order = f"json_extract(data, '$[{columns.index(sort)}]') {direction}, row_number" if sort else 'row_number'
        offset = (page - 1) * size
        if not search and not sort:
            cursor = conn.execute('SELECT data FROM result_rows WHERE row_number>=? ORDER BY row_number LIMIT ?', (offset, size))
        else:
            cursor = conn.execute(f'SELECT data FROM result_rows {where} ORDER BY {order} LIMIT ? OFFSET ?', [*params, size, offset])
        rows = [dict(zip(columns, json.loads(row[0]))) for row in cursor]
    return dict(items=rows, columns=columns, total=total, filtered=count, page=page, page_size=size)
