"""Lightweight web report portal.

The portal intentionally uses only the Python standard library at runtime. It
provides a small, inspectable MVP for report definition, scheduled execution,
process control, result browsing, HTML output and XLSX export.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import os
import re
import sqlite3
import threading
import time
from datetime import datetime, timedelta, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
STATIC_DIR = ROOT / "static"
DB_PATH = Path(os.environ.get("REPORT_DB_PATH", str(DATA_DIR / "report.db"))).expanduser()
if not DB_PATH.is_absolute():
    DB_PATH = (Path.cwd() / DB_PATH).resolve()

UTC = timezone.utc
ALLOWED_CHARTS = {"table", "bar", "line", "pie"}
ALLOWED_FORMATS = {"html", "xlsx", "csv"}


def utc_now() -> datetime:
    return datetime.now(UTC).replace(microsecond=0)


def iso(value: datetime | None) -> str | None:
    return value.isoformat() if value else None


def parse_iso(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def json_value(value: Any) -> Any:
    if isinstance(value, (datetime,)):
        return value.isoformat()
    if isinstance(value, (bytes, bytearray)):
        return value.decode("utf-8", errors="replace")
    try:
        json.dumps(value)
        return value
    except TypeError:
        return str(value)


def connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def seed_demo_database(conn: sqlite3.Connection) -> None:
    """Create a small demo table only for a new, local default database."""

    if DB_PATH != (DATA_DIR / "report.db").resolve() or conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='demo_sales'"
    ).fetchone():
        return
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS demo_sales (
            id INTEGER PRIMARY KEY,
            month TEXT NOT NULL,
            region TEXT NOT NULL,
            amount REAL NOT NULL,
            orders INTEGER NOT NULL
        )
        """
    )
    if conn.execute("SELECT COUNT(*) FROM demo_sales").fetchone()[0] == 0:
        rows = [
            ("2026-01", "华东", 128000, 320),
            ("2026-01", "华南", 97000, 250),
            ("2026-02", "华东", 145000, 360),
            ("2026-02", "华南", 112000, 290),
            ("2026-03", "华东", 151000, 380),
            ("2026-03", "华南", 121000, 305),
            ("2026-04", "华东", 163000, 410),
            ("2026-04", "华南", 134000, 336),
        ]
        conn.executemany(
            "INSERT INTO demo_sales(month, region, amount, orders) VALUES (?, ?, ?, ?)", rows
        )


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with connection() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT NOT NULL DEFAULT '',
                sql TEXT NOT NULL,
                chart_type TEXT NOT NULL DEFAULT 'table',
                x_column TEXT NOT NULL DEFAULT '',
                y_column TEXT NOT NULL DEFAULT '',
                formulas TEXT NOT NULL DEFAULT '[]',
                layout TEXT NOT NULL DEFAULT '{}',
                output_formats TEXT NOT NULL DEFAULT '["html", "xlsx"]',
                schedule_minutes INTEGER NOT NULL DEFAULT 0,
                enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                last_run_at TEXT,
                next_run_at TEXT
            );
            CREATE TABLE IF NOT EXISTS runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id INTEGER NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                status TEXT NOT NULL,
                started_at TEXT NOT NULL,
                finished_at TEXT,
                row_count INTEGER NOT NULL DEFAULT 0,
                output_dir TEXT,
                error TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_runs_report_started ON runs(report_id, started_at DESC);
            """
        )
        # Keep existing installations compatible with the formula feature.
        columns = {row[1] for row in conn.execute("PRAGMA table_info(reports)").fetchall()}
        if "formulas" not in columns:
            conn.execute("ALTER TABLE reports ADD COLUMN formulas TEXT NOT NULL DEFAULT '[]'")
        if "layout" not in columns:
            conn.execute("ALTER TABLE reports ADD COLUMN layout TEXT NOT NULL DEFAULT '{}'")
        seed_demo_database(conn)
        demo_formulas = json.dumps([
            {"name": "最大值", "formula": "=MAX(amount)"},
            {"name": "最小值", "formula": "=MIN(amount)"},
            {"name": "平均值", "formula": "=AVERAGE(amount)"},
            {"name": "积分值", "formula": "=INTEGRAL(amount, month)"},
        ], ensure_ascii=False)
        conn.execute(
            "UPDATE reports SET formulas=? WHERE name=? AND (formulas IS NULL OR formulas='[]')",
            (demo_formulas, "销售趋势示例"),
        )
        if conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0] == 0:
            now = utc_now()
            conn.execute(
                """
                INSERT INTO reports
                (name, description, sql, chart_type, x_column, y_column, formulas, layout,
                 output_formats, schedule_minutes, enabled, created_at, updated_at, next_run_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    "销售趋势示例",
                    "按月份查看演示销售额，修改 SQL 即可连接自己的 SQLite 数据库。",
                    "SELECT month, SUM(amount) AS amount FROM demo_sales GROUP BY month ORDER BY month",
                    "line",
                    "month",
                    "amount",
                    json.dumps([
                        {"name": "最大值", "formula": "=MAX(amount)"},
                        {"name": "最小值", "formula": "=MIN(amount)"},
                        {"name": "平均值", "formula": "=AVERAGE(amount)"},
                        {"name": "积分值", "formula": "=INTEGRAL(amount, month)"},
                    ], ensure_ascii=False),
                    json.dumps(normalize_layout(None), ensure_ascii=False),
                    json.dumps(["html", "xlsx"], ensure_ascii=False),
                    0,
                    1,
                    iso(now),
                    iso(now),
                    None,
                ),
            )


def row_to_report(row: sqlite3.Row) -> dict[str, Any]:
    result = dict(row)
    result["formulas"] = json.loads(result.get("formulas") or "[]")
    result["layout"] = normalize_layout(result.get("layout") or "{}")
    result["output_formats"] = json.loads(result.get("output_formats") or "[]")
    result["enabled"] = bool(result["enabled"])
    return result


def row_to_run(row: sqlite3.Row) -> dict[str, Any]:
    result = dict(row)
    result["report_name"] = result.pop("report_name", None)
    return result


def get_report(report_id: int) -> dict[str, Any] | None:
    with connection() as conn:
        row = conn.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
    return row_to_report(row) if row else None


def normalize_formulas(value: Any) -> list[dict[str, str]]:
    """Normalize editor lines into named, Excel-like summary formulas."""

    if value is None:
        return []
    if isinstance(value, str):
        items: list[Any] = [line.strip() for line in value.splitlines() if line.strip()]
    elif isinstance(value, list):
        items = value
    else:
        raise ValueError("公式定义必须是文本或数组")
    normalized: list[dict[str, str]] = []
    for item in items:
        if isinstance(item, dict):
            name = str(item.get("name", "")).strip()
            formula = str(item.get("formula", "")).strip()
        else:
            line = str(item).strip()
            if "=" not in line:
                raise ValueError(f"公式缺少 '='：{line}")
            name, formula = line.split("=", 1)
            name, formula = name.strip(), "=" + formula.strip()
        if not name or not formula.startswith("="):
            raise ValueError("每条公式必须写成“名称 = 表达式”")
        if len(name) > 80 or len(formula) > 300:
            raise ValueError("公式名称最多 80 个字符，表达式最多 300 个字符")
        normalized.append({"name": name, "formula": formula})
    if len(normalized) > 30:
        raise ValueError("单份报表最多定义 30 个统计公式")
    return normalized


def normalize_layout(value: Any) -> dict[str, Any]:
    """Keep the browser layout model small and safe for HTML/XLSX rendering."""

    if not value:
        return {"rows": 8, "cols": 8, "cells": {}, "merges": []}
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError as exc:
            raise ValueError("版式定义不是有效 JSON") from exc
    if not isinstance(value, dict):
        raise ValueError("版式定义必须是对象")
    rows = min(max(int(value.get("rows", 8)), 1), 30)
    cols = min(max(int(value.get("cols", 8)), 1), 20)
    cells: dict[str, Any] = {}
    allowed_elements = {"text", "value", "table", "chart"}
    allowed_sources = {"static", "report_field", "sql", "formula"}
    for key, raw_cell in (value.get("cells") or {}).items():
        if not re.fullmatch(r"r\d+c\d+", str(key)) or not isinstance(raw_cell, dict):
            continue
        match = re.fullmatch(r"r(\d+)c(\d+)", str(key))
        if not match or int(match.group(1)) >= rows or int(match.group(2)) >= cols:
            continue
        raw_style = raw_cell.get("style") if isinstance(raw_cell.get("style"), dict) else {}
        color = str(raw_style.get("color", "#172033"))
        background = str(raw_style.get("background", "#ffffff"))
        if not re.fullmatch(r"#[0-9a-fA-F]{3,8}", color):
            color = "#172033"
        if not re.fullmatch(r"#[0-9a-fA-F]{3,8}", background):
            background = "#ffffff"
        style = {
            "bold": bool(raw_style.get("bold", False)),
            "italic": bool(raw_style.get("italic", False)),
            "fontSize": min(max(int(raw_style.get("fontSize", 13)), 8), 48),
            "color": color,
            "background": background,
            "align": str(raw_style.get("align", "left")) if str(raw_style.get("align", "left")) in {"left", "center", "right"} else "left",
            "border": str(raw_style.get("border", "thin")) if str(raw_style.get("border", "thin")) in {"none", "thin", "medium", "dashed"} else "thin",
        }
        element = str(raw_cell.get("element", "text"))
        if element not in allowed_elements:
            element = "text"
        raw_source = raw_cell.get("data_source") if isinstance(raw_cell.get("data_source"), dict) else {}
        source_type = str(raw_source.get("type", "static"))
        if source_type not in allowed_sources:
            source_type = "static"
        source_sql = str(raw_source.get("sql", "")).strip()
        if source_sql:
            source_sql_normalized = source_sql.lower().lstrip()
            if not (source_sql_normalized.startswith("select") or source_sql_normalized.startswith("with") or source_sql_normalized.startswith("explain")) or ";" in source_sql.rstrip(";"):
                raise ValueError(f"单元格 {key} 的数据源 SQL 只允许单条只读查询")
        chart_type = str(raw_source.get("chart_type", "line"))
        if chart_type not in {"line", "bar", "pie"}:
            chart_type = "line"
        aggregate = str(raw_source.get("aggregate", "first"))
        if aggregate not in {"first", "sum", "avg", "min", "max", "count"}:
            aggregate = "first"
        data_source = {
            "type": source_type,
            "field": str(raw_source.get("field", ""))[:120],
            "sql": source_sql[:4000],
            "formula": str(raw_source.get("formula", ""))[:300],
            "aggregate": aggregate,
            "chart_type": chart_type,
            "x_field": str(raw_source.get("x_field", ""))[:120],
            "y_field": str(raw_source.get("y_field", ""))[:120],
        }
        cells[str(key)] = {"value": str(raw_cell.get("value", ""))[:500], "element": element, "data_source": data_source, "style": style}
    merges = []
    for raw_merge in value.get("merges") or []:
        if not isinstance(raw_merge, dict):
            continue
        try:
            row, col = int(raw_merge["row"]), int(raw_merge["col"])
            rowspan, colspan = int(raw_merge.get("rowspan", 1)), int(raw_merge.get("colspan", 1))
        except (KeyError, TypeError, ValueError):
            continue
        if 0 <= row < rows and 0 <= col < cols and 1 <= rowspan <= rows - row and 1 <= colspan <= cols - col and (rowspan > 1 or colspan > 1):
            merges.append({"row": row, "col": col, "rowspan": rowspan, "colspan": colspan})
    return {"rows": rows, "cols": cols, "cells": cells, "merges": merges}


def validate_report_payload(payload: dict[str, Any], existing_id: int | None = None) -> dict[str, Any]:
    name = str(payload.get("name", "")).strip()
    sql = str(payload.get("sql", "")).strip()
    if not name:
        raise ValueError("报表名称不能为空")
    if len(name) > 120:
        raise ValueError("报表名称不能超过 120 个字符")
    if not sql:
        raise ValueError("SQL 不能为空")
    normalized = sql.lower().lstrip()
    if not (normalized.startswith("select") or normalized.startswith("with") or normalized.startswith("explain")):
        raise ValueError("只允许 SELECT、WITH 或 EXPLAIN 查询")
    if ";" in sql.rstrip(";" ):
        raise ValueError("不允许执行多条 SQL")
    chart_type = str(payload.get("chart_type", "table"))
    if chart_type not in ALLOWED_CHARTS:
        raise ValueError(f"不支持的图表类型: {chart_type}")
    try:
        schedule_minutes = max(0, int(payload.get("schedule_minutes", 0) or 0))
    except (TypeError, ValueError) as exc:
        raise ValueError("调度间隔必须是非负整数分钟") from exc
    if schedule_minutes and schedule_minutes < 1:
        raise ValueError("调度间隔至少为 1 分钟")
    formats = payload.get("output_formats", ["html", "xlsx"])
    if isinstance(formats, str):
        formats = [item.strip() for item in formats.split(",") if item.strip()]
    formats = list(dict.fromkeys(str(item) for item in formats))
    if not formats or any(item not in ALLOWED_FORMATS for item in formats):
        raise ValueError("输出格式只能是 html、xlsx 或 csv")
    formulas = normalize_formulas(payload.get("formulas", []))
    layout = normalize_layout(payload.get("layout"))
    return {
        "name": name,
        "description": str(payload.get("description", "")).strip(),
        "sql": sql,
        "chart_type": chart_type,
        "x_column": str(payload.get("x_column", "")).strip(),
        "y_column": str(payload.get("y_column", "")).strip(),
        "formulas": formulas,
        "layout": layout,
        "output_formats": formats,
        "schedule_minutes": schedule_minutes,
        "enabled": bool(payload.get("enabled", True)),
    }


def report_rows() -> list[dict[str, Any]]:
    with connection() as conn:
        rows = conn.execute(
            "SELECT * FROM reports ORDER BY updated_at DESC, id DESC"
        ).fetchall()
    return [row_to_report(row) for row in rows]


def run_rows(limit: int = 100) -> list[dict[str, Any]]:
    with connection() as conn:
        rows = conn.execute(
            """
            SELECT runs.*, reports.name AS report_name
            FROM runs JOIN reports ON reports.id = runs.report_id
            ORDER BY runs.started_at DESC LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [row_to_run(row) for row in rows]


def safe_column(row: dict[str, Any], name: str, fallback_index: int) -> Any:
    if name and name in row:
        return row[name]
    values = list(row.values())
    return values[fallback_index] if len(values) > fallback_index else None


def number(value: Any) -> float | None:
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def excel_column_number(column: str) -> int:
    result = 0
    for char in column.upper():
        if not "A" <= char <= "Z":
            raise ValueError(f"无效单元格列：{column}")
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def excel_values(token: str, rows: list[dict[str, Any]], columns: list[str]) -> list[Any]:
    token = token.strip().strip("`\"")
    if not rows:
        return []
    range_match = re.fullmatch(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", token)
    if range_match:
        start_col, start_row, end_col, end_row = range_match.groups()
        start_col_num, end_col_num = excel_column_number(start_col), excel_column_number(end_col)
        if start_col_num != end_col_num:
            raise ValueError("统计公式目前只支持同一列的单元格范围")
        col_index = start_col_num - 1
        if col_index >= len(columns):
            raise ValueError(f"不存在单元格列：{start_col}")
        return [rows[index - 1].get(columns[col_index]) for index in range(int(start_row), int(end_row) + 1) if 0 < index <= len(rows)]
    cell_match = re.fullmatch(r"([A-Za-z]+)(\d+)", token)
    if cell_match:
        column, row_number = cell_match.groups()
        col_index = excel_column_number(column) - 1
        row_index = int(row_number) - 1
        if col_index >= len(columns) or row_index >= len(rows):
            raise ValueError(f"不存在单元格：{token}")
        return [rows[row_index].get(columns[col_index])]
    if rows and token in rows[0]:
        return [row.get(token) for row in rows]
    raise ValueError(f"找不到数据列或单元格：{token}")


def axis_number(value: Any, fallback: float) -> float:
    result = number(value)
    if result is not None:
        return result
    text = str(value or "").strip()
    for pattern in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(text, pattern)
            return parsed.replace(tzinfo=UTC).timestamp()
        except ValueError:
            continue
    return fallback


def aggregate_formula(function: str, args: list[str], rows: list[dict[str, Any]], columns: list[str]) -> float:
    if function.upper() == "INTEGRAL":
        if len(args) != 2:
            raise ValueError("INTEGRAL 需要两个参数：数值列和时间列")
        y_values = [number(item) for item in excel_values(args[0], rows, columns)]
        x_values = excel_values(args[1], rows, columns)
        pairs = [(axis_number(x, index), y) for index, (x, y) in enumerate(zip(x_values, y_values)) if y is not None]
        if len(pairs) < 2:
            return 0.0
        return sum((right_x - left_x) * (left_y + right_y) / 2 for (left_x, left_y), (right_x, right_y) in zip(pairs, pairs[1:]))
    if len(args) != 1:
        raise ValueError(f"{function.upper()} 需要一个参数")
    values = [number(item) for item in excel_values(args[0], rows, columns)]
    values = [item for item in values if item is not None]
    if function.upper() == "MAX":
        return max(values) if values else 0.0
    if function.upper() == "MIN":
        return min(values) if values else 0.0
    if function.upper() in {"AVERAGE", "AVG"}:
        return sum(values) / len(values) if values else 0.0
    if function.upper() == "SUM":
        return sum(values)
    if function.upper() == "COUNT":
        return float(len(values))
    raise ValueError(f"不支持的统计函数：{function}")


def evaluate_formula(formula: str, rows: list[dict[str, Any]]) -> float | str:
    """Evaluate a safe subset of Excel formulas without executing Python code."""

    columns = list(rows[0].keys()) if rows else []
    expression = formula.strip()[1:].strip() if formula.strip().startswith("=") else formula.strip()
    function_pattern = re.compile(r"\b(MAX|MIN|AVERAGE|AVG|SUM|COUNT|INTEGRAL)\s*\(([^()]*)\)", re.IGNORECASE)
    while True:
        match = function_pattern.search(expression)
        if not match:
            break
        args = [item.strip() for item in match.group(2).split(",")]
        value = aggregate_formula(match.group(1), args, rows, columns)
        expression = expression[:match.start()] + repr(value) + expression[match.end():]
    cell_pattern = re.compile(r"\b([A-Za-z]+\d+(?::[A-Za-z]+\d+)?)\b")
    expression = cell_pattern.sub(lambda match: repr(excel_values(match.group(1), rows, columns)[0]), expression)
    import ast
    tree = ast.parse(expression, mode="eval")
    allowed = (ast.Expression, ast.Constant, ast.UnaryOp, ast.UAdd, ast.USub, ast.BinOp, ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow, ast.Mod)
    if any(not isinstance(node, allowed) for node in ast.walk(tree)):
        raise ValueError("公式包含不支持的运算")
    value = eval(compile(tree, "<report-formula>", "eval"), {"__builtins__": {}}, {})  # noqa: S307
    if not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ValueError("公式结果不是有限数字")
    return float(value)


def evaluate_formulas(report: dict[str, Any], rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    results = []
    for item in report.get("formulas", []):
        try:
            value: Any = evaluate_formula(item["formula"], rows)
        except Exception as exc:  # noqa: BLE001 - show formula errors in the report
            value = f"#ERROR: {exc}"
        results.append({"name": item["name"], "formula": item["formula"], "value": value})
    return results


def layout_style(style: dict[str, Any]) -> str:
    border = style.get("border", "thin")
    border_value = "0" if border == "none" else f"{border} #cbd5e1" if border == "dashed" else f"{border} solid #cbd5e1"
    return ";".join([
        f"font-size:{int(style.get('fontSize', 13))}px",
        f"font-weight:{'700' if style.get('bold') else '400'}",
        f"font-style:{'italic' if style.get('italic') else 'normal'}",
        f"color:{html.escape(str(style.get('color', '#172033')))}",
        f"background:{html.escape(str(style.get('background', '#ffffff')))}",
        f"text-align:{style.get('align', 'left')}",
        f"border:{border_value}",
        "padding:8px;min-width:80px;height:34px;white-space:pre-wrap",
    ])


def aggregate_cell_values(values: list[Any], aggregate: str) -> Any:
    if not values:
        return ""
    if aggregate == "first":
        return values[0]
    numbers = [number(item) for item in values]
    numbers = [item for item in numbers if item is not None]
    if aggregate == "count":
        return len(values)
    if not numbers:
        return ""
    if aggregate == "sum":
        return sum(numbers)
    if aggregate == "avg":
        return sum(numbers) / len(numbers)
    if aggregate == "min":
        return min(numbers)
    if aggregate == "max":
        return max(numbers)
    return values[0]


def execute_data_source(source: dict[str, Any], report_rows: list[dict[str, Any]]) -> Any:
    source_type = source.get("type", "static")
    if source_type == "static":
        return source.get("value", "")
    if source_type == "formula":
        return evaluate_formula(str(source.get("formula", "")), report_rows)
    if source_type == "report_field":
        field = str(source.get("field", ""))
        values = [row.get(field) for row in report_rows if field in row]
        return aggregate_cell_values(values, str(source.get("aggregate", "first")))
    if source_type == "sql":
        sql = str(source.get("sql", "")).strip()
        if not sql:
            return []
        with connection() as conn:
            conn.execute("PRAGMA query_only = ON")
            result = conn.execute(sql)
            columns = [item[0] for item in result.description or []]
            return [{column: json_value(value) for column, value in zip(columns, row)} for row in result.fetchall()]
    return ""


def cell_data_rows(source: dict[str, Any], report_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    value = execute_data_source(source, report_rows)
    if isinstance(value, list):
        return [item for item in value if isinstance(item, dict)]
    return []


def render_data_table(data_rows: list[dict[str, Any]]) -> str:
    if not data_rows:
        return '<span class="muted">暂无数据</span>'
    columns = list(data_rows[0].keys())
    head = "".join(f"<th>{html.escape(str(column))}</th>" for column in columns)
    body = "".join("<tr>" + "".join(f"<td>{html.escape(str(row.get(column, '')))}</td>" for column in columns) + "</tr>" for row in data_rows)
    return f'<table class="element-table"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def render_layout_cell(cell: dict[str, Any], report_rows: list[dict[str, Any]]) -> str:
    source = cell.get("data_source") or {"type": "static"}
    source.setdefault("value", cell.get("value", ""))
    value = execute_data_source(source, report_rows)
    element = cell.get("element", "text")
    if element == "table":
        return render_data_table(cell_data_rows(source, report_rows))
    if element == "chart":
        data_rows = cell_data_rows(source, report_rows)
        if not data_rows and isinstance(value, list):
            data_rows = value
        columns = list(data_rows[0].keys()) if data_rows else []
        chart_type = source.get("chart_type", "line")
        x_field = source.get("x_field") or (columns[0] if columns else "")
        y_field = source.get("y_field") or (columns[1] if len(columns) > 1 else "")
        return chart_svg(data_rows, chart_type, x_field, y_field) if data_rows else '<span class="muted">暂无图表数据</span>'
    if isinstance(value, list):
        if value and isinstance(value[0], dict):
            value = value[0].get(source.get("field", ""), next(iter(value[0].values()), ""))
        else:
            value = value[0] if value else ""
    return html.escape(str(value if value is not None else ""))


def render_layout_html(layout: dict[str, Any], report_rows: list[dict[str, Any]]) -> str:
    if not layout.get("cells") and not layout.get("merges"):
        return ""
    cells = layout.get("cells", {})
    merges = layout.get("merges", [])
    covered: set[tuple[int, int]] = set()
    merge_map: dict[tuple[int, int], dict[str, int]] = {}
    for merge in merges:
        row, col = merge["row"], merge["col"]
        merge_map[(row, col)] = merge
        covered.update((row + r, col + c) for r in range(merge["rowspan"]) for c in range(merge["colspan"]) if (r, c) != (0, 0))
    rows_html = []
    for row in range(layout.get("rows", 8)):
        cells_html = []
        for col in range(layout.get("cols", 8)):
            if (row, col) in covered:
                continue
            key = f"r{row}c{col}"
            cell = cells.get(key, {})
            style = cell.get("style", {})
            merge = merge_map.get((row, col), {})
            attrs = f' rowspan="{merge.get("rowspan", 1)}" colspan="{merge.get("colspan", 1)}"' if merge else ""
            cells_html.append(f'<td{attrs} style="{layout_style(style)}">{render_layout_cell(cell, report_rows)}</td>')
        rows_html.append(f"<tr>{''.join(cells_html)}</tr>")
    return f'<section class="designed-layout"><h2>报表版式</h2><table><tbody>{"".join(rows_html)}</tbody></table></section>'


def chart_svg(rows: list[dict[str, Any]], chart_type: str, x_column: str, y_column: str) -> str:
    if chart_type == "table" or not rows:
        return ""
    points = []
    for row in rows[:40]:
        x = safe_column(row, x_column, 0)
        y = number(safe_column(row, y_column, 1))
        if y is not None:
            points.append((str(x), y))
    if not points:
        return '<p class="muted">没有可绘制的数值列</p>'
    width, height = 760, 300
    left, top, right, bottom = 56, 20, 20, 46
    plot_w, plot_h = width - left - right, height - top - bottom
    max_y = max(v for _, v in points) or 1
    min_y = min(0, min(v for _, v in points))
    span = max_y - min_y or 1
    colors = ["#2b6cb0", "#d97706", "#0f766e", "#be123c", "#7c3aed", "#047857"]
    svg: list[str] = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="报表图表">',
        f'<line x1="{left}" y1="{top + plot_h}" x2="{width - right}" y2="{top + plot_h}" class="axis"/>',
        f'<line x1="{left}" y1="{top}" x2="{left}" y2="{top + plot_h}" class="axis"/>',
    ]
    if chart_type == "pie":
        total = sum(abs(v) for _, v in points) or 1
        radius, cx, cy = 88, 190, 140
        circumference = 2 * math.pi * radius
        offset = 0.0
        svg.append(f'<circle cx="{cx}" cy="{cy}" r="{radius}" fill="none" stroke="#e5e7eb" stroke-width="58"/>')
        for index, (label, value) in enumerate(points):
            length = abs(value) / total * circumference
            color = colors[index % len(colors)]
            svg.append(
                f'<circle cx="{cx}" cy="{cy}" r="{radius}" fill="none" stroke="{color}" stroke-width="58" '
                f'stroke-dasharray="{length:.2f} {circumference:.2f}" stroke-dashoffset="{-offset:.2f}" '
                'transform="rotate(-90 190 140)"/>'
            )
            offset += length
            ly = 36 + index * 24
            svg.append(f'<rect x="340" y="{ly - 11}" width="12" height="12" fill="{color}"/>')
            svg.append(f'<text x="360" y="{ly}" class="legend">{html.escape(label[:24])} {value:,.2f}</text>')
    else:
        step = plot_w / max(len(points), 1)
        coords: list[tuple[float, float]] = []
        for index, (label, value) in enumerate(points):
            x = left + step * (index + 0.5)
            y = top + (max_y - value) / span * plot_h
            coords.append((x, y))
            if chart_type == "bar":
                zero_y = top + (max_y - 0) / span * plot_h
                bar_y, bar_h = min(y, zero_y), abs(zero_y - y)
                svg.append(f'<rect x="{x - step * .3:.2f}" y="{bar_y:.2f}" width="{step * .6:.2f}" height="{max(bar_h, 1):.2f}" class="bar"/>')
            svg.append(f'<text x="{x:.2f}" y="{height - 15}" class="label" text-anchor="middle">{html.escape(label[:14])}</text>')
        if chart_type == "line":
            polyline = " ".join(f"{x:.2f},{y:.2f}" for x, y in coords)
            svg.append(f'<polyline points="{polyline}" class="line"/>')
            for x, y in coords:
                svg.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="4" class="point"/>')
    svg.append("</svg>")
    return "".join(svg)


def render_html(report: dict[str, Any], rows: list[dict[str, Any]], run_id: int, summaries: list[dict[str, Any]]) -> str:
    columns = list(rows[0].keys()) if rows else []
    head = "".join(f"<th>{html.escape(str(column))}</th>" for column in columns)
    body = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(row.get(column, '')))}</td>" for column in columns) + "</tr>"
        for row in rows
    )
    summary_rows = "".join(
        f"<tr><th>{html.escape(str(item['name']))}</th><td><code>{html.escape(str(item['formula']))}</code></td><td>{html.escape(str(item['value']))}</td></tr>"
        for item in summaries
    )
    summary = f'<section class="summary"><h2>统计值</h2><table><thead><tr><th>指标</th><th>公式</th><th>结果</th></tr></thead><tbody>{summary_rows}</tbody></table></section>' if summaries else ""
    chart = chart_svg(rows, report["chart_type"], report["x_column"], report["y_column"])
    designed_layout = render_layout_html(report.get("layout", {}), rows)
    return f"""<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><title>{html.escape(report['name'])}</title>
<style>body{{font-family:system-ui, sans-serif;margin:32px;color:#182230}} table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #d8dee8;padding:8px;text-align:left}} th{{background:#eef3f8}} .chart{{max-width:760px;width:100%;margin:12px 0}} .axis{{stroke:#64748b;stroke-width:1}} .bar{{fill:#2b6cb0;opacity:.85}} .line{{fill:none;stroke:#d97706;stroke-width:3}} .point{{fill:#d97706}} .label,.legend{{font-size:12px;fill:#334155}} .muted{{color:#64748b}}</style></head>
<body><h1>{html.escape(report['name'])}</h1><p>{html.escape(report['description'])}</p>
<p>运行编号：{run_id}　数据行数：{len(rows)}</p>{designed_layout}{summary}{chart}<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>
</body></html>"""


def write_outputs(report: dict[str, Any], rows: list[dict[str, Any]], run_id: int, summaries: list[dict[str, Any]]) -> Path:
    output_dir = OUTPUT_DIR / f"run-{run_id}"
    output_dir.mkdir(parents=True, exist_ok=True)
    with (output_dir / "data.json").open("w", encoding="utf-8") as fh:
        json.dump(rows, fh, ensure_ascii=False, indent=2, default=str)
    with (output_dir / "summary.json").open("w", encoding="utf-8") as fh:
        json.dump(summaries, fh, ensure_ascii=False, indent=2, default=str)
    formats = set(report["output_formats"])
    if "html" in formats:
        (output_dir / "report.html").write_text(render_html(report, rows, run_id, summaries), encoding="utf-8")
    if "csv" in formats:
        columns = list(rows[0].keys()) if rows else []
        with (output_dir / "report.csv").open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=columns)
            if columns:
                writer.writeheader()
                writer.writerows(rows)
    if "xlsx" in formats:
        try:
            from openpyxl import Workbook
            from openpyxl.chart import BarChart, LineChart, PieChart, Reference
            from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
            from openpyxl.utils import get_column_letter
        except ImportError as exc:
            raise RuntimeError("生成 XLSX 需要安装 openpyxl") from exc
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        columns = list(rows[0].keys()) if rows else []
        sheet.append(columns)
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2B6CB0")
        for row in rows:
            sheet.append([row.get(column) for column in columns])
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column_cells in sheet.columns:
            width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 36)
            sheet.column_dimensions[column_cells[0].column_letter].width = width
        if report["chart_type"] in {"bar", "line", "pie"} and len(columns) >= 2 and rows:
            chart_type = {"bar": BarChart, "line": LineChart, "pie": PieChart}[report["chart_type"]]
            chart = chart_type()
            data = Reference(sheet, min_col=2, min_row=1, max_row=len(rows) + 1)
            labels = Reference(sheet, min_col=1, min_row=2, max_row=len(rows) + 1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(labels)
            chart.title = report["name"]
            sheet.add_chart(chart, "H2")
        if summaries:
            summary_sheet = workbook.create_sheet("统计值")
            summary_sheet.append(["指标", "公式", "结果"])
            for cell in summary_sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="0F766E")
            for item in summaries:
                summary_sheet.append([item["name"], item["formula"], item["value"]])
            summary_sheet.freeze_panes = "A2"
            for column_cells in summary_sheet.columns:
                width = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 50)
                summary_sheet.column_dimensions[column_cells[0].column_letter].width = width
        layout = report.get("layout") or {}
        if layout.get("cells") or layout.get("merges"):
            layout_sheet = workbook.create_sheet("版式")
            chart_data_sheet = None
            for row in range(layout.get("rows", 8)):
                for col in range(layout.get("cols", 8)):
                    cell = layout_sheet.cell(row=row + 1, column=col + 1)
                    definition = (layout.get("cells") or {}).get(f"r{row}c{col}", {})
                    style = definition.get("style", {})
                    source = dict(definition.get("data_source") or {"type": "static"})
                    source.setdefault("value", definition.get("value", ""))
                    element = definition.get("element", "text")
                    if element == "table":
                        table_rows = cell_data_rows(source, rows)
                        cell.value = ""
                        if table_rows:
                            table_columns = list(table_rows[0].keys())
                            for offset, column_name in enumerate(table_columns):
                                target = layout_sheet.cell(row=row + 1, column=col + offset + 1, value=column_name)
                                target.font = Font(bold=True, color="FFFFFF")
                                target.fill = PatternFill("solid", fgColor="2563EB")
                            for row_offset, data_row in enumerate(table_rows, start=1):
                                for offset, column_name in enumerate(table_columns):
                                    layout_sheet.cell(row=row + row_offset + 1, column=col + offset + 1, value=data_row.get(column_name))
                    elif element == "chart":
                        cell.value = definition.get("value", "")
                        chart_rows = cell_data_rows(source, rows)
                        chart_columns = list(chart_rows[0].keys()) if chart_rows else []
                        if chart_rows and len(chart_columns) >= 2:
                            if chart_data_sheet is None:
                                chart_data_sheet = workbook.create_sheet("图元数据")
                            base_col = 1
                            base_row = chart_data_sheet.max_row + 2
                            chart_data_sheet.cell(row=base_row, column=base_col, value=f"R{row + 1}C{col + 1}")
                            for offset, column_name in enumerate(chart_columns):
                                chart_data_sheet.cell(row=base_row + 1, column=base_col + offset, value=column_name)
                            for row_offset, data_row in enumerate(chart_rows, start=2):
                                for offset, column_name in enumerate(chart_columns):
                                    chart_data_sheet.cell(row=base_row + row_offset, column=base_col + offset, value=data_row.get(column_name))
                            chart_kind = {"bar": BarChart, "line": LineChart, "pie": PieChart}.get(source.get("chart_type", "line"), LineChart)
                            element_chart = chart_kind()
                            data = Reference(chart_data_sheet, min_col=base_col + 1, min_row=base_row + 1, max_row=base_row + len(chart_rows) + 1)
                            labels = Reference(chart_data_sheet, min_col=base_col, min_row=base_row + 2, max_row=base_row + len(chart_rows) + 1)
                            element_chart.add_data(data, titles_from_data=True)
                            element_chart.set_categories(labels)
                            element_chart.title = str(definition.get("value", "图表"))
                            layout_sheet.add_chart(element_chart, cell.coordinate)
                    else:
                        resolved = execute_data_source(source, rows)
                        if isinstance(resolved, list):
                            if resolved and isinstance(resolved[0], dict):
                                resolved = resolved[0].get(source.get("field", ""), next(iter(resolved[0].values()), ""))
                            else:
                                resolved = resolved[0] if resolved else ""
                        cell.value = resolved
                    cell.font = Font(name="Microsoft YaHei", size=style.get("fontSize", 13), bold=style.get("bold", False), italic=style.get("italic", False), color=str(style.get("color", "#172033")).replace("#", ""))
                    cell.alignment = Alignment(horizontal=style.get("align", "left"), vertical="center", wrap_text=True)
                    cell.fill = PatternFill("solid", fgColor=str(style.get("background", "#ffffff")).replace("#", ""))
                    border_style = style.get("border", "thin")
                    side = Side(style=None if border_style == "none" else "dashed" if border_style == "dashed" else border_style, color="CBD5E1")
                    cell.border = Border(left=side, right=side, top=side, bottom=side)
                layout_sheet.row_dimensions[row + 1].height = 28
            for merge in layout.get("merges", []):
                start = layout_sheet.cell(row=merge["row"] + 1, column=merge["col"] + 1).coordinate
                end = layout_sheet.cell(row=merge["row"] + merge["rowspan"], column=merge["col"] + merge["colspan"]).coordinate
                layout_sheet.merge_cells(f"{start}:{end}")
            for col in range(layout.get("cols", 8)):
                layout_sheet.column_dimensions[get_column_letter(col + 1)].width = 16
        workbook.save(output_dir / "report.xlsx")
    return output_dir


def execute_report(report_id: int) -> dict[str, Any]:
    report = get_report(report_id)
    if not report:
        raise ValueError("报表不存在")
    started = utc_now()
    with connection() as conn:
        cursor = conn.execute(
            "INSERT INTO runs(report_id, status, started_at) VALUES (?, ?, ?)",
            (report_id, "running", iso(started)),
        )
        run_id = cursor.lastrowid
    try:
        with connection() as conn:
            # Report definitions are read-only by contract, including WITH
            # statements that could otherwise wrap an UPDATE/DELETE in SQLite.
            conn.execute("PRAGMA query_only = ON")
            result = conn.execute(report["sql"])
            columns = [item[0] for item in result.description or []]
            rows = [{column: json_value(value) for column, value in zip(columns, row)} for row in result.fetchall()]
        summaries = evaluate_formulas(report, rows)
        output_dir = write_outputs(report, rows, run_id, summaries)
        finished = utc_now()
        next_run = finished + timedelta(minutes=report["schedule_minutes"]) if report["schedule_minutes"] else None
        with connection() as conn:
            conn.execute(
                "UPDATE runs SET status=?, finished_at=?, row_count=?, output_dir=? WHERE id=?",
                ("success", iso(finished), len(rows), str(output_dir), run_id),
            )
            conn.execute(
                "UPDATE reports SET last_run_at=?, next_run_at=?, updated_at=? WHERE id=?",
                (iso(finished), iso(next_run), iso(finished), report_id),
            )
        return get_run(run_id) or {"id": run_id, "status": "success"}
    except Exception as exc:  # noqa: BLE001 - persisted for the UI
        finished = utc_now()
        with connection() as conn:
            conn.execute(
                "UPDATE runs SET status=?, finished_at=?, error=? WHERE id=?",
                ("failed", iso(finished), str(exc), run_id),
            )
        raise


def get_run(run_id: int) -> dict[str, Any] | None:
    with connection() as conn:
        row = conn.execute(
            "SELECT runs.*, reports.name AS report_name FROM runs JOIN reports ON reports.id=runs.report_id WHERE runs.id=?",
            (run_id,),
        ).fetchone()
    return row_to_run(row) if row else None


class Scheduler:
    def __init__(self) -> None:
        self.running = False
        self.paused = False
        self.thread: threading.Thread | None = None
        self.stop_event = threading.Event()
        self.lock = threading.Lock()
        self.active: dict[int, threading.Thread] = {}

    def start(self) -> None:
        with self.lock:
            if self.running:
                return
            self.running = True
            self.paused = False
            self.stop_event.clear()
            self.thread = threading.Thread(target=self._loop, name="report-scheduler", daemon=True)
            self.thread.start()

    def stop(self) -> None:
        with self.lock:
            self.running = False
            self.stop_event.set()

    def pause(self) -> None:
        self.paused = True

    def resume(self) -> None:
        if not self.running:
            self.start()
        self.paused = False

    def trigger(self, report_id: int) -> bool:
        with self.lock:
            if report_id in self.active:
                return False
            thread = threading.Thread(target=self._run_one, args=(report_id,), daemon=True)
            self.active[report_id] = thread
            thread.start()
            return True

    def _run_one(self, report_id: int) -> None:
        try:
            execute_report(report_id)
        except Exception:
            pass
        finally:
            with self.lock:
                self.active.pop(report_id, None)

    def _loop(self) -> None:
        while not self.stop_event.wait(1):
            if not self.running or self.paused:
                continue
            now = utc_now()
            for report in report_rows():
                if not report["enabled"] or not report["schedule_minutes"]:
                    continue
                next_run = parse_iso(report.get("next_run_at"))
                if next_run and next_run <= now:
                    self.trigger(report["id"])

    def status(self) -> dict[str, Any]:
        return {
            "running": self.running,
            "paused": self.paused,
            "active_runs": list(self.active),
            "database": str(DB_PATH),
        }


SCHEDULER = Scheduler()


def parse_id(value: str) -> int:
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError("无效的编号") from exc


class PortalHandler(BaseHTTPRequestHandler):
    server_version = "ReportPortal/0.1"

    def log_message(self, format: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {format % args}")

    def send_json(self, payload: Any, status: int = HTTPStatus.OK) -> None:
        raw = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def send_file(self, path: Path, content_type: str, download_name: str | None = None) -> None:
        if not path.exists() or not path.is_file():
            self.send_json({"error": "文件不存在"}, HTTPStatus.NOT_FOUND)
            return
        raw = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        if download_name:
            self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length > 2_000_000:
            raise ValueError("请求体过大")
        return json.loads(self.rfile.read(length) or b"{}")

    def route(self) -> tuple[str, dict[str, list[str]]]:
        parsed = urlparse(self.path)
        return parsed.path.rstrip("/") or "/", parse_qs(parsed.query)

    def do_GET(self) -> None:  # noqa: N802
        path, query = self.route()
        try:
            if path in {"/", "/index.html"}:
                self.send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            elif path.startswith("/static/"):
                rel = unquote(path.removeprefix("/static/")).replace("\\", "/")
                target = (STATIC_DIR / rel).resolve()
                if STATIC_DIR.resolve() not in target.parents:
                    raise ValueError("非法文件路径")
                content_type = "text/plain; charset=utf-8" if target.suffix == ".js" else "text/css; charset=utf-8"
                self.send_file(target, content_type)
            elif path == "/api/status":
                self.send_json(SCHEDULER.status())
            elif path == "/api/reports":
                self.send_json({"items": report_rows()})
            elif path.startswith("/api/reports/"):
                report_id = parse_id(path.split("/")[3])
                report = get_report(report_id)
                if not report:
                    self.send_json({"error": "报表不存在"}, HTTPStatus.NOT_FOUND)
                else:
                    self.send_json(report)
            elif path == "/api/runs":
                self.send_json({"items": run_rows()})
            elif path.startswith("/api/runs/"):
                parts = path.split("/")
                run_id = parse_id(parts[3])
                run = get_run(run_id)
                if not run:
                    self.send_json({"error": "运行记录不存在"}, HTTPStatus.NOT_FOUND)
                    return
                if len(parts) >= 5 and parts[4] == "data":
                    data_path = OUTPUT_DIR / f"run-{run_id}" / "data.json"
                    self.send_file(data_path, "application/json; charset=utf-8")
                elif len(parts) >= 5 and parts[4] == "summary":
                    summary_path = OUTPUT_DIR / f"run-{run_id}" / "summary.json"
                    self.send_file(summary_path, "application/json; charset=utf-8")
                elif len(parts) >= 5 and parts[4] == "file":
                    fmt = query.get("format", ["html"])[0]
                    mapping = {"html": ("report.html", "text/html; charset=utf-8"), "xlsx": ("report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), "csv": ("report.csv", "text/csv; charset=utf-8")}
                    if fmt not in mapping:
                        raise ValueError("不支持的文件格式")
                    filename, content_type = mapping[fmt]
                    self.send_file(OUTPUT_DIR / f"run-{run_id}" / filename, content_type, None if fmt == "html" else filename)
                else:
                    self.send_json(run)
            elif path == "/api/database/tables":
                with connection() as conn:
                    rows = conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name").fetchall()
                self.send_json({"items": [row[0] for row in rows]})
            else:
                self.send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
        except (ValueError, json.JSONDecodeError) as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # noqa: BLE001 - keep API errors visible
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_POST(self) -> None:  # noqa: N802
        path, _ = self.route()
        try:
            payload = self.read_json()
            if path == "/api/process/start":
                SCHEDULER.start()
                self.send_json(SCHEDULER.status())
            elif path == "/api/process/stop":
                SCHEDULER.stop()
                self.send_json(SCHEDULER.status())
            elif path == "/api/process/pause":
                SCHEDULER.pause()
                self.send_json(SCHEDULER.status())
            elif path == "/api/process/resume":
                SCHEDULER.resume()
                self.send_json(SCHEDULER.status())
            elif path == "/api/reports":
                data = validate_report_payload(payload)
                now = utc_now()
                next_run = now + timedelta(minutes=data["schedule_minutes"]) if data["schedule_minutes"] else None
                with connection() as conn:
                    cursor = conn.execute(
                        """
                        INSERT INTO reports
                        (name, description, sql, chart_type, x_column, y_column, formulas, layout, output_formats,
                         schedule_minutes, enabled, created_at, updated_at, next_run_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (data["name"], data["description"], data["sql"], data["chart_type"], data["x_column"], data["y_column"], json.dumps(data["formulas"], ensure_ascii=False), json.dumps(data["layout"], ensure_ascii=False), json.dumps(data["output_formats"], ensure_ascii=False), data["schedule_minutes"], int(data["enabled"]), iso(now), iso(now), iso(next_run)),
                    )
                    report_id = cursor.lastrowid
                self.send_json(get_report(report_id), HTTPStatus.CREATED)
            elif path.endswith("/run") and path.startswith("/api/reports/"):
                report_id = parse_id(path.split("/")[3])
                if not get_report(report_id):
                    self.send_json({"error": "报表不存在"}, HTTPStatus.NOT_FOUND)
                elif not SCHEDULER.trigger(report_id):
                    self.send_json({"error": "该报表已有运行任务"}, HTTPStatus.CONFLICT)
                else:
                    self.send_json({"accepted": True, "report_id": report_id}, HTTPStatus.ACCEPTED)
            else:
                self.send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
        except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_PUT(self) -> None:  # noqa: N802
        path, _ = self.route()
        try:
            if not path.startswith("/api/reports/"):
                self.send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
                return
            report_id = parse_id(path.split("/")[3])
            existing = get_report(report_id)
            if not existing:
                self.send_json({"error": "报表不存在"}, HTTPStatus.NOT_FOUND)
                return
            payload = self.read_json()
            merged = {**existing, **payload}
            data = validate_report_payload(merged, report_id)
            now = utc_now()
            next_run = now + timedelta(minutes=data["schedule_minutes"]) if data["schedule_minutes"] else None
            with connection() as conn:
                conn.execute(
                    """
                    UPDATE reports SET name=?, description=?, sql=?, chart_type=?, x_column=?, y_column=?,
                    formulas=?, layout=?, output_formats=?, schedule_minutes=?, enabled=?, updated_at=?, next_run_at=? WHERE id=?
                    """,
                    (data["name"], data["description"], data["sql"], data["chart_type"], data["x_column"], data["y_column"], json.dumps(data["formulas"], ensure_ascii=False), json.dumps(data["layout"], ensure_ascii=False), json.dumps(data["output_formats"], ensure_ascii=False), data["schedule_minutes"], int(data["enabled"]), iso(now), iso(next_run), report_id),
                )
            self.send_json(get_report(report_id))
        except (ValueError, json.JSONDecodeError, sqlite3.IntegrityError) as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_DELETE(self) -> None:  # noqa: N802
        path, _ = self.route()
        try:
            if not path.startswith("/api/reports/"):
                self.send_json({"error": "接口不存在"}, HTTPStatus.NOT_FOUND)
                return
            report_id = parse_id(path.split("/")[3])
            with connection() as conn:
                cursor = conn.execute("DELETE FROM reports WHERE id=?", (report_id,))
            if cursor.rowcount == 0:
                self.send_json({"error": "报表不存在"}, HTTPStatus.NOT_FOUND)
            else:
                self.send_json({"deleted": True})
        except ValueError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:  # noqa: BLE001
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)


def main() -> None:
    parser = argparse.ArgumentParser(description="Report Portal")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8766)
    args = parser.parse_args()
    init_db()
    SCHEDULER.start()
    server = ThreadingHTTPServer((args.host, args.port), PortalHandler)
    print(f"Report Portal running at http://{args.host}:{args.port}")
    print(f"SQLite database: {DB_PATH}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        SCHEDULER.stop()
        server.server_close()


if __name__ == "__main__":
    main()
