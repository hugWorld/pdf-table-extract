"""把提取结果写成 Excel。"""

from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

from .core import Table, safe_sheet_name

ILLEGAL = re.compile(r"[\x00-\x08\x0b-\x0c\x0e-\x1f\x7f\ufeff\u200b-\u200f]")

HEADER_FILL = PatternFill("solid", fgColor="F2F1EC")
SUSPECT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Font(name="Arial", size=10)
THIN_BOLD = Font(name="Arial", size=10, bold=True)
LINK = Font(name="Arial", size=10, color="1F5FBF", underline="single")

SOURCE_HEADER = ("\u6765\u6e90\u6587\u4ef6", "\u9875\u7801", "\u8868\u53f7")
INDEX_TITLE = "\u76ee\u5f55"
INDEX_HEADER = ("\u5e8f\u53f7", "\u5de5\u4f5c\u8868", "\u6765\u6e90\u6587\u4ef6",
                "\u9875\u7801", "\u8868\u53f7", "\u884c\u00d7\u5217", "\u53ef\u7591")


def _display(text: str) -> str | int | float | None:
    t = ILLEGAL.sub("", text).strip()
    if not t:
        return None
    cleaned = t.replace("\u2212", "-").replace(",", "")
    try:
        if cleaned.lstrip("+-").replace(".", "", 1).isdigit() and len(cleaned) < 16:
            value = float(cleaned)
            return int(value) if value.is_integer() and "." not in cleaned else value
    except ValueError:
        pass
    return t


def _width(values: list[str]) -> float:
    longest = max((len(v) for v in values), default=6)
    return min(max(longest * 1.15 + 3, 8), 46)


def short_source(name: str, room: int) -> str:
    """把文件名压成能塞进工作表名的短标识，尾部（编号）优先保留。"""
    stem = Path(name).stem
    stem = re.sub(r"^(?:journal|article|fulltext|main)[._\-\s]+", "", stem, flags=re.I)
    stem = re.sub(r"[_.]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    if len(stem) <= room:
        return stem
    head = stem[: max(room // 2 - 1, 2)].rstrip()
    tail = stem[-(room - len(head) - 1):].lstrip()
    return f"{head}~{tail}"


def _label(table: Table) -> str:
    return table.label or f"P{table.page}-{table.seq}"


def _sheet_title(table: Table, multi: bool) -> str:
    label = _label(table)
    if not multi:
        return label
    room = max(31 - len(label) - 1, 6)
    short = short_source(table.source_name, room)
    return f"{short} {label}" if short else label


def _write_table(ws, table: Table, offset: int, with_source: bool) -> None:
    for r, row in enumerate(table.data):
        row_idx = r + 1
        for c, cell in enumerate(row):
            target = ws.cell(row=row_idx, column=c + 1 + offset,
                             value=_display(cell.text))
            target.font = THIN_BOLD if r == 0 else THIN
            target.alignment = Alignment(vertical="center")
            if r == 0:
                target.fill = HEADER_FILL
            if cell.suspect:
                target.fill = SUSPECT_FILL
                target.comment = Comment(cell.note or "\u8bf7\u5bf9\u7167\u539f\u6587\u6838\u5bf9",
                                         "extractTab")
        if with_source:
            vals = (SOURCE_HEADER if r == 0
                    else (table.source_name, table.page, table.label))
            for col, value in enumerate(vals, start=1):
                src = ws.cell(row=row_idx, column=col, value=value)
                src.font = THIN_BOLD if r == 0 else THIN
                if r == 0:
                    src.fill = HEADER_FILL

    for c in range(table.ncols):
        col_idx = c + 1 + offset
        column = [str(ws.cell(row=r + 1, column=col_idx).value or "")
                  for r in range(len(table.data))]
        ws.column_dimensions[get_column_letter(col_idx)].width = _width(column)

    if with_source:
        for col, width in enumerate((22, 8, 12), start=1):
            ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = ws.cell(row=2, column=1)
    if ws.max_row > 1:
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(table.ncols + offset)}{ws.max_row}"
        )


def _write_index(wb, rows: list[tuple[str, Table]]) -> None:
    """目录页：一眼看出每个工作表来自哪篇论文、第几页。"""
    ws = wb.create_sheet(INDEX_TITLE, 0)
    for col, name in enumerate(INDEX_HEADER, start=1):
        cell = ws.cell(row=1, column=col, value=name)
        cell.font = THIN_BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")

    for i, (title, table) in enumerate(rows, start=1):
        r = i + 1
        ws.cell(row=r, column=1, value=i).font = THIN
        link = ws.cell(row=r, column=2, value=title)
        link.font = LINK
        link.hyperlink = Hyperlink(ref=link.coordinate,
                                   location=f"'{title}'!A1", display=title)
        ws.cell(row=r, column=3, value=table.source_name).font = THIN
        ws.cell(row=r, column=4, value=table.page).font = THIN
        ws.cell(row=r, column=5, value=table.label).font = THIN
        ws.cell(row=r, column=6, value=f"{table.nrows} \u00d7 {table.ncols}").font = THIN
        flagged = table.suspect_count
        note = ws.cell(row=r, column=7, value=flagged if flagged else "")
        note.font = THIN
        if flagged:
            note.fill = SUSPECT_FILL

    for col, width in enumerate((6, 34, 34, 7, 12, 10, 7), start=1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    if len(rows) > 1:
        ws.auto_filter.ref = f"A1:G{len(rows) + 1}"


def write_workbook(tables: list[Table], out_path: str | Path,
                   with_source: bool = False) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    offset = 3 if with_source else 0
    multi = len({t.source_name for t in tables}) > 1

    index_rows: list[tuple[str, Table]] = []
    for table in tables:
        title = safe_sheet_name(_sheet_title(table, multi), used)
        ws = wb.create_sheet(title)
        _write_table(ws, table, offset, with_source)
        index_rows.append((title, table))

    if not index_rows:
        wb.create_sheet("\u7a7a")
    else:
        _write_index(wb, index_rows)
        wb.active = 0

    out = Path(out_path)
    wb.save(out)
    return out
