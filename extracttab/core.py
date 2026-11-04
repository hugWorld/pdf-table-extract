"""extractTab 核心引擎：从 PDF 中定位表格、重建单元格文本、导出 Excel。"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

SHEET_BAD = re.compile(r"[\\/*?:\[\]]")


@dataclass
class Cell:
    text: str = ""
    suspect: bool = False
    note: str = ""


@dataclass
class Table:
    page: int
    seq: int
    label: str
    nrows: int
    ncols: int
    bbox: tuple
    data: list[list[Cell]] = field(default_factory=list)
    source_name: str = ""

    @property
    def suspect_count(self) -> int:
        return sum(1 for row in self.data for c in row if c.suspect)

    @property
    def numeric_ratio(self) -> float:
        total = filled = numeric = 0
        for row in self.data:
            for c in row:
                total += 1
                t = c.text.strip()
                if not t:
                    continue
                filled += 1
                if is_number(t):
                    numeric += 1
        return numeric / filled if filled else 0.0


def is_number(text: str) -> bool:
    t = text.replace("\u2212", "-").replace(",", "").strip()
    if not t:
        return False
    return bool(re.fullmatch(r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?", t))


class PdfDocument:
    """包装一页 PDF，缓存字符流与碎片字体信息。"""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.doc = pymupdf.open(str(self.path))
        self._chars: dict[int, list[tuple[str, str, float, float]]] = {}
        self.frag_fonts: set[str] = set()
        self._scan_fonts()

    def close(self) -> None:
        self.doc.close()

    def _scan_fonts(self) -> None:
        """找出被误映射的符号字体：字符集极小、却反复出现的字体。"""
        stats: dict[str, list] = {}
        for page in self.doc:
            for span in page.get_texttrace():
                rec = stats.setdefault(span["font"], [0, set()])
                for char in span["chars"]:
                    rec[0] += 1
                    rec[1].add(chr(char[0]))
        for font, (count, charset) in stats.items():
            if count < 3 or len(charset) > 3:
                continue
            if any(c.isdigit() or c.isalpha() for c in charset):
                self.frag_fonts.add(font)

    def page_chars(self, pno: int) -> list[tuple[str, str, float, float, float]]:
        if pno in self._chars:
            return self._chars[pno]
        out = []
        for block in self.doc[pno].get_text("rawdict")["blocks"]:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    font = span.get("font", "")
                    for ch in span.get("chars", []):
                        c = ch.get("c", "")
                        if not c or c in ("\x00", "\ufffd"):
                            continue
                        if not c.isprintable() and c != " ":
                            continue
                        origin = ch.get("origin", (0.0, 0.0))
                        box = ch.get("bbox", (origin[0], origin[1], origin[0], origin[1]))
                        out.append((c, font, origin[0], origin[1],
                                    max(box[2] - box[0], 0.0)))
        self._chars[pno] = out
        return out


def _cluster_rows(items, tol: float = 3.0):
    rows: list[list] = []
    for item in sorted(items, key=lambda t: t[3]):
        for row in rows:
            if abs(row[0][3] - item[3]) <= tol:
                row.append(item)
                break
        else:
            rows.append([item])
    return rows


ODD_CHARS = set("{}|\\^$@")


def _render_line(items, frag_fonts: set[str]) -> tuple[str, bool, set[str]]:
    items = sorted(items, key=lambda t: t[2])

    # 伪粗体会把同一字符画两遍，位置几乎重合，只留一份
    kept: list = []
    for item in items:
        if (kept and item[0] == kept[-1][0]
                and abs(item[2] - kept[-1][2]) < 0.8
                and abs(item[3] - kept[-1][3]) < 1.0):
            continue
        kept.append(item)
    items = kept

    out: list[str] = []
    suspect = False
    notes: set[str] = set()

    for idx, (ch, font, _x, _y, _w) in enumerate(items):
        if font not in frag_fonts or not ch.isdigit():
            if ch in ODD_CHARS:
                suspect = True
                notes.add("\u7b26\u53f7\u53ef\u80fd\u9519\u4f4d\uff0c\u8bf7\u5bf9\u7167\u539f\u6587")
            out.append(ch)
            continue

        prev = next((items[j][0] for j in range(idx - 1, -1, -1)), "")
        nxt = next((items[j][0] for j in range(idx + 1, len(items))), "")
        if prev.isdigit() and nxt.isdigit():
            out.append("\u00b1")
            suspect = True
            notes.add("\u00b1 \u5df2\u81ea\u52a8\u8fd8\u539f")
        else:
            out.append(ch)
            suspect = True
            notes.add("\u5b58\u5728\u65e0\u6cd5\u8bc6\u522b\u7684\u7b26\u53f7")
    return "".join(out), suspect, notes


def rebuild_cell(chars, frag_fonts: set[str]) -> Cell:
    if not chars:
        return Cell()
    parts: list[str] = []
    suspect = False
    notes: set[str] = set()
    for row in _cluster_rows(chars):
        text, sus, nt = _render_line(row, frag_fonts)
        if text.strip():
            parts.append(text.strip())
        if sus:
            suspect = True
            notes |= nt
    return Cell(" ".join(parts), suspect, "\uff1b".join(sorted(notes)))


def _chars_in(pdf: PdfDocument, pno: int, rect) -> list:
    """按字符宽度中点判定归属，避免相邻单元格的字被吸进来。"""
    x0, y0, x1, y1 = rect
    out = []
    for item in pdf.page_chars(pno):
        _ch, _font, x, y, width = item
        center = x + width / 2
        if x0 - 0.5 <= center <= x1 + 0.5 and y0 - 0.6 <= y <= y1 + 0.6:
            out.append(item)
    return out


def extract_table(pdf: PdfDocument, pno: int, raw, seq: int) -> Table:
    data: list[list[Cell]] = []
    for r in range(raw.row_count):
        row: list[Cell] = []
        for c in range(raw.col_count):
            try:
                rect = raw.rows[r].cells[c]
            except (IndexError, TypeError, ValueError):
                row.append(Cell())
                continue
            if rect is None or rect == pymupdf.EMPTY_RECT():
                row.append(Cell())
                continue
            row.append(rebuild_cell(_chars_in(pdf, pno, rect), pdf.frag_fonts))
        data.append(row)

    label = detect_label(pdf, pno, raw.bbox) or f"\u7b2c{pno + 1}\u9875\u8868{seq}"
    return Table(
        page=pno + 1,
        seq=seq,
        label=label,
        nrows=raw.row_count,
        ncols=raw.col_count,
        bbox=tuple(raw.bbox),
        data=data,
    )


def detect_label(pdf: PdfDocument, pno: int, bbox) -> str:
    """从表格上方一小段区域里找 "Table 3" 这样的编号。"""
    page = pdf.doc[pno]
    top = max(bbox[1] - 26, 0)
    clip = pymupdf.Rect(bbox[0] - 4, top, bbox[2] + 4, bbox[1] - 1)
    if clip.is_empty:
        return ""
    try:
        text = page.get_textbox(clip)
    except Exception:
        return ""
    match = re.search(r"(Table|TABLE|\u8868)\s*([0-9IVX]+)", text)
    return f"Table {match.group(2).upper()}" if match else ""


def find_tables(pdf: PdfDocument) -> list[Table]:
    results: list[Table] = []
    for pno in range(pdf.doc.page_count):
        page = pdf.doc[pno]
        try:
            finder = page.find_tables(strategy="lines")
        except Exception:
            continue
        seq = 0
        for raw in finder.tables:
            seq += 1
            if raw.row_count < 2 or raw.col_count < 2:
                continue
            height = raw.bbox[3] - raw.bbox[1]
            if height > page.rect.height * 0.85:
                continue
            table = extract_table(pdf, pno, raw, seq)
            if sum(1 for row in table.data for c in row if c.text.strip()) < 4:
                continue
            table.source_name = pdf.path.name
            results.append(table)
    return results


def open_pdf(path: str | Path) -> PdfDocument:
    return PdfDocument(path)


def safe_sheet_name(name: str, used: set[str]) -> str:
    base = SHEET_BAD.sub(" ", name).strip() or "Sheet"
    base = base[:31]
    candidate = base
    n = 2
    while candidate in used:
        suffix = f"({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate
