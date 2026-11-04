"""暴露给前端的后端接口。"""

from __future__ import annotations

import os
import traceback
from pathlib import Path

import webview

from .core import Table, find_tables, open_pdf
from .export import write_workbook


class Api:
    def __init__(self) -> None:
        self.window: webview.Window | None = None
        self.tables: dict[str, Table] = {}
        self.paths: list[Path] = []
        self._counter = 0

    def bind(self, window: webview.Window) -> None:
        self.window = window

    def _new_uid(self) -> str:
        self._counter += 1
        return f"t{self._counter}"

    def pick_files(self) -> list[str]:
        if not self.window:
            return []
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True,
            file_types=("PDF 文件 (*.pdf)", "所有文件 (*.*)"))
        return [str(p) for p in (result or [])]

    def load(self, paths: list[str]) -> dict:
        summaries = []
        errors = []
        for raw in paths:
            path = Path(raw)
            if not path.exists() or path.suffix.lower() != ".pdf":
                errors.append(f"{path.name}：不是可用的 PDF")
                continue
            if path in self.paths:
                errors.append(f"{path.name}：已经加载过了")
                continue
            try:
                pdf = open_pdf(path)
                try:
                    found = find_tables(pdf)
                finally:
                    pdf.close()
            except Exception as exc:
                errors.append(f"{path.name}：{exc}")
                continue

            if path not in self.paths:
                self.paths.append(path)
            for table in found:
                uid = self._new_uid()
                self.tables[uid] = table
                summaries.append({
                    "uid": uid,
                    "label": table.label,
                    "page": table.page,
                    "rows": table.nrows,
                    "cols": table.ncols,
                    "suspect": table.suspect_count,
                    "numeric": round(table.numeric_ratio, 3),
                    "file": path.name,
                })
        return {"tables": summaries, "errors": errors}

    def preview(self, uid: str) -> dict:
        table = self.tables.get(uid)
        if not table:
            return {"rows": [], "suspect": []}
        return {
            "rows": [[cell.text for cell in row] for row in table.data],
            "suspect": [[cell.suspect for cell in row] for row in table.data],
            "notes": sorted({cell.note for row in table.data
                             for cell in row if cell.note}),
        }

    def export(self, uids: list[str], with_source: bool) -> dict:
        chosen = [self.tables[u] for u in uids if u in self.tables]
        if not chosen:
            return {"ok": False, "message": "没有勾选任何表格"}
        if not self.paths:
            stem = "tables"
        elif len(self.paths) == 1:
            stem = self.paths[0].stem
        else:
            stem = f"{self.paths[0].stem}等{len(self.paths)}篇"
        if not self.window:
            return {"ok": False, "message": "窗口未就绪"}
        target = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=f"{stem}_表格.xlsx",
            file_types=("Excel 工作簿 (*.xlsx)",))
        if not target:
            return {"ok": False, "message": ""}
        path = target if isinstance(target, str) else target[0]
        try:
            out = write_workbook(chosen, path, with_source=with_source)
        except Exception:
            return {"ok": False, "message": traceback.format_exc(limit=3)}
        flagged = sum(1 for t in chosen if t.suspect_count)
        return {
            "ok": True,
            "path": str(out),
            "count": len(chosen),
            "flagged": flagged,
            "name": out.name,
        }

    def open_path(self, path: str) -> None:
        try:
            os.startfile(path)
        except Exception:
            pass

    def reveal(self, path: str) -> None:
        try:
            os.startfile(str(Path(path).parent))
        except Exception:
            pass
