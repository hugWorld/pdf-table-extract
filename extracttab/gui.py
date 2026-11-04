"""extractTab 图形界面（PyWebView + 本地 HTML）。"""

from __future__ import annotations

import json
import os
import tempfile
import time
from pathlib import Path

import webview

from .api import Api

WINDOW_TITLE = "\u8868\u683c\u63d0\u53d6"


def _storage_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    path = Path(base) / "extractTab" / "webview"
    path.mkdir(parents=True, exist_ok=True)
    return path


def main(argv: list[str] | None = None) -> None:
    argv = argv or []
    files = [str(Path(a).resolve()) for a in argv
             if Path(a).suffix.lower() == ".pdf" and Path(a).exists()]

    api = Api()
    html = Path(__file__).with_name("ui.html")
    window = webview.create_window(
        WINDOW_TITLE,
        url=str(html),
        js_api=api,
        width=960,
        height=680,
        min_size=(800, 560),
        background_color="#ffffff",
    )
    api.bind(window)

    def bootstrap() -> None:
        if not files:
            return
        payload = json.dumps(files, ensure_ascii=False)
        script = (
            "(function(){var n=0;var t=setInterval(function(){"
            "if(window.pywebview&&window.pywebview.api){clearInterval(t);"
            f"addFiles({payload});"
            "}else if(++n>60){clearInterval(t);}},100);})();"
        )
        for _ in range(40):
            try:
                window.evaluate_js(script)
                return
            except Exception:
                time.sleep(0.1)

    webview.start(bootstrap, private_mode=False,
                  storage_path=str(_storage_dir()), gui="edgechromium")
