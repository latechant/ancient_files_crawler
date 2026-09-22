# -*- coding: utf-8 -*-
"""从识典《八琼室金石补正续编》卷三十二抽出《宋仇公著墓志》正文。"""
from __future__ import annotations

import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"D:\AI_Project\ancient_files_crawler")

from pipeline.extract_shidian import book_payload, fetch_html, load_router, paragraph_text

CACHE = Path(r"D:\AI_Project\ancient_files_crawler\_cache")
LOG = Path(r"D:\AI_Project\ancient_files_crawler\pipeline\logs")
URL = "https://www.shidianguji.com/book/7549461577028599854/chapter/1lcvz8ltnwv0q"


def main() -> None:
    html = fetch_html(URL, CACHE / "shidian" / "bajiong_v32.html")
    payload = book_payload(load_router(html))
    paras = payload.get("paragraphList") or []
    body = "\n".join(t for t in (paragraph_text(p).strip() for p in paras) if t)
    LOG.mkdir(parents=True, exist_ok=True)
    (LOG / "_bajiong_v32.txt").write_text(body, encoding="utf-8")
    idx = body.find("仇公著")
    if idx < 0:
        raise SystemExit("仇公著 not found")
    chunk = body[max(0, idx - 80) :]
    ends = []
    for mark in ("龙兴寺大悲阁记", "龙兴寺大悲阁", "北岳大殿", "右宋故定州"):
        j = chunk.find(mark, 200)
        if j > 0:
            ends.append(j)
    if ends:
        chunk = chunk[: min(ends)]
    (LOG / "_qiugongzhu.txt").write_text(chunk.strip() + "\n", encoding="utf-8")
    print("chapter_chars", len(body))
    print("clip_chars", len(chunk))
    print("clip_head", chunk[:180].replace("\n", " | "))
    print("clip_tail", chunk[-180:].replace("\n", " | "))


if __name__ == "__main__":
    main()
