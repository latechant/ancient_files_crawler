# -*- coding: utf-8 -*-
"""扫识典金石书目录，找剩余墓志/碑名。"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"D:\AI_Project\ancient_files_crawler")
from pipeline.extract_shidian import book_payload, fetch_html, flatten_chapters, load_router

CACHE = Path(r"D:\AI_Project\ancient_files_crawler\_cache\shidian")
OUT = Path(r"D:\AI_Project\ancient_files_crawler\pipeline\logs\_catalog_hits.txt")

BOOKS = [
    ("7549461577028599854", "1lcvz8ltnwv0q", "八琼室续编"),
    ("CADAL01026682", "1lmm82h8zayib", "山右石刻丛编"),
    ("NGJ892411999028946160395", "1lqd64lans6qi", "益都金石略"),
    ("7659974904619466767", "1m73u9ulbvynn", "金石三编"),
]

NEEDLES = (
    "仇公著",
    "大乘座主",
    "妙因寺",
    "三恪",
    "韩国公",
    "元公墓",
    "李继忠",
    "锜马仁",
    "锜麻仁",
    "关氏",
    "永垂场",
    "摆金",
    "慧隆寺",
)


def main() -> None:
    lines = []
    for book_id, seed, label in BOOKS:
        url = f"https://www.shidianguji.com/book/{book_id}/chapter/{seed}"
        html = fetch_html(url, CACHE / f"{book_id}_{seed}.html")
        payload = book_payload(load_router(html))
        info = payload["bookInfo"]
        chapters = flatten_chapters(info.get("catalog", {}).get("chapters") or [])
        lines.append(f"=== {label} {info.get('bookName')} chapters={len(chapters)} ===")
        for ch in chapters:
            name = ch.get("chapterName") or ""
            if any(n in name for n in NEEDLES):
                lines.append(f"{label}\t{ch.get('chapterId')}\t{name}")
        # also dump all names containing 墓志
        mz = [ch for ch in chapters if "墓" in (ch.get("chapterName") or "")]
        lines.append(f"-- {label} 墓* {len(mz)} --")
        for ch in mz[:80]:
            lines.append(f"{label}\t{ch.get('chapterId')}\t{ch.get('chapterName')}")
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print("wrote", OUT, "lines", len(lines))


if __name__ == "__main__":
    main()
