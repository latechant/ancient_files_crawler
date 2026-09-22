# -*- coding: utf-8 -*-
"""在识典检索页、维基文库 API 上扫剩余高价值条目。"""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from urllib.parse import quote

import requests

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(r"D:\AI_Project\ancient_files_crawler")
OUT = ROOT / "pipeline" / "logs" / "_alt_hunt.json"

UA = "Mozilla/5.0 (compatible; ancient-files-research/1.0; +local archival)"
S = requests.Session()
S.headers.update({"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"})

QUERIES = [
    ("14071", "仇公著墓志铭", "shidian"),
    ("12662", "大乘座主墓志", "both"),
    ("12662", "妙因寺 比丘尼", "both"),
    ("13815", "袭三恪嗣韩国公 墓志", "both"),
    ("13815", "唐故袭三恪嗣韩国公", "both"),
    ("12021", "舒赫德 暂往喀什噶尔", "shidian"),
    ("12022", "福康安 伊犁官马", "shidian"),
    ("8941", "贺丞相墓志铭", "skip"),
    ("13047", "贡桑诺尔布 巴拉根加布", "wiki"),
    ("13047", "蒙藏院总裁贡桑诺尔布呈", "wiki"),
    ("196", "和平縣志", "wiki"),
    ("196", "康熙和平县志", "shidian"),
    ("844", "永垂场碑 摆金", "both"),
    ("164", "关氏墓碑 宁夏", "both"),
    ("348", "锜马仁造像", "shidian"),
    ("276", "正柱 佃户霸地 庄头", "shidian"),
    ("396", "伯兴 严禁私藏鸟枪", "shidian"),
    ("452", "精河屯田 范时绶", "shidian"),
]


def shidian_search(q: str) -> dict:
    url = f"https://www.shidianguji.com/search?query={quote(q)}"
    try:
        r = S.get(url, timeout=30)
        html = r.text
        hits = []
        # 粗抓书名/章节锚点
        for token in ("/book/", "/chapter/", "bookName", "chapterName"):
            if token in html:
                hits.append(token)
        return {
            "ok": r.status_code == 200,
            "status": r.status_code,
            "chars": len(html),
            "has_router": "window._ROUTER_DATA" in html,
            "tokens": hits,
            "url": url,
        }
    except Exception as exc:
        return {"ok": False, "error": type(exc).__name__, "url": url}


def wiki_search(q: str) -> dict:
    try:
        r = S.get(
            "https://zh.wikisource.org/w/api.php",
            params={
                "action": "query",
                "list": "search",
                "srsearch": q,
                "srlimit": 8,
                "format": "json",
            },
            timeout=30,
        )
        r.raise_for_status()
        hits = [
            {"title": it.get("title"), "snippet": it.get("snippet")}
            for it in r.json().get("query", {}).get("search", [])
        ]
        return {"ok": True, "hits": hits}
    except Exception as exc:
        return {"ok": False, "error": type(exc).__name__}


def main() -> None:
    rows = []
    for seq, q, where in QUERIES:
        if where == "skip":
            continue
        item = {"seq": seq, "query": q}
        if where in ("both", "shidian"):
            item["shidian"] = shidian_search(q)
            time.sleep(0.6)
        if where in ("both", "wiki"):
            item["wiki"] = wiki_search(q)
            time.sleep(0.4)
        rows.append(item)
        print("done", seq, q)
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print("wrote", OUT, "n=", len(rows))


if __name__ == "__main__":
    main()
