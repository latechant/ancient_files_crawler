# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import requests

sys.stdout.reconfigure(encoding="utf-8")
OUT = Path(r"D:\AI_Project\ancient_files_crawler\pipeline\logs\_wiki_hunt.json")
S = requests.Session()
S.headers.update({"User-Agent": "ancient-files-research/1.0 (course archival)"})

QS = [
    "和平縣志",
    "仇公著墓誌銘",
    "大乘座主",
    "妙因寺",
    "襲三恪嗣韓國公",
    "貢桑諾爾布 政府公報",
    "巴拉根加布",
    "舒赫德 喀什噶爾",
    "福康安 伊犁官馬",
    "錡馬仁",
    "慧隆寺",
    "永垂場碑",
    "故導丞印",
    "嶺光電 憶往昔",
]


def main() -> None:
    rows = []
    for q in QS:
        r = S.get(
            "https://zh.wikisource.org/w/api.php",
            params={
                "action": "query",
                "list": "search",
                "srsearch": q,
                "srlimit": 6,
                "format": "json",
            },
            timeout=25,
        )
        r.raise_for_status()
        hits = [
            {"title": it.get("title"), "snippet": it.get("snippet", "")[:180]}
            for it in r.json().get("query", {}).get("search", [])
        ]
        rows.append({"q": q, "hits": hits})
        time.sleep(0.25)
    OUT.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print("wrote", len(rows))


if __name__ == "__main__":
    main()
