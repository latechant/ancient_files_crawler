# -*- coding: utf-8 -*-
"""维基文库 MediaWiki API 抽取（公有领域古籍）。"""
from __future__ import annotations

import time
from typing import Any

import requests

API = "https://zh.wikisource.org/w/api.php"
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "ancient-files-research/1.0 (course archival)"})


def get_wikitext(title: str) -> str:
    r = SESSION.get(
        API,
        params={
            "action": "query",
            "prop": "revisions",
            "rvprop": "content",
            "rvslots": "main",
            "titles": title,
            "format": "json",
        },
        timeout=40,
    )
    r.raise_for_status()
    pages = r.json()["query"]["pages"]
    page = next(iter(pages.values()))
    if "missing" in page:
        return ""
    return page["revisions"][0]["slots"]["main"]["*"]


def list_prefix(prefix: str, limit: int = 50) -> list[str]:
    r = SESSION.get(
        API,
        params={
            "action": "query",
            "list": "allpages",
            "apprefix": prefix,
            "aplimit": limit,
            "format": "json",
        },
        timeout=40,
    )
    r.raise_for_status()
    return [p["title"] for p in r.json()["query"]["allpages"]]


def extract_prefixed(prefix: str, sleep_s: float = 0.4) -> dict[str, Any]:
    titles = list_prefix(prefix)
    texts: dict[str, str] = {}
    for i, title in enumerate(titles):
        texts[title] = get_wikitext(title)
        if i + 1 < len(titles):
            time.sleep(sleep_s)
    return {"prefix": prefix, "titles": titles, "texts": texts}
