# -*- coding: utf-8 -*-
"""识典古籍 SSR 抽取：解析 window._ROUTER_DATA，按卷输出 UTF-8 TXT。"""
from __future__ import annotations

import json
import re
import time
from pathlib import Path
from typing import Any

import requests

UA = "Mozilla/5.0 (compatible; ancient-files-research/1.0; +local archival)"
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"})


def load_router(html: str) -> dict[str, Any]:
    start = html.find("window._ROUTER_DATA = ")
    if start < 0:
        raise ValueError("page has no _ROUTER_DATA")
    start += len("window._ROUTER_DATA = ")
    end = html.find("</script>", start)
    raw = html[start:end].strip().rstrip(";")
    return json.loads(raw)


def book_payload(router: dict[str, Any]) -> dict[str, Any]:
    return router["loaderData"]["__session/(lang$)/book/$"]


def chapter_name(ch: dict[str, Any]) -> str:
    name = ch.get("chapterName")
    if isinstance(name, list):
        parts = []
        for n in name:
            if isinstance(n, dict):
                parts.append(str(n.get("content") or ""))
            else:
                parts.append(str(n))
        return "".join(parts).strip()
    return str(name or "").strip()


def flatten_chapters(chaps: list[dict[str, Any]] | None, acc: list[dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    acc = acc if acc is not None else []
    for ch in chaps or []:
        acc.append(
            {
                "chapterId": ch.get("chapterId"),
                "chapterName": chapter_name(ch),
                "hasMainContent": bool(ch.get("hasMainContent")),
            }
        )
        flatten_chapters(ch.get("subChapters") or [], acc)
    return acc


def paragraph_text(p: dict[str, Any]) -> str:
    raw = p.get("content")
    if raw is None:
        return ""
    if isinstance(raw, str):
        try:
            obj = json.loads(raw)
        except json.JSONDecodeError:
            return raw
    else:
        obj = raw
    if isinstance(obj, dict):
        lines = obj.get("lines") or []
        return "".join(str(line.get("content") or "") for line in lines if isinstance(line, dict))
    return str(obj)


def fetch_html(url: str, cache_path: Path, timeout: int = 40) -> str:
    if cache_path.exists() and cache_path.stat().st_size > 1000:
        return cache_path.read_text(encoding="utf-8")
    r = SESSION.get(url, timeout=timeout)
    r.raise_for_status()
    r.encoding = "utf-8"
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(r.text, encoding="utf-8")
    return r.text


def extract_book(book_id: str, seed_chapter_id: str, cache_dir: Path, sleep_s: float = 0.8) -> dict[str, Any]:
    cache_dir.mkdir(parents=True, exist_ok=True)
    seed_url = f"https://www.shidianguji.com/book/{book_id}/chapter/{seed_chapter_id}"
    seed_html = fetch_html(seed_url, cache_dir / f"{book_id}_{seed_chapter_id}.html")
    payload = book_payload(load_router(seed_html))
    info = payload["bookInfo"]
    chapters = flatten_chapters(info.get("catalog", {}).get("chapters") or [])
    texts: dict[str, str] = {}
    sources: list[str] = [f"https://www.shidianguji.com/book/{book_id}"]
    for i, ch in enumerate(chapters):
        cid = ch["chapterId"]
        if not cid or not ch["hasMainContent"]:
            continue
        url = f"https://www.shidianguji.com/book/{book_id}/chapter/{cid}"
        sources.append(url)
        html = fetch_html(url, cache_dir / f"{book_id}_{cid}.html")
        ch_payload = book_payload(load_router(html))
        paragraphs = ch_payload.get("paragraphList") or []
        body = "\n".join(t for t in (paragraph_text(p).strip() for p in paragraphs) if t)
        title = ch["chapterName"] or cid
        texts[title] = body
        if i + 1 < len(chapters):
            time.sleep(sleep_s)
    return {
        "bookId": info.get("bookId") or book_id,
        "bookName": info.get("bookName"),
        "chapters": chapters,
        "texts": texts,
        "sources": sources,
        "seed_url": seed_url,
    }
