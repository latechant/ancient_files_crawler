# -*- coding: utf-8 -*-
"""韦广波本轮：母项方志/别名书批量抽取，归档 TXT 并回写总览表。"""
from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"D:\AI_Project\ancient_files_crawler")
sys.path.insert(0, str(ROOT))

from pipeline.extract_shidian import extract_book
from pipeline.extract_wikisource import extract_prefixed, get_wikitext

sys.stdout.reconfigure(encoding="utf-8")

WEI_DIR = ROOT / "韦广波"
RESULT = WEI_DIR / "爬取结果"
PRODUCT = WEI_DIR / "成果"
LOG_DIR = ROOT / "pipeline" / "logs"
CACHE = ROOT / "_cache" / "shidian"
EXCEL = ROOT / "第22组总览表.xlsx"
STAMP = datetime.now().strftime("%Y-%m-%d")

BOOKS = [
    {
        "seq": 1234,
        "title": "灌记初稿",
        "folder": "1234灌记初稿",
        "kind": "shidian",
        "book_id": "NLG311998004773",
        "seed": "1l2cksxknufbg",
        "publisher": "不详",
        "edition": "清光绪二十年(1894)刻本,4卷;彭洵纂。识典古籍另析为5卷（含目录/拾遗卷）。",
        "note": f"{STAMP} 识典古籍全文已按卷提取；母项为《灌记初稿》。",
    },
    {
        "seq": 3686,
        "title": "马边厅志略",
        "folder": "3686马边厅志略",
        "kind": "shidian",
        "book_id": "CADAL11201307",
        "seed": "1l3b38fba6adc",
        "publisher": "不详",
        "edition": "清嘉庆刻本,6卷首一卷;周斯才纂。CADAL/识典古籍本。",
        "note": f"{STAMP} 识典古籍 CADAL11201307 全文已按序、目录、凡例、卷一至卷六及钞校记提取。",
    },
    {
        "seq": 2402,
        "title": "榆林府志",
        "folder": "2402榆林府志",
        "kind": "shidian",
        "book_id": "NLG312001065669",
        "seed": "1l6a853hg3nz7",
        "publisher": "不详",
        "edition": "清道光二十一年(1841)刻本,50卷;李熙龄纂修。识典古籍转录本。",
        "note": f"{STAMP} 识典古籍全文已按目录卷次提取。原表 gxhj 链接无效，已替换为可核原文。",
    },
    {
        "seq": 3815,
        "title": "云南记",
        "folder": "3815云南记",
        "kind": "wikisource_man_shu",
        "publisher": "武英殿聚珍版/四库全书本",
        "edition": "唐樊绰《蛮书》十卷，亦名云南志、云南史记；袁滋《云南记》已佚。",
        "note": (
            f"{STAMP} 袁滋《云南记》原书已佚，仅类书零星佚文。"
            "现存可核全文为樊绰《蛮书》四库本（别名云南志/云南史记，部分文献亦称云南记）。"
            "本条按别名母项归档四库本《蛮书》，不能径视为袁滋原书。"
        ),
    },
]


def safe_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name).strip()
    return name[:80] or "未名"


def write_sources(path: Path, urls: list[str], extra: str = "") -> None:
    lines = urls + ([extra] if extra else [])
    path.write_text("\n".join(dict.fromkeys(lines)) + "\n", encoding="utf-8")


def write_meta(path: Path, rec: dict, extra: dict) -> None:
    payload = {
        "seq": rec["seq"],
        "title": rec["title"],
        "owner": "韦广波",
        "saved_at": STAMP,
        **extra,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def archive_texts(folder: Path, product_dir: Path, texts: dict[str, str]) -> int:
    folder.mkdir(parents=True, exist_ok=True)
    product_dir.mkdir(parents=True, exist_ok=True)
    n = 0
    for title, body in texts.items():
        if not body or not body.strip():
            continue
        fname = safe_filename(title) + ".txt"
        header = f"# {title}\n# 编码: UTF-8\n# 归档日期: {STAMP}\n\n"
        content = header + body.strip() + "\n"
        (folder / fname).write_text(content, encoding="utf-8")
        (product_dir / fname).write_text(content, encoding="utf-8")
        n += 1
    return n


def save_shidian(rec: dict) -> dict:
    data = extract_book(rec["book_id"], rec["seed"], CACHE / rec["folder"])
    folder = RESULT / rec["folder"]
    product = PRODUCT / rec["title"]
    n = archive_texts(folder, product, data["texts"])
    write_sources(folder / "来源网址.txt", data["sources"])
    write_sources(product / "来源网址.txt", data["sources"])
    meta = {
        "source": data["seed_url"],
        "bookId": data["bookId"],
        "bookName": data["bookName"],
        "chapter_count": len(data["chapters"]),
        "saved_txt": n,
        "ok": n > 0,
        "url": f"https://www.shidianguji.com/book/{rec['book_id']}",
    }
    write_meta(folder / "基本信息.json", rec, meta)
    write_meta(product / "基本信息.json", rec, meta)
    return meta


def save_manshu(rec: dict) -> dict:
    prefix = "蛮书 (四库全书本)"
    data = extract_prefixed(prefix)
    # Keep volume pages; skip talk/redirects.
    texts = {}
    for title, body in data["texts"].items():
        short = title.replace("蛮书 (四库全书本)/", "蛮书_").replace("蛮书 (四库全书本)", "蛮书_目录")
        texts[short] = body
    folder = RESULT / rec["folder"]
    product = PRODUCT / rec["title"]
    n = archive_texts(folder, product, texts)
    sources = [
        "https://zh.wikisource.org/wiki/蛮书_(四库全书本)",
        "https://zh.wikisource.org/wiki/蛮书_(四库全书本)/全览",
        "https://zh.wikipedia.org/wiki/云南志",
    ]
    write_sources(folder / "来源网址.txt", sources)
    write_sources(product / "来源网址.txt", sources)
    (folder / "判定说明.txt").write_text(rec["note"] + "\n", encoding="utf-8")
    (product / "判定说明.txt").write_text(rec["note"] + "\n", encoding="utf-8")
    meta = {
        "source": sources[0],
        "saved_txt": n,
        "ok": n > 0,
        "url": sources[0],
        "pages": data["titles"],
    }
    write_meta(folder / "基本信息.json", rec, meta)
    write_meta(product / "基本信息.json", rec, meta)
    return meta


def update_excel(updates: list[dict]) -> None:
    wb = load_workbook(EXCEL)
    ws = wb[wb.sheetnames[0]]
    by_seq = {u["seq"]: u for u in updates}
    changed = 0
    for row in ws.iter_rows(min_row=5):
        owner = row[0].value
        seq = row[1].value
        if owner != "韦广波" or seq not in by_seq:
            continue
        u = by_seq[seq]
        row[5].value = u["url"]
        row[6].value = "是" if u.get("ok") else "否"
        if u.get("publisher"):
            row[7].value = u["publisher"]
        if u.get("edition"):
            row[8].value = u["edition"]
        old = str(row[9].value or "").strip()
        note = u["note"]
        row[9].value = note if not old or old in ("未找到", "无") else f"{old}；{note}"
        changed += 1
    wb.save(EXCEL)
    print("excel_updated_rows", changed)


def main() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    RESULT.mkdir(parents=True, exist_ok=True)
    PRODUCT.mkdir(parents=True, exist_ok=True)
    results = []
    for rec in BOOKS:
        print("START", rec["seq"], rec["title"], rec["kind"])
        try:
            if rec["kind"] == "shidian":
                meta = save_shidian(rec)
            elif rec["kind"] == "wikisource_man_shu":
                meta = save_manshu(rec)
            else:
                raise ValueError(rec["kind"])
            rec_out = {**rec, **meta}
            results.append(rec_out)
            print("OK", rec["seq"], "txt", meta.get("saved_txt"), "url", meta.get("url"))
        except Exception as e:
            print("FAIL", rec["seq"], rec["title"], type(e).__name__, e)
            results.append({**rec, "ok": False, "error": str(e), "url": rec.get("url", "无")})
    update_excel([r for r in results if r.get("url")])
    log_path = LOG_DIR / f"wei_batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    log_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print("log", log_path)


if __name__ == "__main__":
    main()
