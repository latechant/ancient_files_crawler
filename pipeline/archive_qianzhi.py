# -*- coding: utf-8 -*-
"""归档《遣使－唐朝派使者至吐蕃答贺》：母项《册府元龟》卷九百八十外臣部·通好。"""
from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(r"D:\AI_Project\ancient_files_crawler")
RESULT = ROOT / "韦广波" / "爬取结果"
PRODUCT = ROOT / "韦广波" / "成果"
EXCEL = ROOT / "第22组总览表.xlsx"
STAMP = datetime.now().strftime("%Y-%m-%d")

SEQ = 10736
TITLE = "遣使－唐朝派使者至吐蕃答贺"
FOLDER = "10736遣使－唐朝派使者至吐蕃答贺"
FILE = "册府元龟_卷九百八十_入吐蕃答贺正使.txt"
URL = "https://zh.wikisource.org/wiki/冊府元龜_(四庫全書本)/卷0980"
BODY = """# 母项：《册府元龟》卷九百八十 外臣部·通好（四库全书本）
# 子项：唐朝派遣入吐蕃答贺正使诸条

十月以岳王傅成杭為右庶子兼御史中丞充入吐蕃答賀正使仍賜紫金魚袋以太常博士劉幼復為殿中侍御史為之副仍賜緋魚袋以前萬年縣丞韓澥為大理丞兼殿中侍御史入吐蕃答賀正使判官仍賜緋魚袋。

文宗太和二年十二月以前棣州刺史唐𢎞實為莒王傅兼御史大夫持節充入吐蕃答賀正使。

六年十一月以少府少監田早守本官兼御史中丞持節充入吐蕃答賀正使仍賜紫金魚袋。

十一月以宗正少卿李從簡守本官兼御史中丞持節充入吐蕃答賀正使仍賜紫金魚袋。開成元年十二月以太子少詹事李景儒兼御史中丞充入吐蕃答賀正使。
"""
NOTE = (
    f"{STAMP} 条目为史事子项，非专书。母项《册府元龟》卷九百八十外臣部通好已核四库本："
    "记宝历元年十月成杭、太和二年唐弘实、六年田早、九年李从简、开成元年李景儒等充入吐蕃答贺正使。"
    "已摘录对应原文，不替代全书。"
)


def dump() -> None:
    header = f"# {TITLE}\n# 来源: {URL}\n# 编码: UTF-8\n# 归档日期: {STAMP}\n\n"
    content = header + BODY.strip() + "\n"
    folder = RESULT / FOLDER
    product = PRODUCT / TITLE
    folder.mkdir(parents=True, exist_ok=True)
    product.mkdir(parents=True, exist_ok=True)
    (folder / FILE).write_text(content, encoding="utf-8")
    (product / FILE).write_text(content, encoding="utf-8")
    (folder / "来源网址.txt").write_text(URL + "\n", encoding="utf-8")
    (product / "来源网址.txt").write_text(URL + "\n", encoding="utf-8")
    meta = {
        "seq": SEQ,
        "title": TITLE,
        "url": URL,
        "publisher": "不详",
        "edition": "宋王钦若等《册府元龟》四库全书本卷九百八十外臣部通好。",
        "note": NOTE,
        "saved_at": STAMP,
        "ok": True,
    }
    (folder / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (product / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("saved files", SEQ)


def mark_excel() -> None:
    last = None
    for i in range(10):
        try:
            wb = load_workbook(EXCEL)
            ws = wb[wb.sheetnames[0]]
            for row in ws.iter_rows(min_row=5):
                if row[0].value != "韦广波":
                    continue
                if row[1].value != SEQ:
                    continue
                row[5].value = URL
                row[6].value = "是"
                row[7].value = "不详"
                row[8].value = "《册府元龟》四库本卷九百八十外臣部通好；子项摘录入吐蕃答贺正使。"
                old = str(row[9].value or "").strip()
                row[9].value = NOTE if old in ("", "未找到", "无") else f"{old}；{NOTE}"
            wb.save(EXCEL)
            print("excel ok")
            return
        except Exception as exc:
            last = exc
            print("excel retry", i, type(exc).__name__)
            time.sleep(1.5)
    raise last  # type: ignore[misc]


if __name__ == "__main__":
    dump()
    mark_excel()
