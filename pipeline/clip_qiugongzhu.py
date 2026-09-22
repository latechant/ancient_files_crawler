# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

LOG = Path(r"D:\AI_Project\ancient_files_crawler\pipeline\logs")
src = (LOG / "_bajiong_v32.txt").read_text(encoding="utf-8")
start = src.find("仇公著墓誌銘")
if start < 0:
    start = src.find("仇公著")
# Prefer the catalog heading that includes dimensions.
head = src.rfind("栝", 0, start + 1)
if head >= 0 and start - head < 80:
    start = head
chunk = src[start:]
# Next major item after the 仇志 is 徽宗八行八刑条 / 邢台碑.
cut_marks = (
    "右徽宗御製",
    "徽宗御製學校",
    "學校八行",
    "八行、八刑",
    "邢臺",
)
cuts = []
for m in cut_marks:
    i = chunk.find(m, 200)
    if i > 0:
        cuts.append((i, m))
cuts.sort()
if cuts:
    chunk = chunk[: cuts[0][0]]
# If still too long, cut at the next 石刻 heading-like line after a 跋.
lines = chunk.splitlines()
kept = []
seen_title = False
for line in lines:
    if "宋故定州觀察判官仇府君墓誌銘" in line:
        seen_title = True
    if seen_title and kept and (
        line.startswith("右徽宗")
        or "學校八行" in line
        or (len(line) < 40 and "碑" in line and "仇" not in line and "墓" not in line)
    ):
        break
    kept.append(line)
text = "\n".join(kept).strip() + "\n"
(LOG / "_qiugongzhu.txt").write_text(text, encoding="utf-8")
preview = []
for n, line in enumerate(text.splitlines(), 1):
    preview.append(f"{n:03d}|{line}")
(LOG / "_qiugongzhu_preview.txt").write_text("\n".join(preview), encoding="utf-8")
meta = [
    f"chars={len(text)}",
    f"lines={len(kept)}",
    f"first_cut={cuts[0] if cuts else None}",
    f"has_府君={('仇府君' in text)}",
    f"has_柳子文={('柳子文' in text)}",
    f"has_绍圣={('紹聖' in text or '绍圣' in text)}",
    f"has_云门={('雲門' in text or '云门' in text)}",
    f"has_念={('子念' in text or '子曰念' in text or '男念' in text)}",
]
(LOG / "_qiugongzhu_meta.txt").write_text("\n".join(meta), encoding="utf-8")
