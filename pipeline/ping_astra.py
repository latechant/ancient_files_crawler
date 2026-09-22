# -*- coding: utf-8 -*-
"""Ping aizex gpt-6-astra without printing secrets."""
from __future__ import annotations

import sys
from pathlib import Path

APM = Path(r"D:\AI_Project\automatic_problem_maker")
sys.path.insert(0, str(APM))

from config import get_api_key, get_base_url  # noqa: E402
from tools.llm_client import call_llm  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
key = get_api_key()
print("has_key", bool(key), "key_len", len(key))
print("base_url", get_base_url())
try:
    text = call_llm(
        "只回复：astra-ok",
        system_prompt="你是连通性探针，只回四个字符以内的确认。",
        temperature=0,
        timeout_seconds=90,
        model_name_override="gpt-6-astra",
        max_tokens=32,
    )
    print("ok", True)
    print("reply_len", len(text or ""))
    print("reply", (text or "")[:80].replace("\n", " "))
except Exception as exc:
    print("ok", False)
    print("error_type", type(exc).__name__)
    print("error", str(exc)[:400])
