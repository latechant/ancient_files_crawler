# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

import requests

sys.stdout.reconfigure(encoding="utf-8")
API = "https://zh.wikisource.org/w/api.php"
s = requests.Session()
s.headers["User-Agent"] = "ancient-files-research/1.0"
for prefix in ["蛮书", "蠻書", "蛮书 (", "蠻書 (", "蛮书_"]:
    r = s.get(API, params={"action": "query", "list": "allpages", "apprefix": prefix, "aplimit": 20, "format": "json"}, timeout=30)
    titles = [p["title"] for p in r.json()["query"]["allpages"]]
    print("PREFIX", prefix, "N", len(titles))
    for t in titles[:15]:
        print(" ", t)

for title in ["蛮书 (四库全书本)", "蠻書 (四庫全書本)", "蛮书", "蠻書"]:
    r = s.get(API, params={"action": "query", "titles": title, "format": "json"}, timeout=30)
    print("TITLE", title, r.json()["query"]["pages"])
