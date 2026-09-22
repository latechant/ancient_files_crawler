# -*- coding: utf-8 -*-
"""本轮核验归档：墓志录文 + 咸丰敕谕汉译文。"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(r"D:\AI_Project\ancient_files_crawler")
RESULT = ROOT / "韦广波" / "爬取结果"
PRODUCT = ROOT / "韦广波" / "成果"
EXCEL = ROOT / "第22组总览表.xlsx"
STAMP = datetime.now().strftime("%Y-%m-%d")

ZHANG = """夫人姓張，諱玉憐，齊國西安人也。父慶□，本國大中正。深源峻遠，冑自炎皇。子房處漢，秩窮袞命；茂先在晉，位實台鉉。自茲已降，龜組相尋。夫人幼而明惠，容德絕倫，孝敬自天，貞風內潤。事父母以孝謹著稱，撫弟妹以仁惠垂問。笄年既及，為文貞侯所娉。事舅奉姑，懃孝有聞。承郎接妹，婉順見美。故以流譽於鄉黨，傳芳於州閭矣。文侯少宦，夫人留居奉養，溫清視膳，廿餘載。性不姤忌，寤寐思賢，撫視庶子，同之自生，降恩厚澤，平等無二。正光二年，入朝長信宮，引見顯陽殿。神儀偉穆，進止閑庠。皇太后為之動容，歎曰：令問令望，如珪如璋，其此婦之謂乎。夫才於朝，妻賢於室，頊在一房矣。孝昌中，文侯薨徂。子女煢稚。夫人慈撫訓導，咸得成立。居家理治，嚴明著稱。推尚佛法，深解空相，大悲動心，惟慕慈善，聞聲見形，不食其肉，三長六短，齊誡不爽。福善徒施，上天不弔。以魏天平三年正月乙□奄焉薨逝。道俗行泣，少長進慕。亡上之痛，異口同音；人百之謠，朝野僉詠。嗚呼哀哉！粵四年二月丙寅合葬於黃山文侯之□陵。謹述餘芳，傳美金石。其辭曰：
女德幼成，婦容夙彰。金□稟潔，令問令望。如日之曜，如蘭之芳。蘭芬□何，芬似桂林。峨峨貞峻，婉孌德音。有信有行，如玉如金。金玉惟何，含弘清慎。四德□炳，慈明內潤。烈烈貞覶，溫溫婉婉。順順如□，□□□□。白日照照，長夜攸攸。如何奄忽，□□春秋。□儀永淪，餘芳空留。
"""

XIANFENG = """奉天承运皇帝诏曰：
谕达赖喇嘛之呼毕勒罕。朕统驭天下，惟期率土之民，各享安居之福，推兴政教。顷据驻藏大臣奏称：呼毕勒罕尔降生以来，均有吉祥佳兆，性貌端庄，举止大方，熟谙经典，能辨识前世达赖喇嘛所供佛尊用器，卫藏僧俗众人之所见所闻，皆倾心归顺，故择正月十三日，令布达拉宫众呼图克图预先诵经七日，热振呼图克图等虔诚祝祷诵经之后，与驻藏大臣一同于布达拉宫高宗纯皇帝圣像之前掣签，由金奔巴瓶掣出尔名等情。朕览奏，甚为嘉悦，即降谕令尔为达赖喇嘛之呼毕勒罕。且又降谕特遣驻藏办事大臣满庆、恩庆令呼必勒罕尔于布达拉宫坐床等因咨文。将赏尔之诸项物件交付理藩院司员携往以外，并逐项缮写汉字清单咨呈满庆，抵达之后呼必勒罕尔祇领之。今呼必勒罕正值聪慧长进之时，理宜感戴朕施鸿恩，专心学习所有经典，弘扬黄教，俾僧俗众人均获安康太平，勿得怠惰。特谕。
咸丰十年十月十一日
"""

LIU = """魏故咸陽太守劉府君墓誌銘
君諱玉，字天寶，弘農胡城人也。厥初基冑與日月同開，爵封次弟通君臣之始。周秦大漢，并班名位。遠祖司徒寬之苗。其中易世，舉一足明。值漢中譏凶奴之患，李陵出討，軍勢不利，遂沒虜廷。先人祖宗，便習其俗，婚姻官帶，與之錯雜。大魏開建，託定恒代，以曾祖初萬頭，大族之冑，宜履名宦，從駕之眾，理須督率，依地置官，為何渾地汗。爾時此斑，例亞州牧。義成王南討萇安，以祖可洛侯名家之孫，召接為副，充子都將。與王策謀幃內，制定雍境，遂以土荒，即今鎮押。君數世重蔭，成應引內，為西征子都。出祺之挺，屢有薰跡，宜可昇接，錫之矛土。假咸陽太守。春秋七十八，以孝昌三年歲次丙午十一月廿四日卒於家。肇基雲景，神綵重映，是曰劉族，世立堅明。位綖台司，志含中貞，翼輔王室，唯安唯寧。弈踵相繼，其器易新，召蒞矛土，四裔來冥。綏接恩化，富壤殷民，體含玉潔，不磨自鄰。
"""

FOUND = [
    {
        "seq": 13687,
        "title": "崔鸿妻张玉怜墓志",
        "folder": "13687崔鸿妻张玉怜墓志",
        "file": "夫人张玉怜墓志.txt",
        "body": ZHANG,
        "url": "https://zh.wikipedia.org/wiki/崔鴻",
        "publisher": "天津古籍出版社",
        "edition": "赵超《汉魏南北朝墓志汇编》录文；北魏天平四年合葬。",
        "note": f"{STAMP} 母项《汉魏南北朝墓志汇编》。已核录文并归档。出土线索见《考古学报》1984年第2期《临淄北朝崔氏墓》。",
    },
    {
        "seq": 11123,
        "title": "咸丰皇帝为坐床事给十二世达赖喇嘛的敕谕（1860）",
        "folder": "11123咸丰皇帝为坐床事给十二世达赖喇嘛的敕谕",
        "file": "咸丰十年十月十一日敕谕_汉译文.txt",
        "body": XIANFENG,
        "url": "http://www.livingbuddha.cn/view-6231bba985014649b4ce2c81448bfa5f.html",
        "publisher": "文物出版社",
        "edition": "清咸丰十年(1860)敕谕；汉译文转引自西藏自治区档案馆编《西藏历史档案荟萃》第64号。原件满蒙藏文，西藏自治区档案馆藏。",
        "note": f"{STAMP} 核验原表网址，页面与条目相符，已归档汉译文。原件非汉文，汉译文出自1995年汇编。",
    },
    {
        "seq": 16504,
        "title": "刘玉墓志",
        "folder": "16504刘玉墓志",
        "file": "魏故咸阳太守刘府君墓志铭.txt",
        "body": LIU,
        "url": "http://www.ybsfzx.com/ldbt/beiwei010.htm",
        "publisher": "不详",
        "edition": "北魏孝昌三年；全称《魏故咸阳太守刘府君墓志铭》。母项《汉魏南北朝墓志汇编》。",
        "note": f"{STAMP} 核为北魏刘玉墓志，非明清同名人物。已归档释文。",
    },
]

CLUES = [
    {
        "seq": 5484,
        "url": "http://www.wenwuchina.com/a/15/209857.html",
        "publisher": "文物出版社",
        "edition": "2013年《潜山林新战国秦汉墓》第五章「东汉墓葬」",
        "note": f"{STAMP} 子项。母项叶润清主编《潜山林新战国秦汉墓》，文物出版社2013，ISBN 9787501038114。第五章第172–182页。现代考古报告，未取得可公开全文，仍记未找到。",
    },
    {
        "seq": 14071,
        "url": "https://www.shidianguji.com/mid-page/7568608946472222746",
        "publisher": "不详",
        "edition": "柳子文撰、王同老书；绍圣三年葬。母项《益都金石记》《八琼室金石补正续编》。",
        "note": f"{STAMP} 已定位母项，识典见跋文，完整志文未从可核页面抽出，仍记未找到。",
    },
    {
        "seq": 13431,
        "url": "https://www.gdwx.fudan.edu.cn/fd/ea/c4127a458218/page.htm",
        "publisher": "上海古籍出版社",
        "edition": "仇鹿鸣、夏婧《五代十国墓志汇编》开运〇一三",
        "note": f"{STAMP} 母项已定位为《五代十国墓志汇编》开运〇一三，公开全文未取得，仍记未找到。",
    },
    {
        "seq": 3301,
        "url": "https://www.shidianguji.com/zh/book/LGS10115355/chapter/1m2ad33tvdoha",
        "publisher": "不详",
        "edition": "专书见《朔方文库》；相关记载见《宁夏府志》卷十一兵防·宁夏满营。",
        "note": f"{STAMP} 专书《宁夏满营驻防事宜》未见公开全文。识典《宁夏府志》卷十一兵防有满营兵额俸饷，作母项线索，不替代原书。",
    },
]


def dump_found(it: dict) -> None:
    header = f"# {it['title']}\n# 来源: {it['url']}\n# 编码: UTF-8\n# 归档日期: {STAMP}\n\n"
    content = header + it["body"].strip() + "\n"
    folder = RESULT / it["folder"]
    product = PRODUCT / it["title"]
    folder.mkdir(parents=True, exist_ok=True)
    product.mkdir(parents=True, exist_ok=True)
    (folder / it["file"]).write_text(content, encoding="utf-8")
    (product / it["file"]).write_text(content, encoding="utf-8")
    (folder / "来源网址.txt").write_text(it["url"] + "\n", encoding="utf-8")
    (product / "来源网址.txt").write_text(it["url"] + "\n", encoding="utf-8")
    meta = {k: it[k] for k in ("seq", "title", "url", "publisher", "edition", "note")}
    meta["saved_at"] = STAMP
    meta["ok"] = True
    (folder / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (product / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("FOUND", it["seq"], it["title"])


def main() -> None:
    for it in FOUND:
        dump_found(it)
    by = {it["seq"]: ("是", it) for it in FOUND}
    by.update({c["seq"]: ("否", c) for c in CLUES})
    wb = load_workbook(EXCEL)
    ws = wb[wb.sheetnames[0]]
    n = 0
    for row in ws.iter_rows(min_row=5):
        if row[0].value != "韦广波":
            continue
        seq = row[1].value
        if seq not in by:
            continue
        flag, rec = by[seq]
        row[5].value = rec["url"]
        row[6].value = flag
        if rec.get("publisher"):
            row[7].value = rec["publisher"]
        if rec.get("edition"):
            row[8].value = rec["edition"]
        old = str(row[9].value or "").strip()
        row[9].value = rec["note"] if old in ("", "未找到", "无") else f"{old}；{rec['note']}"
        n += 1
        if flag == "否":
            folder = RESULT / f"{seq}{str(row[2].value or '').strip('《》')}"
            folder.mkdir(parents=True, exist_ok=True)
            (folder / "本轮补充").mkdir(exist_ok=True)
            (folder / "来源网址.txt").write_text(rec["url"] + "\n", encoding="utf-8")
            (folder / "本轮补充" / "线索.txt").write_text(rec["note"] + "\n", encoding="utf-8")
    wb.save(EXCEL)
    print("excel_rows", n)


if __name__ == "__main__":
    main()
