# -*- coding: utf-8 -*-
"""归档本轮已核验的墓志/别名书，并回写总览表。"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"D:\\AI_Project\\ancient_files_crawler")
RESULT = ROOT / "韦广波" / "爬取结果"
PRODUCT = ROOT / "韦广波" / "成果"
EXCEL = ROOT / "第22组总览表.xlsx"
STAMP = datetime.now().strftime("%Y-%m-%d")
sys.stdout.reconfigure(encoding="utf-8")

MANSHU = Path(r"C:\Users\64344\.cursor\projects\d-AI-Project-ruanjiansheji\agent-tools\3fdef153-330b-40d2-be2b-3cac38f6b3c8.txt").read_text(encoding="utf-8")

HOU = """周骠骑大将军开府侯莫陈道生墓志铭

君讳道生，字某，朔州武川人也。本系阴山，出自国族。降及于魏，在秦作刘。父少兴，武川镇将。山河抗拒，关塞被边。早擅威声，咸多雄烈。君子读书马上，清谈剑端。独运六奇，专精三略。虽复身居末将，而勇冠旌门；位在支军，而谋参幕府。

魏正光五年，任统军，颖天柱尔朱荣征北海王。永安三年，随太师贺拔胜入关，寻转别将。滑源卷甲，关城束马，并皆克捷，君有力焉。永熙三年，补都督。

太祖文皇帝奄有关河，令行天下，以君干略，委之爪牙。名洽中涓，功参上造。兵临河曲，前登白马之津；寇发蒲城，先战黄沙之苑。临晋横船，既禽赵将；马陵削树，复下齐兵。班瑞司勋，披图疏爵，授骠骑大将军、金紫光禄大夫、鄠县开国公，食邑五百。赵俨之为骠骑，止驾单车；张堪之拜光禄，长乘白马。以斯连类，朝野荣之。

大统九年，更姓侯莫陈氏，随大将军拓跋远经始阳。二水长斗，三川无市，擐甲搆病，死于辕门，春秋五十一。赠持节都督、潮州刺史。君在武川，文皇帝同乡里。霸功既立，王业克成，不忘舍讲，犹论偿博。今嗣德惟新，功臣追远。东都马鸣，不无见日之叹；北陵车过，终忆平生之言。有诏更赠使持节、骠骑大将军、开府仪同三司、都督宜敷绛三州诸军事、敷州刺史，谥某公，礼也。

夫人拓跋氏，安邑郡夫人。庭有钟鼓，家承箴诫，教容教德，言告以归。天和五年六月薨，即以其年十月同葬于京兆某县洪源乡。武子成寝，请西阶而合葬；平阳下嫁，即庐山而共坟。铭曰：

凝阴远寂，广漠平寒。沙穷瀚海，地尽皋兰。塞鸿秋去，胡桑夏乾。风土壮气，山河凛然。师旅上谷，威雄武川。君则继踵，代不乏贤。匣有忠剑，庭流孝泉。魏室多故，余风未殄。天保让德，当途废典。上将指踪，中涓力展。洛城夜捷，河梁朝剪。铁剑金龟，荣追玉铉。身胄汉祚，门承魏绪。并擅华宗，俱称当路。伉俪云匹，年龄并故。赵瑟秦声，同为丘墓。小陵石椁、洪原乡墓。刍灵两引，池柳双前。队路仍合，松城即连。霜随柳白，月逐坟圆。芝兰几代，陵谷何年。
"""

SHI = """大唐故徵君史府君墓誌銘并序
朝議郎試大理司直兼殿中侍御史渤海李再字安國撰

昔春秋時衛有史魚屍諫，人到于今稱之，此則史氏茂德，其流遠矣。□良夫與晉、魯敗齊師於鞌，銘鍾策勳，垂芳千古，則孫氏盛烈而□□生賢，此之謂不朽矣。於戲！史氏、孫氏皆生於衛。故延州萊季子稱□老□，而二族合歡，鳳皇于飛，挺生賢子也。宜哉賢子之美，此不書也。府君諱光，字重明，其先渤海郡人也。遠祖因宦於絳，子孫遂而家焉。今則爲絳人。曾祖澄，皇吏部常選；祖謹，高道不仕，吟嘯風月，當時不得而臣□。府君則高道之第三子也。懷肥遁之志，守父祖之風，累有辟命而堅辭不起，所謂巢許明時隱也，訓誘子弟，雍睦九族，怡怡自得，陶然守真，□十年間，樂天知命。嗚呼！道不常泰，天不憗遺，以貞元八年九月十九日寢疾終於私第，享年七十有六。邦國殄悴，風煙晦冥，愛子恒微，幼歸真宗，早晤玄理，受二百五十之具誠，總三千六萬之威儀，且弘業隨朝□古寺也。寳剎山立，仁祠洞開，佛事至大，待人弘闡，恒微常清，净辦事實。領都綱，又居上府，今即為佛寺之奧主也。三綱備□，一德日彰。雖鳩摩、道安，今古同流也。尊夫人樂安孫氏，斷機垂教，猶績自勤，真梁鴻之妻，實孟宗之慈母也。不幸以元和二年十二月遘疾終於私第，受年□□載。恒微泣血茹痛，銜哀柴毀，感慈親之永訣，瞻養堂以號慕。痛□□之崩陷，□泉扃而殞絕。以元和三年正月廿七日卜吉兆於良鄉縣（下泐）之原，禮也。夫誌者記也，懼陵谷變遷，誌曰：不□也，□□□，府君存，夫人温。常選碩德，高道蘭薰。□□□□，□里高□。令子淑女，同會祇園。□兆□石，□原爲靈。永（下泐）
"""

ZHAO = """隋故上開府樂安縣開國伯趙君墓誌。公諱惠，字世摸，金城郡人也。源流積而成遠，枝葉蘊而彌芳。門襲軒冕之榮，家傳鍾鼎之食，並詳於汗冊，備在緜圖。考敬，營州長史、昌黎建德二郡太守。奇姿秀挺，雅志淵沖，加以仁澤被於氓黎，拖留所以興詠；至德懷於民吏，來慕所以爲歌。公體秀山川，稟零嶽瀆，儀範當世，師表人倫。懷橘之年，夙標令望；佩觹之歲，早播英聲。齊天統五年，時年十四，合鄉舉爲建德郡功曹，至下空六字大隋開皇三年，爲刺史高寶寧擾動氓黎，阻兵稱亂，公遂被衆潰圍，擁入北狄。抱誠臣之節，豈松竹而喻其志；推奉國之操，非金石而方其心。至四年，將領兵衆一千餘人，俱來歸國，蒙授大都督。乃屬偽陳放命，肆彼姦回，虐用其民，招聚亡叛，作偽不義，綱紀無緒，民思禍亂，兵革恆聞。皇帝敏彼蒼生，離之塗炭，分命問罪，龔行天罰。公遂總鷹揚之衆，帥貔虎之師，矢石蹔交，已瞻亂轍，威略纔布，早見投衡，曾不密朝，一時清蕩。蒙授上開府、樂安縣開國伯，封邑七百戶。方欲長城飲馬，勒石燕山，誰謂與善徒言，奄從朝露，以開皇十年十月廿九日薨於陽州所。悲纏行路，痛結鄉邦，粵以開皇十一年十一月廿四日葬於雍州大興縣堺洪固鄉所，詔贈摸禮。但桑田易改，山谷遷移，德音鐫其玄石，冀千載而恆垂。乃爲銘曰：山川稟秀，辰象降零。家傳軒冕，門襲簪纓。誕惟君子，夙播令名。交遊致信，奉國推誠。浩浩波澜，滔滔江漢。宮牆深遠，莫測涯岸。萬傾同懷，千尋擢幹。春秋斡運，乾象遞旋。悠悠蒼昊，冥冥上玄。寧云與善，不辨才賢。一捐棺舍，永墜虞淵。遠日有期，親徒俱集。悲斯賵馬，賓僚灑泣。膴膴松野，蕭蕭寒隰。蘭菊恆芳，徽猷永立。
"""

ITEMS = [
    {
        "seq": 3815,
        "title": "云南记",
        "folder": "3815云南记",
        "file": "蠻書_四庫全書本_全覽.txt",
        "body": MANSHU,
        "url": "https://zh.wikisource.org/wiki/蠻書_(四庫全書本)/全覽",
        "publisher": "武英殿聚珍版/四库全书本",
        "edition": "唐樊绰《蛮书》十卷；袁滋《云南记》已佚。本条归档四库本《蛮书》全览作别名核验。",
        "note": f"{STAMP} 袁滋《云南记》原书已佚。已归档樊绰《蠻書（四庫全書本）》全览十卷；别名云南志/云南史记，部分文献亦称云南记。不能径视为袁滋原书。",
    },
    {
        "seq": 6636,
        "title": "史光及妻孙氏墓志",
        "folder": "6636史光及妻孙氏墓志",
        "file": "大唐故徵君史府君墓誌銘并序.txt",
        "body": SHI,
        "url": "https://www.shufawang.cn/xzmb/b2d7bea161936e3ea9cd25eb31d1e674.html",
        "publisher": "不详",
        "edition": "唐元和三年；录文转引自《隋唐五代墓誌匯編》北京卷第二冊 / 《唐代墓誌彙編續集》。",
        "note": f"{STAMP} 核为《大唐故徵君史府君墓誌銘并序》，墓主史光，妻乐安孙氏。已保存录文。母项：《唐代墓誌彙編續集》。",
    },
    {
        "seq": 8301,
        "title": "侯莫陈道生及妻拓跋氏墓志",
        "folder": "8301侯莫陈道生及妻拓跋氏墓志",
        "file": "周骠骑大将军开府侯莫陈道生墓志铭.txt",
        "body": HOU,
        "url": "https://www.shidianguji.com/book/SBCK123/chapter/1j73rp2e1fhod_406",
        "publisher": "四部丛刊",
        "edition": "庾信《庾子山集》所收志文；夫人拓跋氏附见。",
        "note": f"{STAMP} 母项为《庾子山集》。识典古籍四部丛刊本已提取志文及铭。原表已有该网址，本轮核验正文并归档。",
    },
    {
        "seq": 9069,
        "title": "隋故上开府乐安县开国伯赵君（惠）墓志",
        "folder": "9069隋故上开府乐安县开国伯赵君惠墓志",
        "file": "隋故上开府乐安县开国伯赵君墓誌.txt",
        "body": ZHAO,
        "url": "https://www.xiaoxiong360.com/shuhua/shiwen/3105.html",
        "publisher": "不详",
        "edition": "隋开皇十一年葬；墓主赵惠字世摸。录文亦见《隋代墓誌銘彙考》《長安新出墓誌》。",
        "note": f"{STAMP} 核为《隋赵惠墓志》/《赵世摸墓志》。已保存释文。母项线索：《隋代墓誌銘彙考》《長安新出墓誌》。",
    },
]


def dump_item(it: dict) -> None:
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
    meta = {
        "seq": it["seq"],
        "title": it["title"],
        "url": it["url"],
        "publisher": it["publisher"],
        "edition": it["edition"],
        "note": it["note"],
        "saved_at": STAMP,
        "ok": True,
    }
    (folder / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (product / "基本信息.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print("saved", it["seq"], it["title"], len(it["body"]))


def main() -> None:
    # attach 赵惠 if seq known later
    for it in ITEMS:
        dump_item(it)
    wb = load_workbook(EXCEL)
    ws = wb[wb.sheetnames[0]]
    by_seq = {it["seq"]: it for it in ITEMS}
    n = 0
    for row in ws.iter_rows(min_row=5):
        if row[0].value != "韦广波":
            continue
        seq = row[1].value
        if seq not in by_seq:
            continue
        it = by_seq[seq]
        row[5].value = it["url"]
        row[6].value = "是"
        row[7].value = it["publisher"]
        row[8].value = it["edition"]
        old = str(row[9].value or "").strip()
        row[9].value = it["note"] if old in ("", "未找到", "无") else f"{old}；{it['note']}"
        n += 1
    wb.save(EXCEL)
    print("excel_rows", n)


if __name__ == "__main__":
    main()
