"""画面を通した通しの検査。

パート職員が受付入力をし、職員が日報を確定し、帳票・週の集計表・入力画面・入力もれチェックの
数がすべて同じ（T_受電 から数えたもの）になることを、期待値と突き合わせて確かめる。
現行 Excel で起きた 5 つの不具合が起きないことも、ここで確かめる。
"""
import datetime
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import (new_sim, get, post, follow, add_people, ids, strip_tags, Fail, check_clean, bridge)  # noqa: E402

DAY = "2026-08-25"
PEOPLE = [
    {"code": "001", "sei": "山田", "mei": "花子", "kind": "パート"},
    {"code": "002", "sei": "佐藤", "mei": "恵", "kind": "パート"},
    {"code": "003", "sei": "鈴木", "mei": "一郎", "kind": "職員"},
    {"code": "004", "sei": "田中", "mei": "美咲", "kind": "パート"},
]
# 区分ID_製品ID → 件数（区分と集計列は db/schema.sql の初期データ）
CALLS = {
    "山田": {"1_1": 5, "1_2": 3, "3_3": 2, "11_0": 2},                       # 申込 10・抽選 2
    "佐藤": {"4_1": 1, "8_2": 1, "6_1": 1, "14_0": 1, "15_0": 1},            # 払込 2（内 返金 1）・発送 1（内 交換 1）・その他 2
    "鈴木": {"2_4": 1, "12_0": 1, "17_0": 1},                               # 職員：申込 2・その他 1
}
EXPECT_COLS = {"申込": (12, 2), "抽選": (2, 0), "払込用紙": (2, 0), "商品発送": (1, 0), "その他": (3, 1)}
EXPECT_TOTAL = (20, 3)
TASKS = {"山田": {"1": 10, "3": 4}, "佐藤": {"1": 5, "13": "２"}}   # 全角の数字も受け付けること
CHECKS = []


def ok(cond, what):
    if not cond:
        raise Fail("期待と違います: " + what)
    CHECKS.append(what)


def ver_of(html):
    m = re.search(r'name="ver" value="([^"]*)"', html)
    return m.group(1) if m else ""


def entry(sim, folder, tid, cells, expect_redirect=True, ver=None):
    r = get(sim, f"/nippou/{folder}/entry.asp?d={DAY}&t={tid}")
    form = {"act": "save", "d": DAY, "t": str(tid), "ver": ver if ver is not None else ver_of(r.text)}
    for k, v in cells.items():
        form["c_" + k] = str(v)
    r = post(sim, f"/nippou/{folder}/entry.asp", form)
    if expect_redirect:
        if r.code != 302:
            raise Fail("受付入力の保存に失敗: " + strip_tags(r.text)[:1500])
        r = follow(sim, r)
    return r


def sheet_numbers(html):
    out = {}
    # 5 列：(職員) を数字の左に小さく、件数を右に
    for cid, sub, big in re.findall(r'data-col="(\d+)"><span class="st">\((\d+)\)</span><span class="v">(\d+)</span>', html):
        out[cid] = (int(big), int(sub))
    m = re.search(r'data-col="total"><span class="num">(\d+)</span> 件<span class="st st-r">\((\d+)\)</span>', html)
    if m:
        out["total"] = (int(m.group(1)), int(m.group(2)))
    m = re.search(r'class="num a-num">(\d+)</span> 名', html)
    out["att"] = int(m.group(1)) if m else None
    out["交換"] = int(re.search(r'data-k="交換">(\d+)<', html).group(1))
    out["返金"] = int(re.search(r'data-k="返金">(\d+)<', html).group(1))
    return out


def col_ids(sim):
    _, rows = sim.bridge.query("SELECT C.[集計列ID], C.[集計列名] FROM [M_集計列] AS C", [])
    return {name: str(cid) for cid, name in rows}


def run():
    sim = new_sim("flow")
    add_people(sim, PEOPLE)
    pid = ids(sim)
    ok(len(pid) == 4, "マスタ保守の画面から担当者を 4 人登録できる")

    # ---- パート職員の受付入力（part）と、職員の代理入力（staff） ----
    entry(sim, "part", pid["山田"], CALLS["山田"])
    entry(sim, "part", pid["佐藤"], CALLS["佐藤"])
    r = entry(sim, "staff", pid["鈴木"], CALLS["鈴木"])
    ok("保存しました" in r.text, "保存のあと「保存しました」が出る")

    # ---- その他業務 ----
    for name, items in TASKS.items():
        form = {"act": "save", "d": DAY, "t": str(pid[name])}
        for k, v in items.items():
            form["g_" + k] = str(v)
        r = post(sim, "/nippou/part/tasks.asp", form)
        ok(r.code == 302, f"その他業務を保存できる（{name}）")

    # ---- 日報（出勤者・回線数・記述欄）を保存して確定 ----
    form = {"act": "fix", "d": DAY, "lines": "5", "t1": "システム切替の初日。", "t2": "記念貨幣の在庫の問合せ 1 件（課長対応）",
            "t3": ""}
    for name in ("山田", "佐藤", "鈴木"):
        form[f"att_{pid[name]}"] = "1"
    form[f"h_{pid['山田']}"] = "9:00-13:00"
    r = post(sim, "/nippou/staff/daily.asp", form)
    if r.code != 302:
        raise Fail("日報の確定に失敗: " + strip_tags(r.text)[:1500])
    r = follow(sim, r)
    ok("確定しました" in r.text, "日報を確定できる")

    # ---- 帳票の数（期待値と突き合わせ） ----
    rep = get(sim, f"/nippou/staff/report.asp?d={DAY}")
    nums = sheet_numbers(rep.text)
    cid = col_ids(sim)
    ok(nums["total"] == EXPECT_TOTAL, f"帳票の合計が {EXPECT_TOTAL[0]}（うち職員 {EXPECT_TOTAL[1]}）: 実際 {nums['total']}")
    for name, exp in EXPECT_COLS.items():
        ok(nums[cid[name]] == exp, f"帳票の {name} が {exp[0]}（うち職員 {exp[1]}）: 実際 {nums[cid[name]]}")
    ok(nums["att"] == 3, f"帳票の出勤者数が 3: 実際 {nums['att']}")
    ok(nums["交換"] == 1 and nums["返金"] == 1, "帳票の 内 交換 1 件・内 返金 1 件")
    ok(re.search(r"令和\s*8\s*年\s*8\s*月\s*25\s*日（火）", strip_tags(rep.text)), "帳票の日付が和暦（令和　8年　　8月　　25日（火））")
    ok(re.search(r'（<span class="num">5</span> 回線）', rep.text), "帳票の出勤者の枠に（ 5 回線）が出る")
    ok(rep.text.count('<table class="s-stamp') == 2 and "藤本課長" in rep.text, "帳票の上に回覧の押印欄（左 4・右 4）が出る")
    ok("山田：9:00-13:00" in rep.text, "出勤者の備考に勤務時間が出る")
    ok(re.search(r'class="t-no">①</td><td class="t-nm">受注入力</td><td class="t-n"><span class="num">15</span> 件', rep.text), "その他業務 ① 受注入力 が 15 件（全員の合計）")
    ok(re.search(r'class="t-no">⑬</td><td class="t-nm">顧客整理</td><td class="t-n"><span class="num">2</span> 件', rep.text), "その他業務 ⑬ 顧客整理 が 2 件（全角の ２ を受け付けた）")

    # ---- 入力画面の合計（1 人ずつ）を足すと帳票と同じ ----
    tot = 0
    for name in CALLS:
        r = get(sim, f"/nippou/staff/entry.asp?d={DAY}&t={pid[name]}")
        n = int(re.search(r'data-col="total">(\d+)<', r.text).group(1))
        ok(n == sum(CALLS[name].values()), f"入力画面の {name} さんの合計が {sum(CALLS[name].values())}")
        tot += n
    ok(tot == EXPECT_TOTAL[0], "入力画面の合計を全員ぶん足すと帳票の合計と同じ")

    # ---- 日報の画面・週の集計表・入力もれチェックも同じ数 ----
    r = get(sim, f"/nippou/staff/daily.asp?d={DAY}")
    ok(re.search(r'data-col="total">20<', r.text), "日報の画面の合計が 20")
    r = get(sim, f"/nippou/staff/summary.asp?d={DAY}")
    ok("合いません" not in r.text, "週の集計表：日ごとの和と週全体の数が一致")
    m = re.search(r'8/25（火）</a></td><td class="n">(\d+) 人</td><td class="n">(\d+) <span class="note">\((\d+)\)', r.text)
    ok(m and (int(m.group(1)), int(m.group(2)), int(m.group(3))) == (3, 20, 3), "週の集計表の 8/25 が 出勤 3 人・合計 20（職員 3）")
    m = re.search(r'週の合計</td><td></td><td class="n">(\d+) <span class="note">\((\d+)\)', r.text)
    ok(m and (int(m.group(1)), int(m.group(2))) == (20, 3), "週の集計表の週の合計が 20（職員 3）")
    r = get(sim, f"/nippou/staff/check.asp?d={DAY}")
    ok("一致" in r.text and "不一致" not in r.text, "入力もれチェック：明細と日報の合計が一致")
    ok("気になる所はありません" in r.text, "入力もれチェック：もれ無し")

    # ---- 確定した日は直せない ----
    r = entry(sim, "part", pid["山田"], {"1_1": 99}, expect_redirect=False)
    ok("確定済み" in r.text, "確定した日は受付入力で保存できない")
    ok(sheet_numbers(get(sim, f"/nippou/staff/report.asp?d={DAY}").text)["total"] == EXPECT_TOTAL, "確定後に保存しようとしても帳票の数は変わらない")

    # ---- 不具合 2：入力が減った日に前の数が残らない（取り消してから減らす） ----
    r = post(sim, "/nippou/staff/daily.asp", {"act": "unfix", "d": DAY})
    ok(r.code == 302, "確定を取り消せる")
    entry(sim, "part", pid["山田"], {"1_1": 2, "1_2": 0})      # 5→2、3→0（行を消す）
    nums = sheet_numbers(get(sim, f"/nippou/staff/report.asp?d={DAY}").text)
    ok(nums["total"] == (14, 3), f"減らした数がそのまま帳票に出る（20→14）: 実際 {nums['total']}")
    ok(nums[cid["申込"]] == (6, 2), f"申込も減る（12→6）: 実際 {nums[cid['申込']]}")
    _, rows = sim.bridge.query("SELECT Count(J.[受電ID]) AS N FROM [T_受電] AS J WHERE J.[件数] = 0", [])
    ok(rows[0][0] == 0, "0 にした欄は行ごと消え、0 件の行が残らない")

    # ---- 同じ人を 2 つの画面で同時に直したら、後から押した方は止まる ----
    r1 = get(sim, f"/nippou/part/entry.asp?d={DAY}&t={pid['佐藤']}")
    old_ver = ver_of(r1.text)
    entry(sim, "part", pid["佐藤"], {"4_1": 3})
    r = entry(sim, "part", pid["佐藤"], {"4_1": 7}, expect_redirect=False, ver=old_ver)
    ok("別の画面で" in r.text, "別の画面で先に保存されていたら、上書きせずに知らせる")
    _, rows = sim.bridge.query("SELECT J.[件数] FROM [T_受電] AS J WHERE J.[区分ID] = 4", [])
    ok(rows == [[3]], "古い画面からの保存は反映されない")

    # ---- 数字でない入力は 1 つでもあれば何も保存しない ----
    r = entry(sim, "part", pid["佐藤"], {"4_1": "5", "14_0": "abc"}, expect_redirect=False)
    ok('class="num ng"' in r.text and "まだ保存していません" in r.text, "数字でない欄を赤くして、保存しない")
    _, rows = sim.bridge.query("SELECT J.[件数] FROM [T_受電] AS J WHERE J.[区分ID] = 4", [])
    ok(rows == [[3]], "誤りのあるときは、正しい欄も保存しない（全部か無しか）")
    entry(sim, "part", pid["佐藤"], {"4_1": "１"})   # 全角
    _, rows = sim.bridge.query("SELECT J.[件数] FROM [T_受電] AS J WHERE J.[区分ID] = 4", [])
    ok(rows == [[1]], "全角の数字（１）を 1 として受け付ける")

    # ---- 不具合 1：名前が変わっても、その人の数は落ちない（ID で持つ） ----
    form = {"id": str(pid["山田"]), "code": "009", "sei": "山田", "mei": "花子", "shimei": "山田　花子（旧姓）", "kana": "",
            "logon": "", "kind": "パート", "start": "2026-04-01", "end": "", "order": "10", "active": "1", "note": ""}
    r = post(sim, "/nippou/staff/master.asp?t=tanto", form)
    ok(r.code == 302, "担当者の氏名・コードを直せる")
    rep = get(sim, f"/nippou/staff/report.asp?d={DAY}")
    ok("山田　花子（旧姓）" in rep.text and sheet_numbers(rep.text)["total"] == (14, 3), "氏名・席番号を変えても、その人の数は帳票から落ちない")

    # ---- 不具合 3：日付が勝手に変わらない（翌日に開いても 8/25 のまま） ----
    import vbsim.builtins as B
    B.NOW_OVERRIDE[0] = datetime.datetime(2026, 8, 26, 9, 0, 0)
    rep = get(sim, f"/nippou/staff/report.asp?d={DAY}")
    ok(re.search(r"令和\s*8\s*年\s*8\s*月\s*25\s*日（火）", strip_tags(rep.text)) and sheet_numbers(rep.text)["total"] == (14, 3), "翌日に開いても、帳票の日付と数は 8/25 のまま")
    r = get(sim, "/nippou/part/entry.asp")
    ok('value="2026-08-26"' in r.text, "日付を指定しないで開くと今日（8/26）になる")
    B.NOW_OVERRIDE[0] = datetime.datetime(2026, 8, 25, 10, 0, 0)

    # ---- 担当者がいない日の出勤・在籍終了 ----
    form = {"id": str(pid["田中"]), "code": "004", "sei": "田中", "mei": "美咲", "shimei": "", "kana": "",
            "logon": "", "kind": "パート", "start": "2026-04-01", "end": "2026-08-24", "order": "40", "active": "1", "note": ""}
    post(sim, "/nippou/staff/master.asp?t=tanto", form)
    r = get(sim, f"/nippou/part/entry.asp?d={DAY}&t={pid['山田']}")
    ok("田中" not in r.text, "在籍終了日を過ぎた人は、入力画面の担当者の候補から消える")
    r = get(sim, "/nippou/part/entry.asp?d=2026-08-20&t=0")
    ok("田中" in r.text, "在籍していた日を開けば、候補に出る（過去は残る）")

    # ---- 記述欄が帳票の罫線に入りきらないときは保存しない ----
    long_form = {"act": "save", "d": DAY, "lines": "5", "t1": "\r\n".join(["あ"] * 5), "t2": "", "t3": ""}
    for name in ("山田", "佐藤", "鈴木"):
        long_form[f"att_{pid[name]}"] = "1"
    r = post(sim, "/nippou/staff/daily.asp", long_form)
    ok(r.code == 200 and "特記事項が 5 行になります" in r.text, "特記事項が帳票の罫線（4 行）に入りきらないときは、理由を出して保存しない")

    # ---- 入力もれ：出勤なのに入力なし ----
    form = {"act": "save", "d": DAY, "lines": "5", "t1": "", "t2": "", "t3": ""}
    for name in ("山田", "佐藤", "鈴木"):
        form[f"att_{pid[name]}"] = "1"
    post(sim, "/nippou/staff/daily.asp", form)
    entry(sim, "staff", pid["鈴木"], {"2_4": 0, "12_0": 0, "17_0": 0})
    r = get(sim, f"/nippou/staff/check.asp?d={DAY}")
    ok("出勤しているのに、入力がありません" in r.text, "入力もれチェック：出勤しているのに入力が無い人を赤で出す")

    # ---- その他 ----
    r = get(sim, "/nippou/part/daily.asp")
    ok(r.code == 404, "part には職員用の画面が無い（開けない）")
    r = get(sim, "/nippou/data/日報集計_be.accdb")
    ok(r.code == 404, "データベースのファイルは URL から取れない")
    r = get(sim, "/nippou/data/web.config")
    ok(r.code == 404, "data の web.config も URL から取れない")
    return sim


if __name__ == "__main__":
    try:
        run()
    except Fail as e:
        print("NG:", e)
        for c in CHECKS:
            print("  OK:", c)
        bridge().shutdown()
        sys.exit(1)
    for c in CHECKS:
        print("  OK:", c)
    print(f"OK: 通しの検査 {len(CHECKS)} 項目")
    bridge().shutdown()
