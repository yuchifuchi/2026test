"""画面の見本（モックアップ）を作る。

本物の ASP を模擬の IIS で動かし、試しのデータを入れた状態の画面を 1 枚の HTML にまとめる。
画面の中身は ASP が出した HTML そのもの（手で写していない）なので、画面を直せば見本も同じに変わる。

  python3 tools/mockup.py
    → build/mockup/画面の見本.html          （納品の zip の 参考 に入れる。閉じた網でも見られるよう、外のファイルを読まない）
    → build/mockup/画面の見本_公開用.html    （Artifact として公開するもの。文字の形だけ Google Fonts から読む）
"""
import base64
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "tests"))
from harness import (new_sim, get, post, follow, add_people, ids, strip_tags, Fail, bridge, DIST, ROOT)  # noqa: E402
import test_docs  # noqa: E402

OUT = os.path.join(ROOT, "build", "mockup")
DAY = "2026-08-25"
PREV = "2026-08-24"

# ---- 試しのデータ（見本の画面に出す数。本物の数ではない） ----
PEOPLE = [
    {"code": "001", "sei": "山田", "mei": "花子", "kind": "パート"},
    {"code": "002", "sei": "佐藤", "mei": "恵", "kind": "パート"},
    {"code": "003", "sei": "高橋", "mei": "由美", "kind": "パート"},
    {"code": "004", "sei": "伊藤", "mei": "久美子", "kind": "パート"},
    {"code": "005", "sei": "渡辺", "mei": "さおり", "kind": "パート"},
    {"code": "006", "sei": "中村", "mei": "明美", "kind": "パート"},
    {"code": "101", "sei": "鈴木", "mei": "一郎", "kind": "職員"},
]
# 区分ID_製品ID → 件数（区分・製品は db/schema.sql の初期データ）
CALLS = {
    DAY: {
        "山田": {"1_1": 6, "1_3": 4, "3_3": 2, "5_1": 1, "11_0": 2, "15_0": 1},
        "佐藤": {"1_2": 3, "4_1": 2, "6_1": 1, "7_2": 2, "12_0": 3, "14_0": 1},
        "高橋": {"1_1": 4, "3_3": 3, "8_2": 1, "10_0": 2, "16_0": 1},
        "伊藤": {"1_3": 5, "2_3": 1, "7_1": 2, "13_0": 1, "17_0": 2},
        "渡辺": {"1_1": 3, "4_2": 1, "9_4": 2, "11_0": 1},
        "中村": {"3_3": 2, "5_1": 2, "12_0": 2, "14_0": 1},
        "鈴木": {"2_1": 1, "17_0": 1},
    },
    PREV: {
        "山田": {"1_1": 4, "1_3": 3, "11_0": 1},
        "佐藤": {"1_2": 2, "4_1": 1, "7_2": 1, "12_0": 2},
        "高橋": {"1_1": 3, "3_3": 1, "10_0": 2},
        "伊藤": {"1_3": 4, "7_1": 1, "17_0": 1},
        "渡辺": {"1_1": 2, "9_4": 1},
        "鈴木": {"2_1": 1},
    },
}
TASKS = {"山田": {"1": 12, "2": 5}, "佐藤": {"3": 8, "5": 4}, "高橋": {"7": 6, "8": 6},
         "伊藤": {"10": 20, "11": 20}, "渡辺": {"4": 3, "13": 2}, "中村": {"9": 15, "12": 2}}
MEMO = {
    "t1": "記念貨幣の申込開始日のため、午前中に申込方法の問合せが集中した。",
    "t2": "払込用紙の再発行の可否について、課長に引き継いだ（1 件）。",
    "t3": "申込サイトの入力例を増やしてほしいとの声あり。",
}

# ---- 画面ごとの説明（見本の左に出す） ----
ABOUT = {
    "part_menu": ("パート職員", "お気に入りに入れた「http://（サーバーの名前）/nippou/part/」を開く", [
        "大きなボタンは「受付入力」と「その他業務」の 2 つだけです。",
        "職員用の画面のファイルは、この part のフォルダに置いていません。アドレスを書き換えても開けません。",
    ]),
    "entry": ("パート職員（職員も、職員メニューから代わりに入れられます）", "メニューの「受付入力」", [
        "担当者は一覧から選びます。人は名前ではなく番号で覚えているので、名前の書き方が違って数が抜け落ちることがありません。",
        "区分の横の「→申込」などは、その件数が日報のどの列に入るかです（今の Excel では 2 行目に隠れていたものです）。",
        "保存は「その欄の今の件数」で上書きします。同じ数を 2 回足してしまうことがありません。",
        "日報が確定した日は、ここでは直せません（紙に出した数と食い違わないように）。",
    ]),
    "entry_ng": ("パート職員", "受付入力で、数字でない文字を入れて「保存する」を押したとき", [
        "数字でない欄が赤くなり、どこが違うかが上に出ます。",
        "1 つでも誤りがあれば、正しい欄も含めて何も保存しません（半分だけ保存されて数が合わなくなることがありません）。",
        "全角の数字（１２）は、そのまま数として受け付けます。",
    ]),
    "tasks": ("パート職員", "メニューの「その他業務」", [
        "電話以外の仕事 ①～⑬ の件数を入れます。名前は帳票の下の欄と同じです。",
        "名前は職員が「マスタ保守」の「業務項目」で直せます。",
    ]),
    "staff_menu": ("職員", "お気に入りに入れた「http://（サーバーの名前）/nippou/staff/」を開く", [
        "毎日の流れ（入力もれチェック → 日報の作成・確定 → 帳票）の順にボタンが並んでいます。",
        "パート職員の画面とは、アドレスのフォルダ（part と staff）で分けています。",
    ]),
    "daily": ("職員", "職員メニューの「日報の作成・確定」", [
        "出勤者に印を付け、回線数と 3 つの記述欄を入れて「保存して確定する」を押します。",
        "記述欄は、帳票の点線の数（特記事項 4 行・職員に代わった案件 5 行・要望 6 行）に入りきらないと保存しません。文字を縮めて押し込むことはしません。",
        "上の「この日の問合せ件数」は、帳票と同じ数です（どちらも受付入力の数から、開くたびに数え直しています）。",
    ]),
    "report": ("職員", "職員メニューの「帳票を見る・PDF にする」", [
        "課の印刷用シートと同じ様式です。罫線の位置は、様式の写真と 1.2mm 以内で合わせています。",
        "日付は選んだ日のままです。翌日に開いても、前の日の帳票は前の日の日付と数で出ます。",
        "画面で見るのも PDF にするのも、同じ 1 つの部品（sheet.asp）で作っています。",
    ]),
    "paper": ("職員", "帳票の画面の「PDF で出す（印刷はこちらから）」", [
        "サーバーの Edge で作った PDF を、そのまま画像にしたものです（この見本では、Edge の代わりに同じ仕組みの Chromium で作りました）。",
        "A4 縦 1 枚に収まります。記述欄や名前が上限いっぱいでも 1 枚に収まることを確かめています。",
    ]),
    "summary": ("職員", "職員メニューの「週の集計表」", [
        "月曜から金曜までの日ごとの数と、週の合計です。",
        "日ごとの数を足した値と、週全体を数えた値が合わないときは、赤い字で知らせます。",
    ]),
    "check": ("職員", "職員メニューの「入力もれチェック」", [
        "出勤しているのに入力が無い人、入力があるのに出勤に印が無い人を出します。",
        "1 人ずつの明細を足した数と、日報の合計が「一致」になっているかを見ます。",
    ]),
    "master": ("職員", "職員メニューの「マスタ保守」", [
        "担当者・区分・製品・業務項目・日報の列・回覧の欄の名前を直します。",
        "消す機能はありません。辞めた人は「在籍終了日」を入れます（過去の日報がそのまま作れるように）。",
        "同じ席番号の人が 2 人とも在籍していると、一覧で知らせます。",
    ]),
    "master_stamp": ("職員", "マスタ保守の「回覧の欄」", [
        "帳票の右上の押印欄 8 つ（左 4・右 4）の名前です。人が替わったら、ここで直します。",
        "名前は 6 文字までです（欄に収まる長さ）。",
    ]),
    "setup_check": ("システムの担当の方・設置する方", "「http://（サーバーの名前）/nippou/staff/setup_check.asp」を開く", [
        "設置のあと、うまく動かないときに開きます。上から順に、赤い所が直す所です。",
        "この画面は、ほかのファイルを読み込まない作りなので、ほかが壊れていても開けます。",
    ]),
    "error": ("だれでも", "画面の途中でエラーが起きたとき", [
        "IIS の英語の画面ではなく、この画面が出ます。",
        "「この 5 行をそのままお知らせください」の 5 行（ファイル名・行番号・内容・種別・URL）があれば、どこで何が起きたかが分かります。",
    ]),
}


def body_of(html):
    m = re.search(r"<body([^>]*)>(.*)</body>", html, re.S)
    if not m:
        raise Fail("画面に <body> がありません")
    cls = re.search(r'class="([^"]*)"', m.group(1))
    body = re.sub(r"<script\b.*?</script>", "", m.group(2), flags=re.S | re.I)
    return (cls.group(1) if cls else ""), body


def sig_of(file, query):
    """どの見本の画面に当たるかの鍵（mockup の JavaScript の sig() と同じ決まり）"""
    q = dict(p.split("=", 1) for p in query.split("&") if "=" in p)
    if file == "master.asp":
        s = q.get("t", "tanto")
        if q.get("id"):
            s += ":new" if q["id"] == "new" else ":edit"
        return s
    return ""


def capture(screens, sim, sid, group, title, url, nav=True, r=None, about=None):
    if r is None:
        r = get(sim, url)
    if r.code != 200:
        raise Fail(f"{url}: 画面が開けません（{r.code}）")
    cls, body = body_of(r.text)
    path, _, query = url.partition("?")
    parts = path.split("/")          # ['', 'nippou', 'part', 'entry.asp']
    folder, file = parts[2], (parts[3] or "default.asp")
    who, how, points = ABOUT[about or sid]
    screens.append({"id": sid, "group": group, "title": title, "nav": nav, "parent": sid if nav else (about or sid), "kind": "screen",
                    "folder": folder, "file": file, "sig": sig_of(file, query), "url": url,
                    "bodyClass": cls, "body": body, "who": who, "how": how, "points": points})


def entry(sim, tid, day, cells):
    r = get(sim, f"/nippou/part/entry.asp?d={day}&t={tid}")
    ver = re.search(r'name="ver" value="([^"]*)"', r.text).group(1)
    form = {"act": "save", "d": day, "t": str(tid), "ver": ver}
    form.update({"c_" + k: str(v) for k, v in cells.items()})
    return post(sim, "/nippou/part/entry.asp", form)


def daily_form(pid, day, act):
    """出勤者は、その日に受付入力のある人"""
    form = {"act": act, "d": day, "lines": "6", "t1": "", "t2": "", "t3": ""}
    for n in CALLS[day]:
        form[f"att_{pid[n]}"] = "1"
    return form


def build_screens():
    sim = new_sim("mockup")
    add_people(sim, PEOPLE)
    pid = ids(sim)
    # 前の日（8/24）は入力して確定まで済ませておく（週の集計表に 2 日ぶん出るように）
    for name, cells in CALLS[PREV].items():
        r = entry(sim, pid[name], PREV, cells)
        if r.code != 302:
            raise Fail("受付入力（前の日）に失敗: " + strip_tags(r.text)[:800])
    r = post(sim, "/nippou/staff/daily.asp", daily_form(pid, PREV, "fix"))
    if r.code != 302:
        raise Fail("前の日の確定に失敗: " + strip_tags(r.text)[:800])
    for name, cells in CALLS[DAY].items():
        r = entry(sim, pid[name], DAY, cells)
        if r.code != 302:
            raise Fail("受付入力に失敗: " + strip_tags(r.text)[:800])
    for name, items in TASKS.items():
        form = {"act": "save", "d": DAY, "t": str(pid[name])}
        form.update({"g_" + k: str(v) for k, v in items.items()})
        if post(sim, "/nippou/part/tasks.asp", form).code != 302:
            raise Fail("その他業務の保存に失敗")

    S = []
    P, ST, TR = "パート職員の画面", "職員の画面", "うまく動かないとき"
    capture(S, sim, "part_menu", P, "パート職員のメニュー", "/nippou/part/")
    # 佐藤さんの入力：保存の前・入れ間違い・保存のあと
    sato = pid["佐藤"]
    r = get(sim, f"/nippou/part/entry.asp?d={DAY}&t={sato}")
    capture(S, sim, "entry", P, "受付入力", f"/nippou/part/entry.asp?d={DAY}&t={sato}", r=r)
    bad = dict(CALLS[DAY]["佐藤"])
    bad["4_1"] = "２"
    bad["12_0"] = "3け"
    r = entry(sim, sato, DAY, bad)
    if "まだ保存していません" not in r.text:
        raise Fail("入れ間違いの画面が出ません")
    capture(S, sim, "entry_ng", P, "受付入力（入れ間違い）", "/nippou/part/entry.asp", r=r)
    S[-1]["sig"] = "ng"
    r = follow(sim, entry(sim, sato, DAY, CALLS[DAY]["佐藤"]))
    capture(S, sim, "entry_saved", P, "受付入力（保存のあと）", f"/nippou/part/entry.asp?d={DAY}&t={sato}", nav=False, r=r, about="entry")
    S[-1]["sig"] = "saved"
    capture(S, sim, "tasks", P, "その他業務", f"/nippou/part/tasks.asp?d={DAY}&t={pid['山田']}")
    form = {"act": "save", "d": DAY, "t": str(pid["山田"])}
    form.update({"g_" + k: str(v) for k, v in TASKS["山田"].items()})
    r = follow(sim, post(sim, "/nippou/part/tasks.asp", form))
    capture(S, sim, "tasks_saved", P, "その他業務（保存のあと）", f"/nippou/part/tasks.asp?d={DAY}&t={pid['山田']}", nav=False, r=r, about="tasks")
    S[-1]["sig"] = "saved"

    capture(S, sim, "staff_menu", ST, "職員メニュー", "/nippou/staff/staff.asp")
    capture(S, sim, "check", ST, "入力もれチェック", f"/nippou/staff/check.asp?d={DAY}")
    # 日報：確定の前（出勤に印を付けた状態で保存）→ 確定のあと
    form = daily_form(pid, DAY, "save")
    form.update(MEMO)
    form[f"h_{pid['山田']}"] = "9:00-13:00"
    form[f"n_{pid['中村']}"] = "午後のみ"
    r = post(sim, "/nippou/staff/daily.asp", form)
    if r.code != 302:
        raise Fail("日報の保存に失敗: " + strip_tags(r.text)[:800])
    capture(S, sim, "daily", ST, "日報の作成・確定", f"/nippou/staff/daily.asp?d={DAY}")
    form["act"] = "fix"
    r = follow(sim, post(sim, "/nippou/staff/daily.asp", form))
    if "確定しました" not in r.text:
        raise Fail("日報の確定に失敗: " + strip_tags(r.text)[:800])
    capture(S, sim, "daily_fixed", ST, "日報の作成・確定（確定のあと）", f"/nippou/staff/daily.asp?d={DAY}", nav=False, r=r, about="daily")
    S[-1]["sig"] = "fixed"
    capture(S, sim, "report", ST, "帳票（画面で見る）", f"/nippou/staff/report.asp?d={DAY}")
    S.append(paper(sim))
    capture(S, sim, "summary", ST, "週の集計表", f"/nippou/staff/summary.asp?d={DAY}")
    capture(S, sim, "master", ST, "マスタ保守（担当者）", "/nippou/staff/master.asp?t=tanto")
    for t, nm in (("kubun", "区分"), ("seihin", "製品"), ("block", "ブロック"), ("gyomu", "業務項目"), ("col", "日報の列")):
        capture(S, sim, "master_" + t, ST, f"マスタ保守（{nm}）", f"/nippou/staff/master.asp?t={t}", nav=False, about="master")
    capture(S, sim, "master_edit", ST, "マスタ保守（担当者を直す）", f"/nippou/staff/master.asp?t=tanto&id={pid['山田']}", nav=False, about="master")
    capture(S, sim, "master_new", ST, "マスタ保守（担当者を足す）", "/nippou/staff/master.asp?t=tanto&id=new", nav=False, about="master")
    capture(S, sim, "master_stamp", ST, "マスタ保守（回覧の欄）", "/nippou/staff/master.asp?t=stamp")
    capture(S, sim, "master_stamp_edit", ST, "マスタ保守（回覧の欄を直す）", "/nippou/staff/master.asp?t=stamp&id=5", nav=False, about="master_stamp")

    capture(S, sim, "setup_check", TR, "設置チェック", "/nippou/staff/setup_check.asp")
    # エラーの画面：わざと配列の範囲を超える 1 枚を置いて開く（test_troubleshoot.py と同じやり方）
    p = os.path.join(sim.webroot, "nippou", "part", "sample_error.asp")
    open(p, "w", encoding="utf-8", newline="\n").write(
        '<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%\nOption Explicit\nResponse.CharSet = "utf-8" : Response.CodePage = 65001\nDim a(2)\na(5) = 1\n%>')
    r = sim.request("GET", f"/nippou/part/sample_error.asp?d={DAY}")
    if "NIPPOU-ERROR" not in r.text:
        raise Fail("エラーの画面が出ません")
    capture(S, sim, "error", TR, "エラーの画面", f"/nippou/part/sample_error.asp?d={DAY}", r=r)
    S[-1]["sig"] = "error"
    return S


def paper(sim):
    r = sim.request("GET", f"/nippou/staff/printpdf.asp?d={DAY}")
    if not r.body.startswith(b"%PDF"):
        raise Fail("PDF ができません: " + strip_tags(r.text)[:800])
    pdf = os.path.join(OUT, "帳票.pdf")
    open(pdf, "wb").write(r.body)
    subprocess.run(["pdftoppm", "-r", "110", "-png", "-singlefile", pdf, os.path.join(OUT, "帳票")], check=True)
    png = open(os.path.join(OUT, "帳票.png"), "rb").read()
    who, how, points = ABOUT["paper"]
    return {"id": "paper", "group": "職員の画面", "title": "印刷される紙（PDF）", "nav": True, "parent": "paper", "kind": "paper",
            "folder": "staff", "file": "printpdf.asp", "sig": "", "url": f"/nippou/staff/printpdf.asp?d={DAY}",
            "img": "data:image/png;base64," + base64.b64encode(png).decode("ascii"),
            "who": who, "how": how, "points": points}


def app_css():
    """画面の css を、見本の中の「窓」だけに効くように直す（html・body の規則を窓の外枠に移す）"""
    css = open(os.path.join(DIST, "staff", "css", "style.css"), encoding="utf-8").read()
    css = css.replace("@charset \"utf-8\";", "")
    css = css.replace("html, body {", ".m-root {")
    css = re.sub(r"(^|[\s,}])body(?=[\s.{,:])", r"\1.m-root", css)
    return css


def page(screens, public):
    tpl = open(os.path.join(HERE, "mockup_template.html"), encoding="utf-8").read()
    data = json.dumps({"day": DAY, "screens": screens, "css": app_css()}, ensure_ascii=False)
    data = data.replace("</", "<\\/")
    fonts = ('<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=BIZ+UDPGothic:wght@400;700'
             '&family=BIZ+UDGothic&family=BIZ+UDPMincho:wght@400;700&display=swap">') if public else ""
    html = tpl.replace("{{FONTS}}", fonts).replace("{{DATA}}", data)
    if not public:
        html = ('<!DOCTYPE html>\n<html lang="ja">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1">\n') + \
            html.replace("<!--HEAD-END-->", "</head>\n<body>") + "\n</body>\n</html>\n"
    else:
        html = html.replace("<!--HEAD-END-->", "")
    return html


def check_offline(html):
    """閉じた網の中で開くので、外のファイルを 1 つも読まないこと"""
    probs = []
    for m in re.finditer(r"""(?:src|href)\s*=\s*["']?(https?:|//)""", html):
        probs.append("外のファイルを読む所があります: " + html[m.start():m.start() + 80])
    for m in re.finditer(r"url\(\s*['\"]?(https?:|//)", html):
        probs.append("外のファイルを読む所があります: " + html[m.start():m.start() + 80])
    if "@import" in html:
        probs.append("@import があります")
    return probs


def check_words(screens, tpl):
    """見本の説明の文にも、手順書と同じ「使わない言葉」の決まりを当てる"""
    text = re.sub(r"<[^>]+>", " ", re.sub(r"<script\b.*?</script>|<style\b.*?</style>", " ", tpl, flags=re.S))
    for s in screens:
        text += " " + s["title"] + " " + s["who"] + " " + s["how"] + " " + " ".join(s["points"])
    return [f"見本の説明に使わない言葉「{w}」があります" for w in test_docs.BANNED if w in text]


def main():
    if os.path.exists(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    screens = build_screens()
    bridge().shutdown()
    probs = check_words(screens, open(os.path.join(HERE, "mockup_template.html"), encoding="utf-8").read())
    offline = page(screens, public=False)
    probs += check_offline(offline)
    ids_ = [s["id"] for s in screens]
    if len(ids_) != len(set(ids_)):
        probs.append("画面の id が重なっています")
    if probs:
        for p in probs:
            print("NG:", p)
        sys.exit(1)
    open(os.path.join(OUT, "画面の見本.html"), "w", encoding="utf-8", newline="\n").write(offline)
    open(os.path.join(OUT, "画面の見本_公開用.html"), "w", encoding="utf-8", newline="\n").write(page(screens, public=True))
    n = sum(1 for s in screens if s["nav"])
    kb = len(offline.encode("utf-8")) // 1024
    print(f"OK: 画面の見本（一覧に {n} 枚・つながる画面を合わせて {len(screens)} 枚、{kb} KB、外のファイルを読まない）")


if __name__ == "__main__":
    try:
        main()
    except Fail as e:
        print("NG:", e)
        sys.exit(1)
