"""現地で動かないときの道具（setup_check.asp / probe1～7 / error.asp / 致命エラー画面）の検査。

わざと壊した状態を作り、
  ・設置チェックが「どこが悪いか」を赤い行で出すこと
  ・probe の「最初に赤くなる番号」で層を言い当てられること
  ・エラー画面に ファイル名／行番号／内容／種別／URL が出ること
を確かめる。
"""
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import new_sim, get, post, strip_tags, Fail, check_clean, bridge  # noqa: E402

CHECKS = []
DB = "日報集計_be.accdb"


def ok(cond, what):
    if not cond:
        raise Fail("期待と違います: " + what)
    CHECKS.append(what)


def rows(html):
    """設置チェックの表を (項目, 結果, 内容) の一覧にする。"""
    out = []
    for cls, sec, item, st, text in re.findall(r'<tr class="(\w+)"><td>(.*?)</td><td>(.*?)</td><td class="st">(.*?)</td><td class="pre">(.*?)</td></tr>', html, re.S):
        out.append((cls, sec, item, text))
    return out


def ng_items(html):
    return [(r[2], r[3]) for r in rows(html) if r[0] == "ng"]


def setup_check(sim, folder):
    r = sim.request("GET", f"/nippou/{folder}/setup_check.asp")
    if r.strict or r.warnings or r.error:
        raise Fail(f"setup_check.asp 自体が失敗: {r.error} {r.strict} {r.warnings}\n{strip_tags(r.text)[:1000]}")
    return r


def wfile(sim, rel, data):
    p = os.path.join(sim.webroot, "nippou", *rel.split("/"))
    open(p, "wb").write(data if isinstance(data, bytes) else data.encode("utf-8"))


def rfile(sim, rel):
    return open(os.path.join(sim.webroot, "nippou", *rel.split("/")), "rb").read()


def set_config(sim, folder, name, value):
    c = rfile(sim, f"{folder}/include/config.asp").decode("utf-8")
    pat = re.compile(rf'Const {name} = "[^"]*"')
    if not pat.search(c):
        raise Fail(f"config.asp の {name} が見つかりません")
    c2 = pat.sub(lambda m: f'Const {name} = "{value}"', c)
    wfile(sim, f"{folder}/include/config.asp", c2)


def first_red_probe(sim, folder):
    r = setup_check(sim, folder)
    for item, text in ng_items(r.text):
        m = re.match(r"probe(\d)\.asp", item)
        if m:
            return int(m.group(1)), text
    return None, ""


def run():
    # ---- 1. ふつうに置いたとき：赤い行が無い ----
    sim = new_sim("ts_ok")
    for folder in ("part", "staff"):
        r = setup_check(sim, folder)
        ok(ng_items(r.text) == [], f"正しく置いたとき、設置チェック（{folder}）に赤い行が無い: {ng_items(r.text)}")
    r = setup_check(sim, "staff")
    ok(any(i[2] == "PDF の試し作成" and i[0] == "ok" for i in rows(r.text)), "設置チェック（staff）から PDF を試しに作れる")
    ok(any(i[2] == "ファイルの形式" and i[0] == "ok" for i in rows(r.text)), "設置チェックが .accdb の形式（0x02）を確かめる")
    ok(any(i[2] == "probe7.asp" and i[0] == "ok" for i in rows(r.text)), "設置チェックが probe1～7 をサーバー自身から開く")
    ok(any(i[2] == "M_担当者" and i[0] == "warn" for i in rows(r.text)), "担当者が 0 人なら黄色で知らせる（出荷時は空）")

    # ---- 2. DB_PATH の 3 通りの書き方 ----
    sim = new_sim("ts_paths")
    src = os.path.join(sim.webroot, "nippou", "data", DB)
    d_drive = sim.winfs.to_local("D:\\data")
    os.makedirs(d_drive)
    shutil.copyfile(src, os.path.join(d_drive, DB))
    unc = sim.winfs.to_local("\\\\NIPPOU-SV\\share\\日報")
    os.makedirs(unc)
    shutil.copyfile(src, os.path.join(unc, DB))
    cases = [("/nippou/data/" + DB, "C:\\inetpub\\wwwroot\\nippou\\data\\" + DB),
             (DB, "C:\\inetpub\\wwwroot\\nippou\\data\\" + DB),
             ("data/" + DB, "C:\\inetpub\\wwwroot\\nippou\\data\\" + DB),
             ("D:\\data\\" + DB, "D:\\data\\" + DB),
             ("\\\\NIPPOU-SV\\share\\日報\\" + DB, "\\\\NIPPOU-SV\\share\\日報\\" + DB)]
    for raw, want in cases:
        set_config(sim, "part", "DB_PATH", raw)
        p4 = get(sim, "/nippou/part/probe4.asp")
        ok(want in p4.text, f"DB_PATH「{raw}」を db.asp が {want} と解く")
        sc = setup_check(sim, "part")
        place = [i for i in rows(sc.text) if i[2] == "場所"]
        ok(place and place[0][3] == want, f"設置チェックも同じ場所に解く（{raw}）")
        ok("PROBE5-OK" in get(sim, "/nippou/part/probe5.asp").text, f"その場所で接続できる（{raw}）")
    set_config(sim, "part", "DB_PATH", "../data/" + DB)
    r = sim.request("GET", "/nippou/part/probe4.asp")
    ok("NIPPOU-FATAL" in r.text and "..」は使えません" in r.text and r.code == 200, "DB_PATH に .. を書いたら、理由を画面に出す（HTTP 200）")

    # ---- 3. ACE が無い（または 32/64 ビットが合わない） ----
    sim = new_sim("ts_ace")
    sim.providers = set()
    n, text = first_red_probe(sim, "part")
    ok(n == 5, f"ACE が無いとき、最初に赤くなるのは probe5（実際 {n}）")
    sc = setup_check(sim, "part")
    ok(any(i == "ACE で接続" and "32 ビット版と 64 ビット版" in t for i, t in ng_items(sc.text)), "設置チェックが ACE の問題だと言い当てる")
    r = sim.request("GET", "/nippou/part/entry.asp")
    ok(r.code == 200 and "NIPPOU-FATAL" in r.text and "Microsoft Access データベース エンジン" in r.text, "入力画面にも ACE の問題だと出る（HTTP 200）")
    ok(all(k in r.text for k in ("ファイル名", "行番号", "内容", "種別", "URL")), "致命エラーの画面に 5 行（ファイル名・行番号・内容・種別・URL）が出る")

    # ---- 4. データベースが無い ----
    sim = new_sim("ts_nodb")
    os.remove(os.path.join(sim.webroot, "nippou", "data", DB))
    n, _ = first_red_probe(sim, "part")
    ok(n == 5, "データベースが無いとき、最初に赤くなるのは probe5")
    sc = setup_check(sim, "part")
    ok(any(i == "ファイル" for i, _ in ng_items(sc.text)), "設置チェックが「ファイルがありません」と出す")
    r = sim.request("GET", "/nippou/part/entry.asp")
    ok("config.asp の DB_PATH" in r.text, "入力画面に「DB_PATH を確かめて」と出る")

    # ---- 5. データベースのフォルダに書き込めない ----
    sim = new_sim("ts_ro")
    sim.readonly_dirs = [os.path.join(sim.webroot, "nippou", "data")]
    sc = setup_check(sim, "part")
    ok(any(i == "フォルダへの書き込み" and "IUSR" in t for i, t in ng_items(sc.text)), "書き込めないとき、設置チェックが IUSR / IIS_IUSRS の許可を案内する")
    r = sim.request("POST", "/nippou/staff/master.asp?t=tanto", form={"id": "new", "code": "001", "sei": "山田", "mei": "", "shimei": "",
                                                                        "kana": "", "logon": "", "kind": "パート", "start": "", "end": "",
                                                                        "order": "", "active": "1", "note": ""})
    ok("NIPPOU-FATAL" in r.text and "書き込みの許可" in r.text, "保存しようとしたら、書き込みの許可の問題だと画面に出る")

    # ---- 6. part と staff の config.asp を取り違えた ----
    sim = new_sim("ts_swap")
    staff_cfg = rfile(sim, "staff/include/config.asp")
    part_cfg = rfile(sim, "part/include/config.asp")
    wfile(sim, "part/include/config.asp", staff_cfg)
    wfile(sim, "staff/include/config.asp", part_cfg)
    sc = setup_check(sim, "part")
    ok(any(i == "いまの役割" and "取り違え" in t for i, t in ng_items(sc.text)), "config.asp を取り違えたら、設置チェックが「取り違え」と出す（part）")
    r = sim.request("GET", "/nippou/staff/daily.asp")
    ok("この画面は職員用です" in r.text, "取り違えた staff では、職員用の画面が止まる（入力画面へ案内）")

    # ---- 7. web.config を取り違えた／書けない設定を入れた ----
    sim = new_sim("ts_wc")
    wfile(sim, "part/web.config", rfile(sim, "staff/web.config"))
    sc = setup_check(sim, "part")
    ok(any(i[2] == "web.config" and i[0] == "warn" for i in rows(sc.text)), "web.config を取り違えたら黄色で知らせる")
    bad = rfile(sim, "staff/web.config").decode("utf-8").replace("<system.webServer>", '<system.webServer>\r\n    <asp scriptErrorSentToBrowser="true" />')
    wfile(sim, "staff/web.config", bad)
    r = sim.request("GET", "/nippou/staff/staff.asp")
    ok(r.code == 500 and "500.19" in r.text, "web.config に <asp> を書くと、そのフォルダ全体が 500.19 になる（だから書かない）")

    # ---- 8. include に BOM ----
    sim = new_sim("ts_bom")
    wfile(sim, "part/include/layout.asp", b"\xef\xbb\xbf" + rfile(sim, "part/include/layout.asp"))
    sc = setup_check(sim, "part")
    ok(any(i == "include\\layout.asp" and "BOM" in t for i, t in ng_items(sc.text)), "include に BOM があると、設置チェックがファイル名つきで知らせる")
    n, _ = first_red_probe(sim, "part")
    ok(n == 7, f"layout.asp が壊れているとき、最初に赤くなるのは probe7（実際 {n}）")

    # ---- 9. 実行時エラー → error.asp に 5 行 ----
    sim = new_sim("ts_err")
    wfile(sim, "part/broken.asp", '<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%\nOption Explicit\nResponse.CharSet = "utf-8" : Response.CodePage = 65001\nDim a(2)\na(5) = 1\n%>')
    r = sim.request("GET", "/nippou/part/broken.asp?d=2026-08-25")
    ok(r.handled_by == "/nippou/part/error.asp", "ASP の実行時エラー（500.100）で error.asp が動く")
    ok(r.code == 200 and "NIPPOU-ERROR" in r.text, "エラーの説明画面は HTTP 200 で返る（IIS に差し替えられない）")
    t = strip_tags(r.text)
    ok("/nippou/part/broken.asp" in t and "行番号 5" in t and "インデックスが有効範囲にありません" in t and "800A0009" in t and "?d=2026-08-25" in t,
       "error.asp にファイル名・行番号・内容・種別（番号）・URL が出る")
    ok("5 行をそのまま" in t, "「この 5 行をそのままお知らせください」と出る")

    # ---- 10. Edge が無い ----
    sim = new_sim("ts_edge", edge_present=False)
    sc = setup_check(sim, "staff")
    ok(any(i == "Microsoft Edge" for i, _ in ng_items(sc.text)), "Edge が無いとき、設置チェックが赤で知らせる")
    r = sim.request("GET", "/nippou/staff/printpdf.asp?d=2026-08-25")
    check_clean(r)
    ok("PDF を作れませんでした" in r.text and "Ctrl キーを押しながら P" in r.text, "PDF を作れないときは理由と、画面から印刷する方法を出す")

    # ---- 11. .accdb が Access 2016 専用形式（0x05） ----
    sim = new_sim("ts_fmt")
    p = os.path.join(sim.webroot, "nippou", "data", DB)
    b = bytearray(open(p, "rb").read())
    b[0x14] = 0x05
    open(p, "wb").write(bytes(b))
    sc = setup_check(sim, "part")
    ok(any(i == "ファイルの形式" and "2016 専用形式" in t for i, t in ng_items(sc.text)), "形式が 0x05 なら、設置チェックが「課の Access で開けません」と出す")
    # ---- 12. ROLE が空のときは、ログオン名で職員かを決める（将来 Windows 認証にしたとき用の道） ----
    sim = new_sim("ts_role")
    from harness import add_people
    add_people(sim, [{"code": "003", "sei": "鈴木", "kind": "職員"}])
    sim.bridge.execute("UPDATE [M_担当者] SET [ログオン名] = ? WHERE [姓] = ?", [{"t": "str", "v": "suzuki"}, {"t": "str", "v": "鈴木"}])
    set_config(sim, "staff", "ROLE", "")
    r = sim.request("GET", "/nippou/staff/daily.asp", sv={"LOGON_USER": "MINT\\suzuki"})
    check_clean(r)
    ok("日報の作成・確定" in r.text and "この画面は職員用です" not in r.text, "ROLE が空なら、ログオン名が M_担当者 の職員に当たる人は職員用の画面を開ける")
    r = sim.request("GET", "/nippou/staff/daily.asp", sv={"LOGON_USER": ""})
    ok("この画面は職員用です" in r.text, "ROLE が空で、ログオン名が無い（匿名）なら職員用の画面は止まる")

    # ---- 13. ACE の作業用の一時フォルダに書けない ----
    sim = new_sim("ts_temp")
    sim.readonly_dirs = [sim.winfs.to_local("C:\\Windows\\Temp")]
    sc = setup_check(sim, "part")
    ok(any(i[2] == "一時フォルダ（ACE の作業用）" and i[0] == "warn" for i in rows(sc.text)), "ACE の一時フォルダに書けないとき、黄色で知らせる")
    return True


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
    print(f"OK: 困ったときの道具の検査 {len(CHECKS)} 項目")
    bridge().shutdown()
