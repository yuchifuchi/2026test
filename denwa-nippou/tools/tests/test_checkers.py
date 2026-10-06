"""検査の道具そのものの検査。わざと誤りを入れた小さなファイルを作り、検査が必ず見つけることを確かめる。

（検査が 1 つも見つけないのは「誤りが無い」のか「検査が働いていない」のか区別できないため）
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import ROOT, Fail  # noqa: E402
from vbsim.lint import lint_page  # noqa: E402
from vbsim import sqlcheck, csscheck  # noqa: E402
from vbsim.sim import AspSim  # noqa: E402

WORK = os.path.join(ROOT, "build", "checkers")
HEAD = '<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%\nOption Explicit\nResponse.CharSet = "utf-8" : Response.CodePage = 65001\n'
CHECKS = []


def page(name, body, head=HEAD, files=None):
    d = os.path.join(WORK, "site", "part")
    os.makedirs(os.path.join(d, "include"), exist_ok=True)
    for fn, txt in (files or {}).items():
        open(os.path.join(d, fn), "wb").write(txt if isinstance(txt, bytes) else txt.encode("utf-8"))
    p = os.path.join(d, name)
    data = head + body
    open(p, "wb").write(data if isinstance(data, bytes) else data.encode("utf-8"))
    return [str(f) for f in lint_page(p, os.path.join(WORK, "site"))]


def expect(what, found, word):
    if not any(word in f for f in found):
        raise Fail(f"検査が見逃しました: {what}（見つかったもの: {found}）")
    CHECKS.append(what)


def clean(what, found):
    if found:
        raise Fail(f"誤りの無いものを誤りとしました: {what}: {found}")
    CHECKS.append(what)


def run():
    if os.path.exists(WORK):
        shutil.rmtree(WORK)
    clean("正しいページには何も言わない", page("ok.asp", 'Dim a(4), i\nFor i = 0 To 4\n  a(i) = i\nNext\n%>'))
    # ---- 7-2 Response.End と On Error Resume Next ----
    expect("On Error Resume Next の中の Response.End", page("t1.asp", 'On Error Resume Next\nResponse.End\nOn Error GoTo 0\n%>'), "Response.End 系")
    expect("On Error Resume Next の中で、Response.End を呼ぶ自作の Sub を呼ぶ",
           page("t2.asp", 'Sub Stop1()\n  Response.Redirect "a.asp"\nEnd Sub\nOn Error Resume Next\nStop1\nOn Error GoTo 0\n%>'), "Response.End 系")
    expect("On Error GoTo 0 に戻さずに Response.End（致命エラー表示の誤り）",
           page("t3.asp", 'Sub Fatal1()\n  On Error Resume Next\n  Response.Clear\n  Response.Write "x"\n  Response.End\n  On Error GoTo 0\nEnd Sub\nFatal1\n%>'), "Response.End 系")
    clean("On Error GoTo 0 に戻してからの Response.End は正しい",
          page("t4.asp", 'Sub Fatal2()\n  On Error Resume Next\n  Response.Clear\n  On Error GoTo 0\n  Response.End\nEnd Sub\nFatal2\n%>'))
    expect("On Error Resume Next を戻していない", page("t5.asp", 'Dim x\nOn Error Resume Next\nx = 1\n%>'), "戻していません")
    # ---- 7-6 配列 ----
    expect("Dim a(4) を For i = 0 To 7 で回す", page("t6.asp", 'Dim a(4), i\nFor i = 0 To 7\n  a(i) = i\nNext\n%>'), "添字が範囲を超えます")
    expect("Const の上限で回して添字あふれ", page("t7.asp", 'Const N = 8\nDim a(4), i\nFor i = 0 To N - 1\n  a(i) = i\nNext\n%>'), "添字が範囲を超えます")
    # ---- 名前・構文 ----
    expect("定義されていない関数", page("t8.asp", 'Dim x\nx = NoSuchFunc(1)\n%>'), "定義されていない名前")
    expect("定義されていない変数（Option Explicit）", page("t9.asp", 'y = 1\n%>'), "定義されていない名前")
    expect("引数の数が違う", page("t10.asp", 'Function F(a, b)\n  F = a\nEnd Function\nDim x\nx = F(1)\n%>'), "引数の数")
    expect("End If が無い", page("t11.asp", "Dim x\nIf x Then\n  x = 1\n%>"), "End If")
    expect("Next が無い", page("t12.asp", "Dim i\nFor i = 0 To 1\n%>"), "Next")
    expect("Sub をかっこ付き・引数 2 つで呼ぶ", page("t13.asp", 'Sub S(a, b)\nEnd Sub\nS(1, 2)\n%>'), "かっこは使えません")
    expect("関数と変数の名前が大文字小文字違いで同じ", page("t14.asp", 'Dim fv\nFunction FV(n)\n  FV = n\nEnd Function\n%>'), "再定義")
    expect("Const に式", page("t15.asp", 'Const A = "x" & "y"\n%>'), "Const には定数しか")
    expect("Session を使う", page("t16.asp", 'Dim x\nx = Session("a")\n%>'), "Session")
    # ---- <% %> と include ----
    expect("<% が閉じていない", page("t17.asp", 'Dim x\n<p>'), "閉じていません")
    expect("include 先が無い", page("t18.asp", '%><!--#include file="include/none.asp"--><%\n%>'), "取り込み先がありません")
    expect("include に .. を使う", page("t19.asp", '%><!--#include file="../x.asp"--><%\n%>'), "..")
    expect("ページの先頭の決まりを守っていない", page("t20.asp", 'Dim x\n%>', head='<%@ LANGUAGE="VBScript" %>\n<%\n'), "1 行目")
    expect("ページに BOM", page("t21.asp", b'', head=b'\xef\xbb\xbf' + HEAD.encode() + b'%>'), "BOM")
    # ---- 7-8 SQL ----
    expect("SQL を文字列連結で組み立てる",
           page("t22.asp", 'Function DbQuery(sql, args)\n  DbQuery = 1\nEnd Function\nDim n, r\nn = Request.QueryString("n")\nr = DbQuery("SELECT A.[x] FROM [A] AS A WHERE A.[n] = " & n, Array())\n%>'), "文字列連結")
    for sql, word, what in [
        ("SELECT Nz(J.[件数], 0) AS N FROM [T_受電] AS J", "Nz", "Nz() を使う"),
        ("SELECT A.[x] FROM [A] AS A INNER JOIN [B] AS B ON A.[i] = B.[i] INNER JOIN [C] AS C ON B.[j] = C.[j]", "かっこ", "JOIN 2 つをかっこで囲まない"),
        ("SELECT A.[x], Sum(A.[n]) AS [合計] FROM [A] AS A GROUP BY A.[x] ORDER BY [合計]", "ORDER BY", "ORDER BY に別名"),
        ("SELECT A.[x] FROM [A] AS A LEFT JOIN [B] AS B ON A.[i] = B.[i] AND B.[d] = ?", "ON 句", "ON に ? を書く"),
        ("SELECT CASE WHEN A.[x] = 1 THEN 1 ELSE 0 END AS F FROM [A] AS A", "CASE", "CASE を使う"),
        ("SELECT Count(DISTINCT A.[x]) AS N FROM [A] AS A", "DISTINCT", "Count(DISTINCT) を使う"),
        ("SELECT A.[x] FROM [A] AS A WHERE A.[d] = #2026-08-25#", "#", "日付を # で書く"),
    ]:
        expect("SQL: " + what, sqlcheck.check(sql), word)
    # ---- CSS ----
    expect("CSS: 囲まない表の規則で罫線を消す", csscheck.check("tbody tr:last-child td { border-bottom: 0 }"), ".tw / .sheet")
    expect("CSS: 帳票の中で罫線を消す", csscheck.check(".sheet td { border-bottom: none }"), "罫線を消して")
    # ---- 更新用の zip に、現地のデータと設定を入れない ----
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    import package
    old = {"nippou\\data\\日報集計_be.accdb": "a", "nippou\\part\\include\\config.asp": "a", "nippou\\part\\error.asp": "a",
           "nippou\\staff\\include\\config.asp": "a"}
    new = {"nippou\\data\\日報集計_be.accdb": "b", "nippou\\part\\include\\config.asp": "b", "nippou\\part\\error.asp": "b",
           "nippou\\staff\\include\\config.asp": "b", "nippou\\staff\\new.asp": "b"}
    got = package.update_files(old, new)
    if got != ["nippou\\part\\error.asp", "nippou\\staff\\new.asp"]:
        raise Fail(f"更新用の zip の中身の選び方が違います: {got}")
    CHECKS.append("更新用の zip には、データベース（.accdb）と config.asp を決して入れない")
    t = package.placement_table(got, old, new)
    if "nippou\\staff\\new.asp（新しいファイル）" not in t or "上書きしないもの" not in t:
        raise Fail("置き場所の表の書き方が違います:\n" + t)
    CHECKS.append("置き場所の表に、置くフォルダと「上書きしないもの」を書く")

    # ---- 模擬実行でも Response.End の握りつぶしを見つける ----
    site = os.path.join(WORK, "site2", "nippou", "part")
    os.makedirs(site)
    open(os.path.join(site, "t.asp"), "w", encoding="utf-8").write(HEAD + 'On Error Resume Next\nResponse.End\nOn Error GoTo 0\nResponse.Write "続いてしまった"\n%>')
    sim = AspSim(os.path.join(WORK, "simw"), os.path.join(WORK, "site2", "nippou"), None)
    r = sim.request("GET", "/nippou/part/t.asp")
    if not (r.warnings and "続いてしまった" in r.text):
        raise Fail("模擬実行が Response.End の握りつぶしを再現していません")
    CHECKS.append("模擬実行：On Error Resume Next の中の Response.End は止まらない（本物と同じ）ことを再現し、警告する")
    return True


if __name__ == "__main__":
    try:
        run()
    except Fail as e:
        print("NG:", e)
        sys.exit(1)
    for c in CHECKS:
        print("  OK:", c)
    print(f"OK: 検査の道具の自己検査 {len(CHECKS)} 項目")
