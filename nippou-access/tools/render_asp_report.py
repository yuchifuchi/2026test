# -*- coding: utf-8 -*-
"""web/include/sheet.asp が組み立てる帳票を、IIS 無しで HTML に起こす。

帳票の HTML は sheet.asp の VBScript 文字列連結で作られているので、
その組み立て部分を Python に置き換えて、同じ HTML が出るようにしている。
HTML を .py 側に書き写すと sheet.asp と食い違うため、必ず読み取って変換する。

    python3 tools/render_asp_report.py [出力先.html]
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SHEET = os.path.join(HERE, "web", "include", "sheet.asp")

CRLF = "\r\n"


# --- 検証用のサンプル値 (実データ 8/25 に合わせてある) ------------------------

HEADER = {
    "出勤者数": 6, "回線数": 5, "合計": 119,
    "申込": 41, "申込_職員": 1,
    "抽選": 14, "抽選_職員": 0,
    "払込用紙": 0, "払込用紙_職員": 0,
    "商品発送": 3, "商品発送_職員": 0,
    "その他": 61, "その他_職員": 1,
    "内交換": 0, "内返金": 0,
    "特記事項": "昭和100年記念貨幣の抽選結果について、発表日の問合せが集中した。"
                "発表は 9/1（月）である旨を案内した。",
    "職員代替案件": "支払方法の変更依頼 1 件を職員が対応。",
    "要望": "抽選結果の発表日をハガキにも明記していただけると、問合せが減ると思われます。",
}
ATTEND = ["髙嶋 名保美", "石原 裕子", "石井 淳子", "鈴木 由美", "寺本 法子", "顧客Ｇ"]
TASKS = [
    ("①　受注入力", 12), ("②　受注チェック", 8),
    ("③　戻り郵便処理／払込書チェック", 3), ("④　ＤＭ処理／戻り郵便", 0),
    ("⑤　架電", 5), ("⑥　その他（　　　　　　　　）", 0),
    ("⑦　エクセル入力", 21), ("⑧　エクセルチェック", 4),
    ("⑨　アンケート入力", 30), ("⑩　ﾊｶﾞｷﾃﾞｰﾀ入力", 44),
    ("⑪　ハガキデータチェック", 6), ("⑫　新規ｺｰﾄﾞ取り", 2),
    ("⑬　顧客整理", 0),
]
SAMPLE_DATE = "2026-08-25"


class Arr(object):
    """VBScript の配列は names(i) と丸括弧で引くので、呼び出せる形にする。"""

    def __init__(self, values, size):
        self.v = list(values) + [""] * (size + 1 - len(values))

    def __call__(self, i):
        return self.v[int(i)]


def H(v):
    if v is None:
        return ""
    s = str(v)
    for a, b in (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"), ('"', "&quot;")):
        s = s.replace(a, b)
    return s


def Len(v):
    return len("" if v is None else str(v))


def IIfS(cond, a, b):
    return a if cond else b


def Wareki(v):
    y, m, d = (int(x) for x in str(v).split("-"))
    import datetime
    w = "月火水木金土日"[datetime.date(y, m, d).weekday()]
    return "令和%2d年%2d月%2d日（%s）" % (y - 2018, m, d, w)


# --- sheet.asp を Python に移し替える ----------------------------------------

KEYWORDS = {"vbCrLf": "CRLF", "And": "and", "Or": "or", "Not": "not",
            "True": "True", "False": "False"}


def conv_expr(src):
    """VBScript の式を Python の式にする。文字列リテラルの中は触らない。"""
    out, i, n = [], 0, len(src)
    while i < n:
        c = src[i]
        if c == '"':                                   # 文字列リテラル
            j, buf = i + 1, []
            while j < n:
                if src[j] == '"':
                    if j + 1 < n and src[j + 1] == '"':  # "" は " 1 文字
                        buf.append('"')
                        j += 2
                        continue
                    j += 1
                    break
                buf.append(src[j])
                j += 1
            out.append(repr("".join(buf)))
            i = j
        elif c == "'":                                 # ここから行末までコメント
            break
        elif c == "&":
            out.append("+")
            i += 1
        elif src.startswith("<>", i):
            out.append("!=")
            i += 2
        elif c.isalpha() or c == "_" or ord(c) > 127:
            j = i
            while j < n and (src[j].isalnum() or src[j] == "_" or ord(src[j]) > 127):
                j += 1
            word = src[i:j]
            out.append(KEYWORDS.get(word, word))
            i = j
        else:
            out.append(c)
            i += 1
    return "".join(out).strip()


def join_continuations(text):
    """行末の「 _」でつながった行を 1 行にまとめる。"""
    lines, buf, out = text.split("\n"), "", []
    for ln in lines:
        ln = ln.rstrip()
        if ln.endswith(" _"):
            buf += ln[:-2]
        else:
            out.append(buf + ln)
            buf = ""
    if buf:
        out.append(buf)
    return out


def transpile(asp_text):
    """sheet.asp の HTML 組み立て部分だけを Python の関数本体にする。

    DB を読む部分 (DbQuery / Do While …) は読み飛ばし、その代わりに
    サンプル値を渡す。罫線や桁数の検証はそれで足りる。
    """
    body, indent = [], 1
    skipping = False
    for raw in join_continuations(asp_text):
        s = raw.strip()
        if not s or s.startswith("'") or s in ("<%", "%>"):
            continue
        if skipping:                                    # Do While … Loop の中身
            if s == "Loop":
                skipping = False
            continue
        low = s.lower()
        if low.startswith("do while"):
            skipping = True
            continue
        if (low.startswith("dim ") or low.startswith("set ")
                or low.startswith("ensure") or ".close" in low
                or low.startswith("end function")):
            continue
        if low.startswith("function "):
            name = s.split("(")[0].split()[1]
            args = s[s.index("(") + 1:s.rindex(")")]
            body.append("def %s(%s):" % (name, args))
            body.append("    _ret = ''")
            indent = 1
            continue
        m = re.match(r"For\s+(\w+)\s*=\s*(\d+)\s+To\s+(\d+)\s*$", s, re.I)
        if m:
            body.append("    " * indent + "for %s in range(%s, %s + 1):"
                        % (m.group(1), m.group(2), m.group(3)))
            indent += 1
            continue
        if low == "next":
            indent -= 1
            continue
        m = re.match(r"([A-Za-z_]\w*)\s*=\s*(.+)$", s)
        if m:
            target, expr = m.group(1), conv_expr(m.group(2))
            if target in ("SheetHtml", "Cell", "Memo"):
                target = "_ret"
            body.append("    " * indent + "%s = %s" % (target, expr))
            continue
        raise SystemExit("読み取れない行: " + s)

    # 関数ごとに return を足す
    src, cur = [], None
    for ln in body + ["def __end__():"]:
        if ln.startswith("def "):
            if cur is not None:
                src.append("    return _ret")
            cur = ln
            src.append(ln)
        else:
            src.append(ln)
    src = src[:-1]                                      # 番人の def を落とす
    src.append("    return _ret")
    return "\n".join(src)


def render():
    asp = io.open(SHEET, encoding="utf-8").read()
    code = transpile(asp)

    env = {"CRLF": CRLF, "H": H, "Len": Len, "IIfS": IIfS, "Wareki": Wareki}
    env["hd"] = lambda k: HEADER.get(k, "")
    env["names"] = Arr(ATTEND, 14)
    env["leftN"] = Arr([t[0] for t in TASKS[:8]], 7)
    env["leftV"] = Arr([t[1] for t in TASKS[:8]], 7)
    env["rightN"] = Arr([t[0] for t in TASKS[8:13]], 7)
    env["rightV"] = Arr([t[1] for t in TASKS[8:13]], 7)
    exec(compile(code, "sheet.asp", "exec"), env)
    return env["SheetHtml"](SAMPLE_DATE), code


def main():
    sheet, code = render()
    css = io.open(os.path.join(HERE, "web", "css", "style.css"), encoding="utf-8").read()
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "web", "_report_preview.html")

    # report.asp と同じ外枠 (PageHead / PageFoot 相当) を付ける。
    # @media print が効くかどうかも、これで一緒に確かめられる。
    html = ("<!doctype html><html lang='ja'><head><meta charset='utf-8'>"
            "<title>日報 印刷</title><style>" + css + "</style></head><body>"
            "<header class='appbar'><a class='brand' href='#'>電話応対日報 集計システム</a>"
            "<nav><a href='#'>メニュー</a><a href='#' class='on'>帳票印刷</a></nav></header>"
            "<main><h1>日報 印刷</h1><p class='lead'>画面の見出し（印刷では消える）</p>"
            + sheet +
            "</main><footer>電話応対日報 集計システム</footer></body></html>")
    io.open(out, "w", encoding="utf-8").write(html)

    # printpdf.asp が変換プログラムに渡すのと同じ、帳票だけの HTML も出す
    pdfhtml = ("<!doctype html><html lang='ja'><head><meta charset='utf-8'>"
               "<title>電話応対報告書日報集計表</title><style>" + css +
               " body{background:#fff;margin:0;padding:0}"
               " .sheet{width:auto;margin:0;padding:0;border:0;"
               "border-radius:0;box-shadow:none}"
               "</style></head><body>" + sheet + "</body></html>")
    out2 = os.path.splitext(out)[0] + "_pdf.html"
    io.open(out2, "w", encoding="utf-8").write(pdfhtml)

    print("画面 (report.asp 相当):", out)
    print("PDF (printpdf.asp 相当):", out2)


if __name__ == "__main__":
    main()
