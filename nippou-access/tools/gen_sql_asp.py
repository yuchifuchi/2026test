# -*- coding: utf-8 -*-
"""web/include/sql.asp を src/modSetupQuery.bas から生成する。

なぜ Web 版が保存クエリ (Q_...) を使わないのか
  1. Access を 1 台も使わずに .accdb を作れるようにしたため。
     Jackcess (tools/gen_accdb.py) はテーブルは作れるが、保存クエリは作れない。
  2. ACE (OLEDB) 経由では Access 専用の関数が使えないため。
     とくに Nz() は「関数 'Nz' が定義されていません」で失敗する。
     Access の中では動くのに Web からだけ失敗する、という分かりにくい形になる。

そこで、クエリの SQL は Access ではなく ASP 側に持つ。
ただし出どころは Access 版と同じ src/modSetupQuery.bas 1 つに保ち、
この生成器が次の 3 つを機械的に行う。

  ・末尾の ORDER BY と ";" を外す (副問い合わせとして埋め込むため)
  ・Nz(x,0) を IIf(IsNull(x),0,x) に置き換える (ACE で動く形にする)
  ・クエリの中の [Q_...] を、その中身に展開する (入れ子の解決)

    python3 tools/gen_sql_asp.py
"""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_installer_vbs as g                                    # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(HERE, "web", "include", "sql.asp")

# Web 版が使うクエリ (と、その中で使われるもの)
WANT = [
    ("SQL_選択_区分",     "Q_選択_区分",     "受付入力の区分の選択肢"),
    ("SQL_選択_製品",     "Q_選択_製品",     "受付入力の製品の選択肢"),
    ("SQL_受電明細",      "Q_受電明細",      "集計表の明細 1 行 = 受電 1 件"),
    ("SQL_日報_出勤",     "Q_日報_出勤",     "その日の出勤者"),
    ("SQL_日報_ヘッダ",   "Q_日報_ヘッダ",   "日報 1 日分。件数はすべて受電から計算"),
    ("SQL_未入力チェック", "Q_未入力チェック", "出勤しているのに実績が 1 件も無い人"),
]


def strip_tail(sql):
    """末尾の ";" と ORDER BY を外す。副問い合わせとして使うため。"""
    s = re.sub(r"\s+", " ", sql).strip().rstrip(";").strip()
    # 末尾の ORDER BY 句だけを外す (副問い合わせの中の ORDER BY は無い前提)
    m = re.search(r"\sORDER\s+BY\s[^)]*$", s, re.I)
    if m:
        s = s[:m.start()].strip()
    return s


def unnz(sql):
    """Nz(x,0) を IIf(IsNull(x),0,x) にする。ACE では Nz が使えない。"""
    out = sql
    while True:
        m = re.search(r"Nz\(([^(),]+),\s*([^(),]+)\)", out)
        if not m:
            break
        out = (out[:m.start()] +
               "IIf(IsNull(%s),%s,%s)" % (m.group(1), m.group(2), m.group(1)) +
               out[m.end():])
    if "Nz(" in out:
        raise SystemExit("展開できない Nz が残りました: " + out[:120])
    return out


def expand(sql, byname, seen=()):
    """SQL の中の [Q_...] を、そのクエリの中身に置き換える。"""
    for name in sorted(byname, key=len, reverse=True):
        token = "[%s]" % name
        if token not in sql:
            continue
        if name in seen:
            raise SystemExit("クエリが循環しています: " + name)
        inner = expand(strip_tail(byname[name]), byname, seen + (name,))
        sql = sql.replace(token, "(" + inner + ")")
    return sql


def vb_lines(fn, sql, indent):
    """SQL を、読める幅で折り返した VBScript にする。

    「s = s & ...」で足していく形にする。行継続 (_) を長く連ねると、
    処理系によっては上限に当たるため。
    """
    words, line, lines = sql.split(" "), "", []
    for w in words:
        if len(line) + len(w) + 1 > 88:
            lines.append(line)
            line = w
        else:
            line = (line + " " + w) if line else w
    if line:
        lines.append(line)
    body = []
    for i, ln in enumerate(lines):
        lit = '"%s%s"' % (ln.replace('"', '""'), "" if i == len(lines) - 1 else " ")
        body.append("%ss = %s%s" % (indent, "" if i == 0 else "s & ", lit))
    body.append("%s%s = s" % (indent, fn))
    return "\n".join(body)


def validate(name, sql):
    """かっこと引用符の対応を確かめる。ここが崩れると実行時まで気付けない。"""
    depth, quote = 0, False
    for c in sql:
        if quote:
            if c == "'":
                quote = False
            continue
        if c == "'":
            quote = True
        elif c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth < 0:
                raise SystemExit("%s: かっこが閉じすぎです" % name)
    if depth != 0:
        raise SystemExit("%s: かっこが %d 個閉じていません" % (name, depth))
    if quote:
        raise SystemExit("%s: 引用符が閉じていません" % name)
    if not sql.upper().startswith("SELECT"):
        raise SystemExit("%s: SELECT で始まっていません" % name)


def main():
    _ddl, _ix, _master, queries, _rels = g.collect()
    byname = dict(queries)

    L = []
    a = L.append
    a("<%")
    a("' " + "=" * 74)
    a("'  クエリの SQL")
    a("'")
    a("'  Web 版は Access の保存クエリ (Q_...) を使わない。")
    a("'  理由は 2 つ。")
    a("'    1. Access を 1 台も使わずに .accdb を配れるようにするため。")
    a("'    2. ACE 経由では Access 専用の関数 (Nz など) が使えないため。")
    a("'       Access の中では動くのに Web からだけ失敗する、という形を避ける。")
    a("'")
    a("'  使いかた: 副問い合わせとして埋め込み、並び順は呼ぶ側で付ける。")
    a("'      Set rs = DbQuery(\"SELECT * FROM (\" & SQL_未入力チェック() & \") AS C \" & _")
    a("'                       \"WHERE C.[対象日]=? ORDER BY C.[表示順]\", Array(dt))")
    a("'")
    a("'  ※ このファイルは tools/gen_sql_asp.py が src/modSetupQuery.bas から")
    a("'     自動生成しています。手で編集せず、生成し直してください。")
    a("' " + "=" * 74)
    for fn, qn, note in WANT:
        sql = expand(unnz(strip_tail(byname[qn])), byname)
        validate(qn, sql)
        a("")
        a("' %s (Access 版の %s と同じ)" % (note, qn))
        a("Function %s()" % fn)
        a("    Dim s")
        a(vb_lines(fn, sql, "    "))
        a("End Function")
    a("%>")

    io.open(OUT, "w", encoding="utf-8", newline="\n").write("\n".join(L) + "\n")
    print("出力:", OUT)
    for fn, qn, _n in WANT:
        sql = expand(unnz(strip_tail(byname[qn])), byname)
        print("  %-18s %5d 文字" % (fn, len(sql)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
