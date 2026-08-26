# -*- coding: utf-8 -*-
"""data/*.csv から Access 投入用の VBA モジュール src/modSetupMaster.bas を生成する。

CSV を Access に読ませると文字コードで事故るので、INSERT 文を VBA に直接埋め込む。
VBA のソースは VBE に貼り付ければ文字化けしないため、この方式が最も確実。
"""
import csv
import io
import os

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "src", "modSetupMaster.bas")

# VBA の 1 行あたりの文字数上限 (1023) と継続行数上限 (24) を避けるため、
# 1 つの Sub に詰め込みすぎないよう分割する。
CHUNK = 60


def lit(v, kind):
    """CSV の値を VBA/Jet SQL のリテラルに変換する。"""
    v = (v or "").strip()
    if kind == "n":
        return v if v not in ("", "None") else "0"
    if kind == "b":
        return "True" if v in ("1", "True", "true") else "False"
    if kind == "d":
        return "Null" if not v else "#%s#" % v
    if v == "":
        return "Null"
    return "'" + v.replace("'", "''") + "'"


def load(name):
    with io.open(os.path.join(DATA, name), encoding="utf-8") as f:
        return list(csv.DictReader(f))


def emit_table(out, sub_name, table, columns, rows, transform=None):
    """INSERT 文の並びを生成する。columns は (CSV列名, 出力列名, 型) の並び。"""
    stmts = []
    for row in rows:
        if transform:
            row = transform(dict(row))
            if row is None:
                continue
        cols = ",".join("[%s]" % c[1] for c in columns)
        vals = ",".join(lit(row.get(c[0], ""), c[2]) for c in columns)
        stmts.append('INSERT INTO [%s] (%s) VALUES (%s)' % (table, cols, vals))

    parts = [stmts[i:i + CHUNK] for i in range(0, len(stmts), CHUNK)] or [[]]
    names = []
    for i, part in enumerate(parts, start=1):
        nm = sub_name if len(parts) == 1 else "%s_%d" % (sub_name, i)
        names.append(nm)
        out.append("Private Sub %s()" % nm)
        for s in part:
            out.append('    ExecSQL "%s"' % s.replace('"', '""'))
        out.append("End Sub")
        out.append("")
    return names


def main():
    out = [
        'Attribute VB_Name = "modSetupMaster"',
        "Option Compare Database",
        "Option Explicit",
        "",
        "'" + "=" * 78,
        "' マスタ初期データ",
        "'",
        "' このモジュールは tools/gen_master_vba.py が data/*.csv から自動生成する。",
        "' 手で編集せず、CSV を直してから再生成すること。",
        "'",
        "' 出典: 現行の記入用フォーム (Sheet1 の見出しは日報集計管理用フォームへの外部リンク)",
        "'       および Sheet2 の 2 行目 (集計表の列番号) / 3 行目 (旧転記名)。",
        "'" + "=" * 78,
        "",
    ]

    body = []
    calls = []

    calls += emit_table(body, "Master_集計列", "M_集計列",
                        [("集計列ID", "集計列ID", "n"), ("集計列名", "集計列名", "s"),
                         ("表示順", "表示順", "n")],
                        load("M_集計列.csv"))

    blocks = load("M_ブロック.csv")
    blocks.insert(0, {"ブロックID": "0", "ブロック名": "（なし）", "製品別": "0", "表示順": "0"})
    calls += emit_table(body, "Master_ブロック", "M_ブロック",
                        [("ブロックID", "ブロックID", "n"), ("ブロック名", "ブロック名", "s"),
                         ("製品別", "製品別", "b"), ("表示順", "表示順", "n")],
                        blocks)

    # 製品を伴わない区分のために、製品ID = 0 の「製品指定なし」を必ず用意する。
    # これで T_受電.製品ID を NOT NULL にでき、UNIQUE 制約が確実に効く。
    products = load("M_製品.csv")
    products.insert(0, {"製品ID": "0", "ブロックID": "0", "製品名": "（製品指定なし）",
                        "表示順": "0", "有効": "1"})
    calls += emit_table(body, "Master_製品", "M_製品",
                        [("製品ID", "製品ID", "n"), ("ブロックID", "ブロックID", "n"),
                         ("製品名", "製品名", "s"), ("表示順", "表示順", "n"),
                         ("有効", "有効", "b")],
                        products)

    def kubun_tr(r):
        # 帳票の「内 交換」「内 返金」欄はこの印で引く (区分名の表記ゆれに依存しないため)
        legacy = r.get("旧転記名", "")
        r["内訳区分"] = {"返品": "返品", "返金": "返金"}.get(legacy, "")
        return r

    calls += emit_table(body, "Master_区分", "M_区分",
                        [("区分ID", "区分ID", "n"), ("ブロックID", "ブロックID", "n"),
                         ("区分名", "区分名", "s"), ("集計列ID", "集計列ID", "n"),
                         ("内訳区分", "内訳区分", "s"), ("表示順", "表示順", "n"),
                         ("有効", "有効", "b"), ("旧転記名", "旧転記名", "s")],
                        load("M_区分.csv"), kubun_tr)

    def op_tr(r):
        note = []
        if r.get("名前設定ミス") == "1":
            note.append("現行 Excel の Sheet1!B1 が『%s』のままで、集計表に別人として"
                        "計上されていた。Access では担当者マスタで管理する。"
                        % r.get("旧B1値", ""))
        r["備考"] = " ".join(note)
        r["在籍開始日"] = ""
        r["在籍終了日"] = ""
        r["カナ"] = ""
        return r

    calls += emit_table(body, "Master_担当者", "M_担当者",
                        [("担当者ID", "担当者ID", "n"), ("担当者コード", "担当者コード", "s"),
                         ("姓", "姓", "s"), ("名", "名", "s"), ("氏名", "氏名", "s"),
                         ("カナ", "カナ", "s"), ("職員区分", "職員区分", "s"),
                         ("在籍開始日", "在籍開始日", "d"), ("在籍終了日", "在籍終了日", "d"),
                         ("表示順", "表示順", "n"), ("有効", "有効", "b"),
                         ("備考", "備考", "s")],
                        load("M_担当者.csv"), op_tr)

    calls += emit_table(body, "Master_業務項目", "M_業務項目",
                        [("業務項目ID", "業務項目ID", "n"), ("番号", "番号", "s"),
                         ("項目名", "項目名", "s"), ("帳票表示名", "帳票表示名", "s"),
                         ("表示順", "表示順", "n"), ("有効", "有効", "b")],
                        load("M_業務項目.csv"))

    out.append("Public Sub Setup_Master()")
    out.append('    Echo_ "マスタを登録しています..."')
    out.append("")
    out.append("    ' マスタは毎回作り直すので、まず全消しする")
    for t in ("M_区分", "M_製品", "M_ブロック", "M_集計列", "M_業務項目", "M_担当者"):
        out.append('    ExecSQL "DELETE FROM [%s]"' % t)
    out.append("")
    for c in calls:
        out.append("    %s" % c)
    out.append("")
    out.append('    Echo_ "  マスタ登録 完了"')
    out.append("End Sub")
    out.append("")
    out += body

    with io.open(OUT, "w", encoding="utf-8", newline="\r\n") as f:
        f.write("\n".join(out))
    print("生成: %s (%d 行)" % (OUT, len(out)))


if __name__ == "__main__":
    main()
