# -*- coding: utf-8 -*-
"""Access ファイルを自動で作る .vbs を生成する。

なぜ必要か
    DB は Access、画面は ASP という構成では、Access 側に VBA は 1 行も要らない。
    必要なのはテーブル・マスタ・クエリだけで、すべて SQL で作れる。
    それなのに「VBE を開いて 7 個インポートして F5」という手順を課していた。
    この .vbs をダブルクリックすれば、その工程がまるごと消える。

    おまけに、VBA を含まない .accdb はマクロの警告も出ない。

どこから SQL を取るか
    src/*.bas を唯一の出典とし、そこから SQL を機械的に抜き出す。
    手で書き写すと必ずずれるので、同じ内容を二重に持たない。
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(HERE, "src")


def logical_lines(path):
    """VBA の行継続 (_) を連結して、1 文ずつ返す。"""
    out, buf = [], ""
    for raw in io.open(path, encoding="utf-8"):
        line = raw.rstrip("\r\n")
        stripped = line.rstrip()
        if stripped.endswith(" _"):
            buf += stripped[:-1]
            continue
        buf += stripped
        out.append(buf)
        buf = ""
    if buf:
        out.append(buf)
    return out


def vba_strings(text):
        # VBA の "..." を取り出す。"" は 1 個の " を表す。
    parts, i, n = [], 0, len(text)
    while i < n:
        if text[i] != '"':
            if text[i] == "'":            # 以降はコメント
                break
            i += 1
            continue
        i += 1
        cur = []
        while i < n:
            if text[i] == '"':
                if i + 1 < n and text[i + 1] == '"':
                    cur.append('"')
                    i += 2
                    continue
                i += 1
                break
            cur.append(text[i])
            i += 1
        parts.append("".join(cur))
    return parts


def concat_after(stmt, keyword):
    """`keyword` に続く連結された文字列リテラルを 1 本につなげて返す。"""
    body = stmt.split(keyword, 1)[1]
    return "".join(vba_strings(body))


def collect():
    """DDL・マスタ投入・クエリ定義を、実行順に取り出す。"""
    ddl, index, master, queries, rels = [], [], [], [], []

    # --- テーブルとインデックス (modSetup.CreateTables) ---
    for stmt in logical_lines(os.path.join(SRC, "modSetup.bas")):
        s = stmt.strip()
        if s.startswith("ExecDDL "):
            sql = concat_after(s, "ExecDDL ")
            (index if sql.upper().startswith("CREATE INDEX") else ddl).append(sql)

        # 実績→マスタの参照整合性。SQL ではなく DAO の API で作る。
        m = re.match(r'^AddRelation\s+(.+)$', s)
        if m:
            args = vba_strings(m.group(1))
            if len(args) == 4:
                rels.append(tuple(args))

    # --- 取込ログ (modImportExcel。移行のときだけ使うが、表は作っておく) ---
    for stmt in logical_lines(os.path.join(SRC, "modImportExcel.bas")):
        s = stmt.strip()
        if s.startswith("ExecDDL "):
            ddl.append(concat_after(s, "ExecDDL "))

    # --- マスタ (modSetupMaster) ---
    for stmt in logical_lines(os.path.join(SRC, "modSetupMaster.bas")):
        s = stmt.strip()
        if s.startswith("ExecSQL "):
            sql = concat_after(s, "ExecSQL ")
            if sql.upper().startswith("DELETE FROM"):
                continue                   # 新規作成なので消す対象が無い
            master.append(sql)

    # --- クエリ (modSetupQuery) ---
    acc, name_for_var = "", None
    for stmt in logical_lines(os.path.join(SRC, "modSetupQuery.bas")):
        s = stmt.strip()
        m = re.match(r'^s\s*=\s*s\s*&\s*(.+)$', s)
        if m:
            acc += "".join(vba_strings(m.group(1)))
            continue
        m = re.match(r'^s\s*=\s*(".+)$', s)
        if m:
            acc = "".join(vba_strings(m.group(1)))
            continue
        if s.startswith("SaveQuery "):
            body = s[len("SaveQuery "):]
            if body.rstrip().endswith(", s"):
                nm = vba_strings(body)[0]
                queries.append((nm, acc))
                acc = ""
            else:
                lits = vba_strings(body)
                queries.append((lits[0], "".join(lits[1:])))
    return ddl, index, master, queries, rels


def vbs_lit(s):
    """VBScript の文字列リテラルにする。長い行は連結して 1023 文字制限を避ける。"""
    esc = s.replace('"', '""')
    if len(esc) <= 400:
        return '"%s"' % esc
    chunks = [esc[i:i + 400] for i in range(0, len(esc), 400)]
    return (" & _\n        ").join('"%s"' % c for c in chunks)


def build():
    ddl, index, master, queries, rels = collect()
    L = []
    a = L.append

    a("' " + "=" * 72)
    a("'  電話応対日報 集計システム")
    a("'  データベース (Access) を自動で作ります")
    a("'")
    a("'  このファイルをダブルクリックすると、同じフォルダに")
    a("'  日報集計_be.accdb を作ります。10 秒ほどで終わります。")
    a("'")
    a("'  Access がインストールされているパソコンで実行してください。")
    a("'  作られる accdb には VBA が入らないので、マクロの警告も出ません。")
    a("'")
    a("'  ※ このファイルは tools/gen_installer_vbs.py が src/*.bas から")
    a("'     自動生成しています。手で編集せず、生成し直してください。")
    a("' " + "=" * 72)
    a("Option Explicit")
    a("")
    a("Dim fso, shell, here, dbPath, acc, db, i, total, done")
    a('Set fso = CreateObject("Scripting.FileSystemObject")')
    a("here = fso.GetParentFolderName(WScript.ScriptFullName)")
    a('dbPath = fso.BuildPath(here, "日報集計_be.accdb")')
    a("")
    a("' すでにある場合は作り直してよいか確認する")
    a("If fso.FileExists(dbPath) Then")
    a('  If MsgBox("すでに 日報集計_be.accdb があります。" & vbCrLf & vbCrLf & _')
    a('            "作り直すと、入力済みのデータはすべて消えます。" & vbCrLf & _')
    a('            "本当に作り直しますか？", _')
    a('            vbExclamation + vbYesNo + vbDefaultButton2, "確認") <> vbYes Then')
    a('    MsgBox "中止しました。", vbInformation, "電話応対日報"')
    a("    WScript.Quit")
    a("  End If")
    a("  On Error Resume Next")
    a("  fso.DeleteFile dbPath")
    a("  If Err.Number <> 0 Then")
    a('    MsgBox "古いファイルを消せませんでした。" & vbCrLf & vbCrLf & _')
    a('           "Access で開いたままになっていないか確認してください。", _')
    a('           vbCritical, "電話応対日報"')
    a("    WScript.Quit")
    a("  End If")
    a("  On Error GoTo 0")
    a("End If")
    a("")
    a("On Error Resume Next")
    a('Set acc = CreateObject("Access.Application")')
    a("If Err.Number <> 0 Then")
    a('  MsgBox "Access が見つかりませんでした。" & vbCrLf & vbCrLf & _')
    a('         "このパソコンに Microsoft Access が入っているか確認してください。", _')
    a('         vbCritical, "電話応対日報"')
    a("  WScript.Quit")
    a("End If")
    a("On Error GoTo 0")
    a("")
    a("acc.Visible = False")
    a("acc.NewCurrentDatabase dbPath")
    a("Set db = acc.CurrentDb")
    a("")
    a("done = 0")
    a("total = %d" % (len(ddl) + len(index) + len(master) + len(queries) + len(rels)))
    a("")

    a("' --- 表を作る ---")
    for sql in ddl + index:
        a("Run %s" % vbs_lit(sql))
    a("")
    a("' --- 選択肢のもとになるデータを入れる ---")
    for sql in master:
        a("Run %s" % vbs_lit(sql))
    a("")
    a("' --- 集計のしかたを登録する ---")
    for nm, sql in queries:
        a("MakeQuery %s, %s" % (vbs_lit(nm), vbs_lit(sql)))
    a("")
    a("' --- 表どうしのつながりを登録する ---")
    a("'     ここが効いていると、実績から使われているマスタは消せなくなる")
    for nm, parent, child, fld in rels:
        a("AddRel %s, %s, %s, %s"
          % (vbs_lit(nm), vbs_lit(parent), vbs_lit(child), vbs_lit(fld)))
    a("")

    a("acc.CloseCurrentDatabase")
    a("acc.Quit")
    a("Set db = Nothing")
    a("Set acc = Nothing")
    a("")
    a('MsgBox "データベースを作りました。" & vbCrLf & vbCrLf & _')
    a('       dbPath & vbCrLf & vbCrLf & _')
    a('       "このファイルを共有フォルダに置いてください。" & vbCrLf & _')
    a('       "置いた場所は、あとで Web サーバーの設定に書きます。", _')
    a('       vbInformation, "電話応対日報"')
    a("")
    a("' " + "-" * 72)
    a("Sub Run(sql)")
    a("  On Error Resume Next")
    a("  db.Execute sql, 128        ' 128 = dbFailOnError")
    a("  If Err.Number <> 0 Then Fail sql, Err.Description")
    a("  On Error GoTo 0")
    a("  done = done + 1")
    a("End Sub")
    a("")
    a("Sub MakeQuery(nm, sql)")
    a("  On Error Resume Next")
    a("  db.CreateQueryDef nm, sql")
    a("  If Err.Number <> 0 Then Fail nm, Err.Description")
    a("  On Error GoTo 0")
    a("  done = done + 1")
    a("End Sub")
    a("")
    a("Sub AddRel(nm, parentTable, childTable, fld)")
    a("  Dim rel")
    a("  On Error Resume Next")
    a("  Set rel = db.CreateRelation(nm, parentTable, childTable, 0)   ' 0 = 整合性を守る")
    a("  rel.Fields.Append rel.CreateField(fld)")
    a("  rel.Fields(fld).ForeignName = fld")
    a("  db.Relations.Append rel")
    a("  If Err.Number <> 0 Then Err.Clear    ' つながりは作れなくても致命傷ではない")
    a("  On Error GoTo 0")
    a("  done = done + 1")
    a("End Sub")
    a("")
    a("Sub Fail(what, why)")
    a("  On Error Resume Next")
    a("  acc.CloseCurrentDatabase")
    a("  acc.Quit")
    a("  On Error GoTo 0")
    a('  MsgBox "作成に失敗しました。" & vbCrLf & vbCrLf & _')
    a('         "処理: " & Left(what, 120) & vbCrLf & _')
    a('         "理由: " & why & vbCrLf & vbCrLf & _')
    a('         "この画面を写真に撮って、担当者にお知らせください。", _')
    a('         vbCritical, "電話応対日報"')
    a("  WScript.Quit")
    a("End Sub")
    return "\n".join(L) + "\n", (len(ddl), len(index), len(master), len(queries), len(rels))


if __name__ == "__main__":
    text, counts = build()
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "dist", "データベースを作る.vbs")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    # 日本語 Windows の WSH がそのまま読めるよう Shift_JIS + CRLF で出す
    with open(out, "wb") as f:
        f.write(text.replace("\n", "\r\n").encode("cp932"))
    print("生成: %s" % out)
    print("  表 %d / 索引 %d / マスタ投入 %d / クエリ %d / つながり %d = 全 %d 手順"
          % (counts[0], counts[1], counts[2], counts[3], counts[4], sum(counts)))
