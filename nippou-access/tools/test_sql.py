# -*- coding: utf-8 -*-
"""生成した .accdb に、ASP が実際に流す SQL を通してみる。

Windows も IIS も Access も無いところで、
「列名を間違えた」「別名が解決できない」「Access 専用関数を使ってしまった」
の類を見つけるための検査。UCanAccess (JDBC) で実行する。

  ・テスト用の .accdb を作る (マスタ + 小さな実績データ)
  ・web/include/sql.asp から SQL を取り出す
  ・ASP と同じ形に組み立てて実行する
  ・日報の集計値が手計算と合うかまで見る

    python3 tools/test_sql.py
"""
import io
import json
import os
import re
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_accdb                                                # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(HERE, "tools", "accdb")
DB = "/tmp/nippou_test.accdb"

DT = "#2026-08-25#"
D1, D2 = "#2026-08-24#", "#2026-08-28#"

# --- テスト用の実績 ----------------------------------------------------------
#   区分 1=申込(集計列3) 3=抽選(4) 4=払込用紙(5) 9=商品発送(6) 11=交換(8) 57=返金(7)
#   担当者 1=パート / 18=職員
SAMPLE = {
    # 納品する .accdb の担当者は空なので、テストではここで入れる
    "M_担当者": [
        {"担当者ID": 1, "担当者コード": "001", "姓": "試験", "名": "一郎",
         "氏名": "試験 一郎", "職員区分": "パート", "表示順": 1, "有効": True},
        {"担当者ID": 2, "担当者コード": "002", "姓": "試験", "名": "二郎",
         "氏名": "試験 二郎", "職員区分": "パート", "表示順": 2, "有効": True},
        {"担当者ID": 18, "担当者コード": "100", "姓": "試験", "名": "職員",
         "氏名": "試験 職員", "職員区分": "職員", "表示順": 18, "有効": True},
    ],
    "T_日報": [{"対象日": "2026-08-25", "回線数": 5, "状態": "入力中",
                "特記事項": "テスト", "職員代替案件": "", "要望": ""}],
    "T_出勤": [{"対象日": "2026-08-25", "担当者ID": 1},
               {"対象日": "2026-08-25", "担当者ID": 2},
               {"対象日": "2026-08-25", "担当者ID": 18}],
    "T_受電": [
        {"対象日": "2026-08-25", "担当者ID": 1,  "区分ID": 1,  "製品ID": 1, "件数": 10},
        {"対象日": "2026-08-25", "担当者ID": 1,  "区分ID": 3,  "製品ID": 0, "件数": 4},
        {"対象日": "2026-08-25", "担当者ID": 18, "区分ID": 1,  "製品ID": 0, "件数": 2},
        {"対象日": "2026-08-25", "担当者ID": 18, "区分ID": 11, "製品ID": 0, "件数": 1},
        {"対象日": "2026-08-25", "担当者ID": 1,  "区分ID": 57, "製品ID": 0, "件数": 3},
    ],
    "T_業務実績": [{"対象日": "2026-08-25", "担当者ID": 1, "業務項目ID": 1, "件数": 7}],
}

# 手計算した答え。ここが合わなければ集計の考え方が壊れている。
EXPECT = {
    "判定": "職員",
    "合計": 20, "申込": 12, "申込_職員": 2, "抽選": 4, "払込用紙": 0,
    "商品発送": 0, "その他": 4, "その他_職員": 1, "内交換": 1, "内返金": 3,
    "出勤者数": 3, "合計_職員": 3,
}


def read_sql_asp():
    """sql.asp の Function から SQL 文字列を取り出す。"""
    src = io.open(os.path.join(HERE, "web", "include", "sql.asp"),
                  encoding="utf-8").read()
    out = {}
    for m in re.finditer(r"Function\s+(\S+?)\(\)(.*?)End Function", src, re.S):
        name, body, sql = m.group(1), m.group(2), ""
        for line in body.split("\n"):
            lit = re.match(r'\s*s\s*=\s*(?:s\s*&\s*)?"(.*)"\s*$', line)
            if lit:
                sql += lit.group(1).replace('""', '"')
        out[name] = sql
    return out


def cases(q):
    """ASP が実際に流している形。? は日付などのリテラルに置き換えてある。"""
    return [
        # --- 帳票と日報 (include/sheet.asp, daily.asp) ---
        ("日報ヘッダ",
         "SELECT D.[合計],D.[申込],D.[申込_職員],D.[抽選],D.[払込用紙],D.[商品発送],"
         "D.[その他],D.[その他_職員] FROM (%s) AS D WHERE D.[対象日]=%s"
         % (q["SQL_日報_ヘッダ"], DT)),
        ("日報ヘッダ_内訳",
         "SELECT D.[内交換],D.[内返金],D.[出勤者数],D.[合計_職員],D.[回線数],D.[状態] "
         "FROM (%s) AS D WHERE D.[対象日]=%s" % (q["SQL_日報_ヘッダ"], DT)),
        ("日報_出勤",
         "SELECT S.[氏名] FROM (%s) AS S WHERE S.[対象日]=%s ORDER BY S.[表示順]"
         % (q["SQL_日報_出勤"], DT)),
        ("未入力チェック",
         "SELECT * FROM (%s) AS C WHERE C.[対象日]=%s ORDER BY C.[表示順]"
         % (q["SQL_未入力チェック"], DT)),
        ("未入力チェック_件数",
         "SELECT Count(*) FROM (%s) AS C WHERE C.[対象日]=%s"
         % (q["SQL_未入力チェック"], DT)),
        # --- 受付入力 (entry.asp) ---
        ("選択_区分",
         "SELECT K.[区分ID],K.[表示名] FROM (%s) AS K ORDER BY K.[並び順]"
         % q["SQL_選択_区分"]),
        ("選択_製品",
         "SELECT P.[製品ID],P.[表示名] FROM (%s) AS P ORDER BY P.[ブロックID],P.[表示順]"
         % q["SQL_選択_製品"]),
        # --- 集計表 (summary.asp) ---
        ("受電明細",
         "SELECT * FROM (%s) AS M WHERE M.[対象日] BETWEEN %s AND %s "
         "ORDER BY M.[対象日],M.[氏名],M.[区分ID]" % (q["SQL_受電明細"], D1, D2)),
        ("日別集計",
         "SELECT J.[対象日], Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法], "
         "Sum(J.[件数]) AS [計] "
         "FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID] "
         "WHERE J.[対象日] BETWEEN %s AND %s GROUP BY J.[対象日] ORDER BY J.[対象日]"
         % (D1, D2)),
        # --- 帳票の業務欄 (include/sheet.asp) ---
        ("帳票_業務項目",
         "SELECT TM.[業務項目ID],TM.[帳票表示名], "
         "(SELECT IIf(IsNull(Sum(W.[件数])),0,Sum(W.[件数])) FROM [T_業務実績] AS W "
         "WHERE W.[対象日]=%s AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] "
         "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]" % DT),
        # --- 日報の出勤者一覧 (daily.asp) ---
        ("日報_出勤者件数",
         "SELECT OP.[担当者ID],OP.[担当者コード],OP.[氏名],OP.[職員区分], "
         "(SELECT IIf(IsNull(Sum(J.[件数])),0,Sum(J.[件数])) FROM [T_受電] AS J "
         "WHERE J.[対象日]=A.[対象日] AND J.[担当者ID]=A.[担当者ID]) AS [件数] "
         "FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] "
         "WHERE A.[対象日]=%s ORDER BY OP.[表示順]" % DT),
        # --- その他業務 (tasks.asp) ---
        ("その他業務_本人分",
         "SELECT TM.[業務項目ID],TM.[番号],TM.[項目名], "
         "(SELECT IIf(IsNull(Sum(W.[件数])),0,Sum(W.[件数])) FROM [T_業務実績] AS W "
         "WHERE W.[対象日]=%s AND W.[担当者ID]=1 AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] "
         "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]" % DT),
        # --- マスタ保守の使用件数 (master.asp) ---
        ("マスタ_使用件数",
         "SELECT Count(*) FROM [T_受電] WHERE [区分ID]=1"),

        # --- 書き込み。ASP が流すものと同じ形にしてある ---------------------
        ("日報を作る (db.asp)",
         "INSERT INTO [T_日報] ([対象日],[状態],[更新日時]) "
         "VALUES (#2026-08-26#,'入力中',Now())"),
        ("出勤を登録 (db.asp)",
         "INSERT INTO [T_出勤] ([対象日],[担当者ID]) VALUES (#2026-08-26#,1)"),
        ("受付を登録 (entry.asp)",
         "INSERT INTO [T_受電] ([対象日],[担当者ID],[区分ID],[製品ID],[件数],[備考2],"
         "[登録日時],[更新日時],[登録者]) "
         "VALUES (#2026-08-26#,1,1,1,5,'てすと',Now(),Now(),'testuser')"),
        ("同じ行に足す (entry.asp)",
         "UPDATE [T_受電] SET [件数]=[件数]+2, [更新日時]=Now() "
         "WHERE [対象日]=#2026-08-26# AND [担当者ID]=1 AND [区分ID]=1 AND [製品ID]=1"),
        ("件数を直す (entry.asp)",
         "UPDATE [T_受電] SET [件数]=9,[備考2]='なおした',[更新日時]=Now() "
         "WHERE [受電ID]=1"),
        ("日報を保存 (daily.asp)",
         "UPDATE [T_日報] SET [回線数]=6,[特記事項]='あ',[職員代替案件]='い',[要望]='う',"
         "[更新日時]=Now() WHERE [対象日]=#2026-08-26#"),
        ("日報を確定 (daily.asp)",
         "UPDATE [T_日報] SET [状態]='確定',[確定日時]=Now(),[更新日時]=Now() "
         "WHERE [対象日]=#2026-08-26#"),
        ("確定を解除 (daily.asp)",
         "UPDATE [T_日報] SET [状態]='入力中',[確定日時]=Null,[更新日時]=Now() "
         "WHERE [対象日]=#2026-08-26#"),
        ("その他業務を登録 (tasks.asp)",
         "INSERT INTO [T_業務実績] ([対象日],[担当者ID],[業務項目ID],[件数]) "
         "VALUES (#2026-08-26#,1,2,3)"),
        ("担当者を追加 (master.asp)",
         "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[ログオン名],"
         "[職員区分],[在籍開始日],[表示順],[有効]) "
         "VALUES (900,'900','試験','太郎','試験 太郎','シケン','t-shiken','パート',"
         "#2026-04-01#,900,True)"),
        ("担当者を更新 (master.asp)",
         "UPDATE [M_担当者] SET [担当者コード]='900',[姓]='試験',[名]='太郎',[氏名]='試験 太郎',"
         "[カナ]='シケン',[ログオン名]='t-shiken',[職員区分]='職員',[在籍開始日]=#2026-04-01#,"
         "[在籍終了日]=Null,[表示順]=900,[有効]=True WHERE [担当者ID]=900"),
        ("職員かどうか (auth.asp)",
         "SELECT [職員区分] AS [判定] FROM [M_担当者] WHERE [ログオン名]='t-shiken' AND [有効]=True "
         "AND ([在籍終了日] Is Null Or [在籍終了日]>=%s)" % DT),
        ("退職後は職員でなくなる (auth.asp)",
         "SELECT Count(*) AS [残り] FROM [M_担当者] WHERE [ログオン名]='t-shiken' "
         "AND [有効]=True AND ([在籍終了日] Is Null Or [在籍終了日]>=#2026-12-31#) "
         "AND [在籍終了日]=#2026-09-30#"),
        ("退職日を入れる (master.asp)",
         "UPDATE [M_担当者] SET [在籍終了日]=#2026-09-30# WHERE [担当者ID]=900"),
        ("復帰させる (master.asp)",
         "UPDATE [M_担当者] SET [在籍終了日]=Null,[有効]=True WHERE [担当者ID]=900"),
        ("担当者を消す (master.asp)",
         "DELETE FROM [M_担当者] WHERE [担当者ID]=900"),
        ("受付を消す (entry.asp)",
         "DELETE FROM [T_受電] WHERE [受電ID]=6"),
        ("出勤を外す (daily.asp)",
         "DELETE FROM [T_出勤] WHERE [対象日]=#2026-08-26# AND [担当者ID]=1"),
    ]


def main():
    schema = gen_accdb.build_schema()
    for t, rows in SAMPLE.items():
        schema["data"][t] = rows
    sfile = os.path.join(WORK, "_schema_test.json")
    io.open(sfile, "w", encoding="utf-8").write(
        json.dumps(schema, ensure_ascii=False, indent=1))

    cp = gen_accdb.ensure_jars() + os.pathsep + os.pathsep.join(
        os.path.join("/tmp/jars", n) for n in
        ("ucanaccess-5.0.1.jar", "hsqldb-2.5.2.jar", "jackcess-encrypt-4.0.1.jar"))
    classes = os.path.join(WORK, "_classes")
    os.makedirs(classes, exist_ok=True)
    for src in ("BuildAccdb.java", "TestSql.java"):
        subprocess.check_call(["javac", "-encoding", "UTF-8", "-cp", cp,
                               "-d", classes, os.path.join(WORK, src)])
    if os.path.exists(DB):
        os.remove(DB)
    print("--- テスト用の .accdb を作る ---")
    subprocess.check_call(["java", "-Dfile.encoding=UTF-8", "-cp", cp + os.pathsep + classes,
                           "BuildAccdb", sfile, DB],
                          stdout=subprocess.DEVNULL)

    q = read_sql_asp()
    missing = [k for k in ("SQL_日報_ヘッダ", "SQL_未入力チェック", "SQL_選択_区分",
                           "SQL_選択_製品", "SQL_受電明細", "SQL_日報_出勤") if k not in q]
    if missing:
        raise SystemExit("sql.asp から取り出せません: " + ", ".join(missing))

    cfile = os.path.join(WORK, "_cases.json")
    io.open(cfile, "w", encoding="utf-8").write(
        json.dumps(dict(cases(q)), ensure_ascii=False, indent=1))

    print("--- ASP と同じ SQL を流す ---")
    p = subprocess.run(["java", "-Dfile.encoding=UTF-8",
                        "-cp", cp + os.pathsep + classes, "TestSql", DB, cfile],
                       capture_output=True, text=True)
    out = p.stdout
    print(out)
    if p.returncode != 0:
        print(p.stderr[:2000])
        return 1

    # 集計値が手計算と合うか
    ng = 0
    for k, v in EXPECT.items():
        m = re.search(re.escape(k) + r"=(\S+?)(?:,|$)", out, re.M)
        if not m:
            continue
        got = m.group(1).rstrip(",")
        if str(v) != got:
            print("  値が違います: %s 期待 %s / 実際 %s" % (k, v, got))
            ng += 1
    print("集計値の検算: " + ("合致" if ng == 0 else "%d 件ずれ" % ng))
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main())
