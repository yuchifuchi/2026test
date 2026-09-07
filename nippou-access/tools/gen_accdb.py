# -*- coding: utf-8 -*-
"""Access ファイル (.accdb) を、Windows も Access も使わずに作る。

なぜ必要か
    「Access を持っている人が .vbs をダブルクリックする」という手順を、
    納品物から消すため。できあがった .accdb をそのまま共有フォルダに置けば、
    その日から使える。Access は 1 台も要らない。

どうやって作るか
    Jackcess (Java の Access ファイル読み書きライブラリ) を使う。
    テーブル・主キー・インデックス・リレーションシップ・マスタの中身まで、
    本物の .accdb (V2016 形式) として書き出せる。

出どころは 1 つ
    テーブル定義もマスタも src/*.bas から機械的に取り出す。
    .vbs 版 (tools/gen_installer_vbs.py) とまったく同じ出典なので、
    どちらで作っても中身は同じになる。

    ただし「クエリ (Q_...)」だけは Jackcess では作れない。
    そのため Web 版は保存クエリに依存せず、SQL を web/include/sql.asp に
    持っている。Access ファイルは純粋に「データの置き場所」になっている。

使い方
    python3 tools/gen_accdb.py [出力先.accdb]
    python3 tools/gen_accdb.py --with-history <xlsm フォルダ> 出力先.accdb
"""
import io
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gen_installer_vbs as g                                    # noqa: E402

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JAR_DIR = os.environ.get("JACKCESS_DIR", "/tmp/jars")
JARS = [
    ("jackcess-4.0.5.jar",
     "https://repo1.maven.org/maven2/com/healthmarketscience/jackcess/"
     "jackcess/4.0.5/jackcess-4.0.5.jar"),
    ("commons-lang3-3.12.0.jar",
     "https://repo1.maven.org/maven2/org/apache/commons/commons-lang3/"
     "3.12.0/commons-lang3-3.12.0.jar"),
    ("commons-logging-1.2.jar",
     "https://repo1.maven.org/maven2/commons-logging/commons-logging/"
     "1.2/commons-logging-1.2.jar"),
]


# --- SQL を読み取る ----------------------------------------------------------

def split_top(s):
    """かっこの外側のカンマで分ける。"""
    out, depth, cur, quote = [], 0, [], False
    for c in s:
        if quote:
            cur.append(c)
            if c == "'":
                quote = False
            continue
        if c == "'":
            quote = True
        elif c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
        elif c == "," and depth == 0:
            out.append("".join(cur).strip())
            cur = []
            continue
        cur.append(c)
    if "".join(cur).strip():
        out.append("".join(cur).strip())
    return out


TYPES = {"LONG": "LONG", "COUNTER": "COUNTER", "MEMO": "MEMO",
         "DATETIME": "DATETIME", "BIT": "BIT"}


def parse_create_table(sql):
    m = re.match(r"CREATE TABLE \[([^\]]+)\] \((.*)\)$", sql, re.S)
    if not m:
        raise SystemExit("読み取れない CREATE TABLE: " + sql[:80])
    tbl = {"name": m.group(1), "columns": [], "pk": [],
           "unique": [], "indexes": [], "fk": []}
    for item in split_top(m.group(2)):
        if item.upper().startswith("CONSTRAINT"):
            u = re.match(r"CONSTRAINT \[([^\]]+)\] UNIQUE \((.*)\)$", item, re.S)
            if u:
                cols = [c.strip(" []") for c in u.group(2).split(",")]
                tbl["unique"].append({"name": u.group(1), "cols": cols})
                continue
            f = re.match(r"CONSTRAINT \[([^\]]+)\] FOREIGN KEY \(\[([^\]]+)\]\) "
                         r"REFERENCES \[([^\]]+)\]\(\[([^\]]+)\]\)$", item, re.S)
            if f:
                tbl["fk"].append({"name": f.group(1), "column": f.group(2),
                                  "parent": f.group(3), "parentColumn": f.group(4)})
                continue
            raise SystemExit("読み取れない制約: " + item)

        c = re.match(r"\[([^\]]+)\]\s+([A-Z]+)(?:\((\d+)\))?(.*)$", item, re.S)
        if not c:
            raise SystemExit("読み取れない列: " + item)
        rest = c.group(4).upper()
        col = {"name": c.group(1),
               "type": "TEXT" if c.group(2) == "TEXT" else TYPES[c.group(2)],
               "len": int(c.group(3)) if c.group(3) else None,
               "notNull": "NOT NULL" in rest}
        tbl["columns"].append(col)
        if "PRIMARY KEY" in rest:
            tbl["pk"].append(c.group(1))
    return tbl


def parse_create_index(sql, tables):
    m = re.match(r"CREATE INDEX \[([^\]]+)\] ON \[([^\]]+)\] \((.*)\)$", sql, re.S)
    if not m:
        raise SystemExit("読み取れない CREATE INDEX: " + sql)
    cols = [c.strip(" []") for c in m.group(3).split(",")]
    tables[m.group(2)]["indexes"].append({"name": m.group(1), "cols": cols})


def parse_values(s):
    """VALUES (...) の中身を Python の値にする。"""
    out, i, n = [], 0, len(s)
    cur = []
    quote = False
    while i < n:
        c = s[i]
        if quote:
            if c == "'":
                if i + 1 < n and s[i + 1] == "'":
                    cur.append("'")
                    i += 2
                    continue
                quote = False
                i += 1
                continue
            cur.append(c)
            i += 1
            continue
        if c == "'":
            quote = True
            cur.append("\0STR")            # 文字列だった印
            i += 1
            continue
        if c == ",":
            out.append("".join(cur).strip())
            cur = []
            i += 1
            continue
        cur.append(c)
        i += 1
    out.append("".join(cur).strip())

    vals = []
    for v in out:
        if v.startswith("\0STR"):
            vals.append(v[4:])
        elif v == "Null":
            vals.append(None)
        elif v == "True":
            vals.append(True)
        elif v == "False":
            vals.append(False)
        else:
            vals.append(int(v))
    return vals


def parse_insert(sql):
    m = re.match(r"INSERT INTO \[([^\]]+)\] \((.*?)\) VALUES \((.*)\)$", sql, re.S)
    if not m:
        raise SystemExit("読み取れない INSERT: " + sql[:80])
    cols = [c.strip(" []") for c in m.group(2).split(",")]
    vals = parse_values(m.group(3))
    if len(cols) != len(vals):
        raise SystemExit("列と値の数が合いません: " + sql[:80])
    return m.group(1), dict(zip(cols, vals))


def build_schema():
    ddl, index, master, queries, rels = g.collect()
    tables, order = {}, []
    for sql in ddl:
        t = parse_create_table(sql)
        tables[t["name"]] = t
        order.append(t["name"])
    for sql in index:
        parse_create_index(sql, tables)

    relationships = []
    for name, parent, child, col in rels:
        relationships.append({"name": name, "parent": parent, "child": child,
                              "parentColumn": col, "childColumn": col})
    for t in tables.values():                     # FOREIGN KEY 制約の分
        for fk in t["fk"]:
            relationships.append({"name": fk["name"], "parent": fk["parent"],
                                  "child": t["name"],
                                  "parentColumn": fk["parentColumn"],
                                  "childColumn": fk["column"]})

    data = {}
    for sql in master:
        tbl, row = parse_insert(sql)
        data.setdefault(tbl, []).append(row)

    return {"tables": [tables[n] for n in order],
            "relationships": relationships,
            "data": data,
            "queries": [{"name": n, "sql": s} for n, s in queries]}


# --- Java 側を呼ぶ -----------------------------------------------------------

def ensure_jars():
    os.makedirs(JAR_DIR, exist_ok=True)
    for name, url in JARS:
        p = os.path.join(JAR_DIR, name)
        if os.path.exists(p) and os.path.getsize(p) > 10000:
            continue
        print("取得:", name)
        with urllib.request.urlopen(url) as r, open(p, "wb") as f:
            shutil.copyfileobj(r, f)
    return os.pathsep.join(os.path.join(JAR_DIR, n) for n, _ in JARS)


def main():
    args = [a for a in sys.argv[1:]]
    history = None
    if "--with-history" in args:
        i = args.index("--with-history")
        history = args[i + 1]
        del args[i:i + 2]
    out = args[0] if args else os.path.join(HERE, "dist", "日報集計_be.accdb")
    out = os.path.abspath(out)

    schema = build_schema()

    if history:
        sys.path.insert(0, os.path.join(HERE, "tools"))
        import migrate_rows                                   # noqa: E402
        rows = migrate_rows.build_rows(history)
        schema["data"]["T_受電"] = rows
        print("過去データ: %d 行" % len(rows))

    work = os.path.join(HERE, "tools", "accdb")
    sfile = os.path.join(work, "_schema.json")
    io.open(sfile, "w", encoding="utf-8").write(
        json.dumps(schema, ensure_ascii=False, indent=1))

    cp = ensure_jars()
    classes = os.path.join(work, "_classes")
    os.makedirs(classes, exist_ok=True)
    subprocess.check_call(["javac", "-encoding", "UTF-8", "-cp", cp,
                           "-d", classes, os.path.join(work, "BuildAccdb.java")],
                          stderr=subprocess.STDOUT)
    if os.path.exists(out):
        os.remove(out)
    subprocess.check_call(["java", "-Dfile.encoding=UTF-8",
                           "-cp", cp + os.pathsep + classes,
                           "BuildAccdb", sfile, out])
    print("出力: %s (%.0f KB)" % (out, os.path.getsize(out) / 1024.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
