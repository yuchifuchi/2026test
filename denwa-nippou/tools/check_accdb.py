"""作った .accdb を、schema.sql と「別の読み方」で突き合わせる。

BuildAccdb.java の読み取り器に誤りがあっても見逃さないように、ここでは正規表現で
schema.sql を読み直し、Jackcess で開き直した結果（accdb_dump.json）と比べる。

  python3 tools/check_accdb.py <.accdb> <accdb_dump.json>
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TYPE = {"LONG": ("LONG", False), "COUNTER": ("LONG", True), "MEMO": ("MEMO", False), "BIT": ("BOOLEAN", False),
        "DATETIME": ("SHORT_DATE_TIME", False), "TEXT": ("TEXT", False)}


def strip_comments(sql):
    out = []
    for line in sql.splitlines():
        q = False
        cut = len(line)
        for i, c in enumerate(line):
            if c == "'":
                q = not q
            if not q and line[i:i + 2] == "--":
                cut = i
                break
        out.append(line[:cut])
    return "\n".join(out)


def expected():
    sql = strip_comments(open(os.path.join(ROOT, "db", "schema.sql"), encoding="utf-8").read())
    stmts = [s.strip().rstrip(";").strip() for s in re.split(r";\s*\n", sql + "\n") if s.strip()]
    tables, rels, rows = {}, [], {}
    for s in stmts:
        m = re.match(r"CREATE TABLE \[([^\]]+)\] \((.*)\)$", s, re.S)
        if m:
            name, body = m.groups()
            cols, idx = [], []
            parts = re.split(r",\s*\n", body)
            for p in parts:
                p = " ".join(p.split())
                cm = re.match(r"CONSTRAINT \[([^\]]+)\] UNIQUE \((.*)\)$", p)
                if cm:
                    idx.append((cm.group(1), re.findall(r"\[([^\]]+)\]", cm.group(2)), False, True))
                    continue
                cm = re.match(r"CONSTRAINT \[([^\]]+)\] FOREIGN KEY \(\[([^\]]+)\]\) REFERENCES \[([^\]]+)\]\(\[([^\]]+)\]\)$", p)
                if cm:
                    rels.append((cm.group(1), name, cm.group(2), cm.group(3), cm.group(4)))
                    continue
                cm = re.match(r"\[([^\]]+)\] (LONG|COUNTER|MEMO|BIT|DATETIME|TEXT\((\d+)\))( NOT NULL)?( CONSTRAINT \[([^\]]+)\] PRIMARY KEY)?$", p)
                if not cm:
                    raise SystemExit(f"NG: schema.sql の列定義が読めません: {p}")
                t = cm.group(2).split("(")[0]
                cols.append((cm.group(1), t, int(cm.group(3) or 0), bool(cm.group(4))))
                if cm.group(5):
                    idx.append((cm.group(6), [cm.group(1)], True, True))
            tables[name] = {"cols": cols, "idx": idx}
            continue
        m = re.match(r"CREATE INDEX \[([^\]]+)\] ON \[([^\]]+)\] \(\[([^\]]+)\]\)$", s)
        if m:
            tables[m.group(2)]["idx"].append((m.group(1), [m.group(3)], False, False))
            continue
        m = re.match(r"ALTER TABLE \[([^\]]+)\] ADD CONSTRAINT \[([^\]]+)\] FOREIGN KEY \(\[([^\]]+)\]\) REFERENCES \[([^\]]+)\]\(\[([^\]]+)\]\)$", s)
        if m:
            rels.append((m.group(2), m.group(1), m.group(3), m.group(4), m.group(5)))
            continue
        m = re.match(r"INSERT INTO \[([^\]]+)\] \((.*?)\) VALUES \((.*)\)$", s, re.S)
        if m:
            cols = re.findall(r"\[([^\]]+)\]", m.group(2))
            vals = re.findall(r"'(?:[^']|'')*'|NULL|TRUE|FALSE|-?\d+", m.group(3))
            conv = []
            for v in vals:
                if v.startswith("'"):
                    conv.append(v[1:-1].replace("''", "'"))
                elif v == "NULL":
                    conv.append(None)
                elif v in ("TRUE", "FALSE"):
                    conv.append(v == "TRUE")
                else:
                    conv.append(int(v))
            rows.setdefault(m.group(1), []).append(dict(zip(cols, conv)))
            continue
        raise SystemExit(f"NG: schema.sql の文が読めません: {s[:60]}")
    return tables, rels, rows


def main():
    accdb, dump = sys.argv[1], sys.argv[2]
    probs = []
    head = open(accdb, "rb").read(0x20)
    if head[0x14] != 0x02:
        probs.append(f"ファイル形式のバイト（0x14）が 0x{head[0x14]:02X} です。0x02（Access 2007-2016 形式）でなければなりません")
    if head[4:19] != b"Standard ACE DB":
        probs.append("ファイルの先頭が Standard ACE DB ではありません")
    d = json.load(open(dump, encoding="utf-8"))
    if d["formatByte"] != 2 or d["fileFormat"] != "V2007":
        probs.append(f"Jackcess の読み直し結果が V2007 / 0x02 ではありません: {d['fileFormat']} / {d['formatByte']}")
    tables, rels, rows = expected()
    got = {t["name"]: t for t in d["tables"]}
    if set(got) != set(tables):
        probs.append(f"表の顔ぶれが違います: 余分 {set(got) - set(tables)} / 不足 {set(tables) - set(got)}")
    for name, spec in tables.items():
        t = got.get(name)
        if not t:
            continue
        gcols = t["columns"]
        if [c["name"] for c in gcols] != [c[0] for c in spec["cols"]]:
            probs.append(f"{name}: 列の並びが違います")
        for (cn, ty, ln, nn), gc in zip(spec["cols"], gcols):
            et, auto = TYPE[ty]
            if gc["type"] != et or gc["autoNumber"] != auto:
                probs.append(f"{name}.{cn}: 型が違います（{gc['type']} / 自動番号 {gc['autoNumber']}）")
            if ty == "TEXT" and gc["lengthInUnits"] != ln:
                probs.append(f"{name}.{cn}: 長さが違います（{gc['lengthInUnits']}、正しくは {ln}）")
            want_req = nn and ty not in ("BIT", "COUNTER")
            if gc["required"] != want_req:
                probs.append(f"{name}.{cn}: 値要求（NOT NULL）が違います")
            if ty in ("TEXT", "MEMO") and gc["allowZeroLength"] is not False:
                probs.append(f"{name}.{cn}: 空文字を許す設定になっています")
        gidx = {i["name"]: i for i in t["indexes"]}
        for iname, icols, pk, uq in spec["idx"]:
            gi = gidx.get(iname)
            if not gi:
                probs.append(f"{name}: 索引 {iname} がありません")
                continue
            if gi["columns"] != icols or gi["pk"] != pk or gi["unique"] != uq:
                probs.append(f"{name}: 索引 {iname} の中身が違います（{gi}）")
        want_rows = rows.get(name, [])
        if t["rows"] != len(want_rows):
            probs.append(f"{name}: 行数が {t['rows']} です（正しくは {len(want_rows)}）")
        for w, g in zip(want_rows, t["data"]):
            for k, v in w.items():
                if g.get(k) != v:
                    probs.append(f"{name}: 初期データが違います（{k}: {g.get(k)!r} ≠ {v!r}）")
    grels = {r["name"]: r for r in d["relationships"]}
    for rn, ft, fc, pt, pc in rels:
        g = grels.get(rn)
        if not g:
            probs.append(f"関連 {rn} がありません")
            continue
        if not g["integrity"]:
            probs.append(f"関連 {rn} に参照整合性がありません")
        if (g["primaryTable"], g["primaryColumns"], g["foreignTable"], g["foreignColumns"]) != (pt, [pc], ft, [fc]):
            probs.append(f"関連 {rn} の向きか列が違います: {g}")
    if len(grels) != len(rels):
        probs.append(f"関連の数が違います（{len(grels)}、正しくは {len(rels)}）")
    # 出荷の決まり
    if got.get("M_担当者", {}).get("rows") != 0:
        probs.append("M_担当者 が空ではありません（空で出荷する決まり）")
    if not any(r.get("製品ID") == 0 for r in got.get("M_製品", {}).get("data", [])):
        probs.append("M_製品 に 製品ID = 0（製品指定なし）の行がありません")
    for t in ("T_日報", "T_出勤", "T_受電", "T_業務実績"):
        if got.get(t, {}).get("rows"):
            probs.append(f"{t} に行があります（実績は空で出荷する）")
    if probs:
        for p in probs:
            print("NG:", p)
        sys.exit(1)
    print(f"OK: .accdb の検査（形式 0x02、表 {len(tables)}、関連 {len(rels)}、初期行 {sum(len(v) for v in rows.values())}）")


if __name__ == "__main__":
    main()
