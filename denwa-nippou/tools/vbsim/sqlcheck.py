"""ACE（本番の Access データベース エンジン）だけが受け付けない SQL の書き方を止める。

UCanAccess（検査で SQL を実際に流す仕組み）は ACE より寛容で、Nz や CASE も通してしまう。
そのため「UCanAccess で通った」だけでは足りず、ここで書き方そのものを検査する。
"""
import re

# ACE OLEDB 経由（＝ ASP から）では使えない、または使わない決まりの書き方
BANNED = [
    (r"\bNz\s*\(", "Nz() は Access アプリの関数で、ASP から（ACE OLEDB 経由）では「関数 'Nz' は定義されていません」で落ちます。IIf(IsNull(x), 0, x) を使う"),
    (r"\bCASE\b", "CASE は ACE にありません（IIf を使う）"),
    (r"\bCOALESCE\s*\(", "COALESCE は ACE にありません（IIf(IsNull(..)) を使う）"),
    (r"\bIFNULL\s*\(|\bNVL\s*\(", "IFNULL / NVL は ACE にありません"),
    (r"\|\|", "|| での文字列連結は ACE にありません（& を使う）"),
    (r"\bLIMIT\b|\bOFFSET\b|\bFETCH\b", "LIMIT / OFFSET は ACE にありません（TOP を使う）"),
    (r"\bCOUNT\s*\(\s*DISTINCT\b", "Count(DISTINCT ...) は ACE にありません（派生表で DISTINCT してから Count）"),
    (r"\bCAST\s*\(|\bCONVERT\s*\(", "CAST / CONVERT は ACE にありません"),
    (r"\bReplace\s*\(", "Replace() は ACE OLEDB 経由では使えません（サンドボックスで止められる）"),
    (r"\bDLookup\s*\(|\bDSum\s*\(|\bDCount\s*\(|\bEval\s*\(", "定義域集計関数・Eval は Access アプリでしか使えません"),
    (r"#\d", "日付を #...# で直接書かない決まりです（? で渡す）"),
    (r"\bSELECT\s+\*|\.\*", "SELECT * は使わない決まりです（列を名前で書く）"),
    (r"\bNOW\s*\(\s*\)|\bDATE\s*\(\s*\)", "Date() / Now() を SQL に書かない決まりです（サーバー側の今日と画面の日付がずれるのを防ぐ。? で渡す）"),
]

ALLOWED_FUNCS = {"iif", "isnull", "sum", "count", "max", "min", "abs", "int", "clng", "year", "month", "day",
                 "dateserial", "len", "exists", "in"}

KW = {"select", "from", "where", "group", "by", "order", "having", "inner", "left", "right", "join", "on", "and",
      "or", "not", "as", "insert", "into", "values", "update", "set", "delete", "distinct", "top", "is", "null",
      "union", "all", "asc", "desc", "true", "false", "exists", "in", "between", "like"}


class SqlProblem(Exception):
    pass


def strip_literals(sql):
    """文字列定数を '' に、[名前] を [x] に置き換えた写しを返す（位置は保たない）。"""
    out, i, n = [], 0, len(sql)
    while i < n:
        c = sql[i]
        if c == "'":
            j = i + 1
            while j < n:
                if sql[j] == "'" and j + 1 < n and sql[j + 1] == "'":
                    j += 2
                    continue
                if sql[j] == "'":
                    break
                j += 1
            out.append("''")
            i = j + 1
            continue
        if c == "[":
            j = sql.find("]", i)
            if j < 0:
                raise SqlProblem("[ が閉じていません")
            out.append("[x]")
            i = j + 1
            continue
        out.append(c)
        i += 1
    return "".join(out)


def tokens(s):
    return re.findall(r"\[x\]|''|\?|[A-Za-z_][A-Za-z0-9_]*|\d+|<>|<=|>=|[()=<>,.&+\-*/]", s)


def check(sql):
    """問題があれば説明の一覧を返す（空なら合格）。"""
    probs = []
    if '"' in sql:
        probs.append('SQL の中で " を使わない決まりです（文字列は \'...\' で書く。ACE と検査用エンジンで " の意味が違う）')
    try:
        s = strip_literals(sql)
    except SqlProblem as e:
        return [str(e)]
    for pat, msg in BANNED:
        if re.search(pat, s, re.I):
            probs.append(msg)
    # 日本語の名前は必ず [ ] で囲む
    for ch in s:
        if ord(ch) > 127:
            probs.append(f"[ ] で囲まれていない日本語があります（{ch}…）")
            break
    toks = tokens(s)
    low = [t.lower() for t in toks]
    # 使ってよい関数だけ
    for i, t in enumerate(low):
        if i + 1 < len(low) and toks[i + 1] == "(" and re.match(r"[a-z_]", t) and t not in KW:
            if t not in ALLOWED_FUNCS:
                probs.append(f"関数 {toks[i]}() は ACE で使えるか確かめていないため使わない決まりです")
        if t in ("in",) and i + 1 < len(low) and toks[i + 1] == "(" and i + 2 < len(low) and low[i + 2] != "select":
            pass
    probs += check_levels(toks)
    return probs


def check_levels(toks):
    """かっこの深さごとに FROM / JOIN / ON / ORDER BY / GROUP BY を調べる。"""
    probs = []
    # かっこの対応
    depth = 0
    for t in toks:
        if t == "(":
            depth += 1
        elif t == ")":
            depth -= 1
            if depth < 0:
                return ["かっこの対応が取れていません"]
    if depth != 0:
        return ["かっこの対応が取れていません"]

    def level_segments(start_idx):
        """start_idx の位置と同じ深さのトークン位置を、終わりのかっこまで集める。"""
        d = 0
        out = []
        for i in range(start_idx, len(toks)):
            t = toks[i]
            if t == "(":
                d += 1
            elif t == ")":
                d -= 1
                if d < 0:
                    break
            if d == 0:
                out.append(i)
        return out

    # SELECT ごとに
    for si, t in enumerate(toks):
        if t.lower() != "select":
            continue
        idxs = level_segments(si)
        words = [(i, toks[i].lower()) for i in idxs]
        # FROM 句の JOIN の数
        from_pos = [i for i, w in words if w == "from"]
        if not from_pos:
            continue
        fpos = from_pos[0]
        # FROM 以降、同じ「SELECT の階層」の JOIN を数える（入れ子のかっこ（JOIN の組）の中も含めるため深さで追う）
        d = 0
        joins = 0
        end = len(toks)
        i = fpos + 1
        while i < len(toks):
            tt = toks[i].lower()
            if tt == "(" and i + 1 < len(toks) and toks[i + 1].lower() == "select":
                # 派生表（サブクエリ）の中は別の SELECT として調べるので、ここでは丸ごと飛ばす
                dd = 0
                while i < len(toks):
                    if toks[i] == "(":
                        dd += 1
                    elif toks[i] == ")":
                        dd -= 1
                        if dd == 0:
                            break
                    i += 1
                i += 1
                continue
            if tt == "(":
                d += 1
            elif tt == ")":
                d -= 1
                if d < 0:
                    end = i
                    break
            elif d == 0 and tt in ("where", "group", "order", "having", "union"):
                end = i
                break
            elif tt == "join":
                joins += 1
            i += 1
        if joins >= 2:
            lead = 0
            k = fpos + 1
            while k < len(toks) and toks[k] == "(":
                lead += 1
                k += 1
            if lead < joins - 1:
                probs.append(f"JOIN が {joins} つあるのに、FROM の直後のかっこが {lead} 個です（ACE は JOIN を 2 つずつかっこで囲む必要があります）")
        # ON 句: X.[a] = Y.[b] (AND ...)* だけ
        for i in range(fpos, end):
            if toks[i].lower() == "on":
                j = i + 1
                cond = []
                dd = 0
                while j < end:
                    tt = toks[j].lower()
                    if tt == "(":
                        dd += 1
                    elif tt == ")":
                        if dd == 0:
                            break
                        dd -= 1
                    if dd == 0 and tt in ("left", "inner", "right", "join", "where", "group", "order"):
                        break
                    cond.append(toks[j])
                    j += 1
                if not on_ok(cond):
                    probs.append("ON 句には「表.[列] = 表.[列]」を AND でつないだものだけを書く決まりです（定数や ? を書くと ACE が「JOIN 式はサポートされていません」で落ちることがある。絞り込みは派生表の中で）: " + " ".join(cond))
        # ORDER BY / GROUP BY は「別名.[列]」だけ
        for i, w in words:
            if w in ("order", "group") and i + 1 < len(toks) and toks[i + 1].lower() == "by":
                j = i + 2
                item = []
                items = []
                dd = 0
                while j < len(toks):
                    tt = toks[j]
                    if tt == "(":
                        dd += 1
                    elif tt == ")":
                        if dd == 0:
                            break
                        dd -= 1
                    if dd == 0 and tt.lower() in ("having", "order", "union", "group"):
                        break
                    if dd == 0 and tt == ",":
                        items.append(item)
                        item = []
                    else:
                        item.append(tt)
                    j += 1
                items.append(item)
                for it in items:
                    core = [x for x in it if x.lower() not in ("asc", "desc")]
                    if not (len(core) == 3 and re.match(r"[A-Za-z]\w*$", core[0]) and core[1] == "." and core[2] == "[x]"):
                        probs.append(f"{w.upper()} BY には「別名.[列]」だけを書く決まりです（式や SELECT 句の別名は ACE で「パラメーターが少なすぎます」になることがある）: " + " ".join(it))
    return probs


def on_ok(cond):
    # 外側のかっこは許す
    c = list(cond)
    while c and c[0] == "(" and c[-1] == ")":
        c = c[1:-1]
    parts = []
    cur = []
    for t in c:
        if t.lower() == "and":
            parts.append(cur)
            cur = []
        else:
            cur.append(t)
    parts.append(cur)
    for p in parts:
        p = [x for x in p if x not in ("(", ")")]
        if not (len(p) == 7 and re.match(r"[A-Za-z]\w*$", p[0]) and p[1] == "." and p[2] == "[x]" and p[3] == "="
                and re.match(r"[A-Za-z]\w*$", p[4]) and p[5] == "." and p[6] == "[x]"):
            return False
    return True
