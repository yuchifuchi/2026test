# -*- coding: utf-8 -*-
"""IIS が無い環境で、ASP ページの明らかな壊れを見つける。

見るもの
  ・<% %> の対応
  ・Function / Sub の対応
  ・#include で読んでいるファイルが実在するか
  ・呼んでいる自前の関数が、そのページで読み込まれる範囲に定義されているか
    (RequireStaff を消したのに呼び出しが残っている、等)
  ・Dim で取った配列の大きさを超える添字 (VBScript では実行時エラーになる)

見られないもの (IIS と Access が要る)
  ・SQL が通るか / 実際の画面
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(HERE, "web")

# VBScript の組み込み関数。定義が無くても呼んでよいもの。
BUILTIN = set("""
abs array asc atn cbool cbyte ccur cdate cdbl chr cint clng cos createobject cstr csng
date dateadd datediff datepart dateserial datevalue day escape eval exp filter fix
formatcurrency formatdatetime formatnumber formatpercent getobject getref hex hour
instr instrrev int isarray isdate isempty isnull isnumeric isobject join lbound lcase
left len log ltrim mid minute month monthname now oct replace rgb right rnd round rtrim
second sgn sin space split sqr strcomp string strreverse tan time timer timeserial
timevalue trim typename ubound ucase unescape vartype weekday weekdayname year
""".split())

# 組み込みオブジェクト (Response.Write などの受け側)
OBJECTS = set("response request server session application err".split())

DEF = re.compile(r"^\s*(?:public\s+|private\s+)?(?:function|sub)\s+([^\s(]+)", re.I)
END = re.compile(r"^\s*end\s+(?:function|sub)\b", re.I)
DIMNAME = re.compile(r"([A-Za-z_]\w*)\s*(?:\((\d+)\))?")


def read(p):
    return io.open(p, encoding="utf-8").read()


def strip_strings(line):
    """文字列リテラルと ' コメントを落とす。"""
    out, i, n, ins = [], 0, len(line), False
    while i < n:
        c = line[i]
        if ins:
            if c == '"':
                if i + 1 < n and line[i + 1] == '"':
                    i += 2
                    continue
                ins = False
            i += 1
            continue
        if c == '"':
            ins = True
            out.append(" ")
            i += 1
            continue
        if c == "'":
            break
        out.append(c)
        i += 1
    return "".join(out)


def code_lines(text):
    """<% %> の中だけを、行番号つきで返す。"""
    out = []
    for m in re.finditer(r"<%=?(.*?)%>", text, re.S):
        base = text.count("\n", 0, m.start(1)) + 1
        for k, ln in enumerate(m.group(1).split("\n")):
            out.append((base + k, strip_strings(ln)))
    return out


def include_tree(page, problems, name):
    """page が読み込むファイル一式を、include の順に返す。"""
    files, stack = [], [page]
    while stack:
        f = stack.pop(0)
        if f in files:
            continue
        if not os.path.exists(f):
            problems.append("%s: #include のファイルが無い: %s" % (name, f))
            continue
        files.append(f)
        for inc in re.findall(r'#include\s+file="([^"]+)"', read(f)):
            stack.append(os.path.normpath(os.path.join(os.path.dirname(f), inc)))
    return files


def scan(f):
    """1 ファイルから 定義 / 変数 / 呼び出し / 配列の大きさ を集める。"""
    defs, variables, calls, arrays = {}, set(), [], {}
    fn = None
    for no, ln in code_lines(read(f)):
        if not ln.strip():
            continue
        m = DEF.match(ln)
        if m:
            fn = m.group(1)
            defs[fn.lower()] = no
            variables.add(fn.lower())
            # 引数も変数
            if "(" in ln:
                for a in ln[ln.index("(") + 1:ln.rfind(")")].split(","):
                    a = a.strip().replace("ByVal ", "").replace("ByRef ", "")
                    if a:
                        variables.add(a.lower())
            continue
        if END.match(ln):
            fn = None
            continue
        if re.match(r"\s*(dim|redim|const)\b", ln, re.I):
            body = re.sub(r"^\s*(dim|redim preserve|redim|const)\b", "", ln, flags=re.I)
            for nm, sz in DIMNAME.findall(body):
                variables.add(nm.lower())
                if sz:
                    arrays[(fn, nm)] = int(sz)
            continue
        for nm in re.findall(r"\bSet\s+([A-Za-z_]\w*)\s*=", ln, re.I):
            variables.add(nm.lower())
        for nm in re.findall(r"\bFor\s+Each\s+([A-Za-z_]\w*)\b", ln, re.I):
            variables.add(nm.lower())
        # 呼び出し: 直前が「.」でないもの (メソッド呼び出しは対象外)
        for m2 in re.finditer(r"(?<![.\w])([A-Za-z_]\w*)\s*\(", ln):
            calls.append((no, m2.group(1), fn))
    return defs, variables, calls, arrays


def loop_ranges(f):
    """For i = 0 To N の N を、関数ごとに覚えておく。"""
    out, fn = {}, None
    for no, ln in code_lines(read(f)):
        m = DEF.match(ln)
        if m:
            fn = m.group(1)
        elif END.match(ln):
            fn = None
        m = re.match(r"\s*For\s+([A-Za-z_]\w*)\s*=\s*[-\d]+\s+To\s+(\d+)", ln, re.I)
        if m:
            out[(fn, m.group(1))] = int(m.group(2))
    return out


def check_page(page, problems):
    name = os.path.relpath(page, HERE)
    text = read(page)
    if text.count("<%") != text.count("%>"):
        problems.append("%s: <%% と %%> の数が合わない (%d / %d)"
                        % (name, text.count("<%"), text.count("%>")))

    files = include_tree(page, problems, name)
    defs, variables, calls = {}, set(), []
    for f in files:
        d, v, c, _a = scan(f)
        defs.update(d)
        variables |= v
        calls += [(f, no, nm, fn) for no, nm, fn in c]

    known = set(defs) | BUILTIN | OBJECTS | variables
    for f, no, nm, _fn in calls:
        if nm.lower() not in known:
            problems.append("%s(%d): 定義の無い呼び出し %s()"
                            % (os.path.relpath(f, HERE), no, nm))

    # 配列の添字が大きさを超えていないか
    for f in files:
        _d, _v, _c, arrays = scan(f)
        if not arrays:
            continue
        loops = loop_ranges(f)
        fn = None
        for no, ln in code_lines(read(f)):
            m = DEF.match(ln)
            if m:
                fn = m.group(1)
            elif END.match(ln):
                fn = None
            for nm, idx in re.findall(r"([A-Za-z_]\w*)\s*\(\s*([\w+ ]+?)\s*\)", ln):
                size = arrays.get((fn, nm))
                if size is None:
                    continue
                hi = None
                if idx.isdigit():
                    hi = int(idx)
                else:
                    m2 = re.match(r"^(\w+)(?:\s*\+\s*(\d+))?$", idx)
                    if m2 and (fn, m2.group(1)) in loops:
                        hi = loops[(fn, m2.group(1))] + int(m2.group(2) or 0)
                if hi is not None and hi > size:
                    problems.append(
                        "%s(%d): %s(%s) は最大 %d まで行くが Dim %s(%d)"
                        % (os.path.relpath(f, HERE), no, nm, idx, hi, nm, size))


def main():
    problems = []
    pages = sorted(p for p in os.listdir(WEB) if p.endswith(".asp"))
    for p in pages:
        check_page(os.path.join(WEB, p), problems)
    problems = sorted(set(problems))
    if problems:
        print("見つかった問題:")
        for p in problems:
            print("  !", p)
        return 1
    print("ASP %d ページ: 問題なし" % len(pages))
    return 0


if __name__ == "__main__":
    sys.exit(main())
