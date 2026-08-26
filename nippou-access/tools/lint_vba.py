# -*- coding: utf-8 -*-
"""VBA モジュールの静的チェック。

Windows で貼り付ける前に、Linux 側で拾える種類の間違いだけを潰しておく。
  - 論理行の引用符の対応
  - 行継続 (_) の連続数 (VBA の上限 24)
  - 論理行の長さ (VBA の上限 1023 文字)
  - Sub / Function の対応
  - 呼び出しているのに定義が無いプロシージャ (モジュール横断)
"""
import glob
import io
import os
import re
import sys

VBA_MAX_CONT = 24
VBA_MAX_LEN = 1023

# VBA / Access / DAO の組込みで、定義が無くても呼べるもの
BUILTIN = set("""
Abs Array Asc CBool CByte CCur CDate CDbl CInt CLng CSng CStr CVar Choose Chr
Command Cos Date DateAdd DateDiff DatePart DateSerial DateValue Day DDB Dir
DoEvents Environ EOF Error Exp FileLen Filter Fix Format FormatNumber FreeFile
Hex Hour IIf IMEStatus InStr InStrRev Int IPmt IsArray IsDate IsEmpty IsError
IsMissing IsNull IsNumeric IsObject Join LBound LCase Left Len Loc LOF Log LTrim
Mid Minute Month MonthName MsgBox Now Nz Oct Partition Replace RGB Right Rnd
Round RTrim Second Seek Sgn Shell Sin Space Split Sqr StrComp StrConv String
StrReverse Switch Tan Time Timer TimeSerial TimeValue Trim TypeName UBound UCase
Val VarType Weekday WeekdayName Year InputBox CreateObject GetObject
DCount DSum DAvg DMin DMax DLookup DFirst DLast DVar DStDev
CreateForm CreateReport CreateControl CreateReportControl DeleteControl
CurrentDb CurrentProject CurrentUser SysCmd Eval Print Debug Err
DoCmd Application Screen Forms Reports Me RefreshTitleBar
""".split())

# メンバ呼び出し (.Foo) と配列参照は対象外にしたいので、直前の 1 文字も捕まえる
CALL_RE = re.compile(r"(^|[^.\w])([A-Za-z_぀-ヿ一-鿿][\w぀-ヿ一-鿿]*)\s*\(")
LOCAL_RE = re.compile(r"\b(?:Dim|Const|ReDim|ByVal|ByRef)\s+([\w぀-ヿ一-鿿]+)")
BARE_RE = re.compile(r"^\s*(?:Call\s+)?([A-Za-z_぀-ヿ一-鿿][\w぀-ヿ一-鿿]*)\s*$")
DEF_RE = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
                    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
                    r"([\w぀-ヿ一-鿿]+)")
END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)


def logical_lines(path):
    """行継続を連結した論理行を (開始行番号, 継続本数, テキスト) で返す。"""
    out = []
    buf, start, cont = "", None, 0
    for i, raw in enumerate(io.open(path, encoding="utf-8"), 1):
        line = raw.rstrip("\r\n")
        if start is None:
            start = i
        stripped = line.rstrip()
        if stripped.endswith(" _") or stripped == "_":
            buf += stripped[:-1]
            cont += 1
            continue
        buf += stripped
        out.append((start, cont, buf))
        buf, start, cont = "", None, 0
    if buf:
        out.append((start, cont, buf))
    return out


def strip_strings_and_comment(text):
    """文字列リテラルとコメントを除いたコードを返す。引用符が奇数なら None。"""
    res, i, n = [], 0, len(text)
    in_str = False
    while i < n:
        ch = text[i]
        if in_str:
            if ch == '"':
                if i + 1 < n and text[i + 1] == '"':
                    i += 2
                    continue
                in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            i += 1
            continue
        if ch == "'":
            break
        res.append(ch)
        i += 1
    if in_str:
        return None
    return "".join(res)


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    files = sorted(glob.glob(os.path.join(root, "src", "*.bas")))
    problems = 0
    defined, called, locals_seen = set(), {}, set()

    for path in files:
        rel = os.path.relpath(path, root)
        depth = 0
        for start, cont, text in logical_lines(path):
            if cont > VBA_MAX_CONT:
                print("%s:%d  行継続が %d 本 (上限 %d)" % (rel, start, cont, VBA_MAX_CONT))
                problems += 1
            if len(text) > VBA_MAX_LEN:
                print("%s:%d  論理行が %d 文字 (上限 %d)" % (rel, start, len(text), VBA_MAX_LEN))
                problems += 1

            code = strip_strings_and_comment(text)
            if code is None:
                print("%s:%d  引用符が閉じていない: %s" % (rel, start, text.strip()[:80]))
                problems += 1
                continue

            m = DEF_RE.match(code)
            if m:
                defined.add(m.group(2))
                depth += 1
            elif END_RE.match(code):
                depth -= 1

            for name in LOCAL_RE.findall(code):
                locals_seen.add(name)
            for _, name in CALL_RE.findall(code):
                called.setdefault(name, (rel, start))
            bm = BARE_RE.match(code)
            if bm and bm.group(1) not in ("Else", "End", "Exit", "Loop", "Wend", "Stop"):
                called.setdefault(bm.group(1), (rel, start))

        if depth != 0:
            print("%s  Sub/Function の対応が %+d ずれています" % (rel, depth))
            problems += 1

    keywords = set("""If Then Else ElseIf Select Case For Next Do Loop While Wend With
    Set Let Dim Const ReDim Erase On Error Resume GoTo Exit End Sub Function Property
    Public Private Friend Static Optional ByVal ByRef ParamArray As New Not And Or Xor
    Is Like Mod To Step Each In Nothing True False Null Empty Type Enum Declare Lib
    Option Compare Database Explicit Attribute VB_Name Call""".split())
    for name, (rel, line) in sorted(called.items()):
        if name in defined or name in BUILTIN or name in keywords or name in locals_seen:
            continue
        if name.startswith("ac") or name.startswith("db") or name.startswith("vb"):
            continue
        print("%s:%d  定義が見つからない呼び出し: %s" % (rel, line, name))
        problems += 1

    print("---")
    print("モジュール %d 本 / プロシージャ %d 個 / 指摘 %d 件"
          % (len(files), len(defined), problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
