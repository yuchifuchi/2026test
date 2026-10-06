"""VBScript の字句解析。"""
import re

from .values import VDate
import datetime


class Tok:
    __slots__ = ("kind", "val", "line", "low")

    def __init__(self, kind, val, line):
        self.kind = kind      # id num str date op nl eof wb
        self.val = val
        self.line = line      # 生成コードの行番号（0 始まり）
        self.low = val.lower() if kind == "id" else None

    def __repr__(self):
        return f"{self.kind}:{self.val!r}@{self.line}"


class LexError(Exception):
    def __init__(self, line, msg):
        super().__init__(msg)
        self.line, self.msg = line, msg


OPS3 = ("<=", ">=", "<>")
OPS1 = "=<>+-*/\\^&(),.:"


def _date_lit(s, line):
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return VDate(datetime.datetime.strptime(s, fmt))
        except ValueError:
            pass
    raise LexError(line, f"日付の定数が読めません: #{s}#")


def lex(code):
    toks = []
    i, n, line = 0, len(code), 0

    def prev_sig():
        return toks[-1] if toks else None

    while i < n:
        c = code[i]
        if c in " \t":
            i += 1
            continue
        if c == "\n":
            toks.append(Tok("nl", "\n", line))
            line += 1
            i += 1
            continue
        if c == "_":
            # 行の継続: 空白 + _ + 改行
            j = i + 1
            while j < n and code[j] in " \t":
                j += 1
            if (j >= n or code[j] == "\n") and (i == 0 or code[i - 1] in " \t"):
                i = j + 1
                line += 1
                continue
        if c == "'":
            while i < n and code[i] != "\n":
                i += 1
            continue
        if c == '"':
            j = i + 1
            buf = []
            while True:
                if j >= n or code[j] == "\n":
                    raise LexError(line, "文字列定数が終了していません")
                if code[j] == '"':
                    if j + 1 < n and code[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(code[j])
                j += 1
            toks.append(Tok("str", "".join(buf), line))
            i = j + 1
            continue
        if c == "#":
            j = code.find("#", i + 1)
            if j < 0 or "\n" in code[i:j]:
                raise LexError(line, "日付の定数 # が閉じていません")
            toks.append(Tok("date", _date_lit(code[i + 1:j], line), line))
            i = j + 1
            continue
        if c == "&" and i + 1 < n and code[i + 1] in "hH":
            m = re.match(r"&[hH]([0-9A-Fa-f]+)&?", code[i:])
            if m:
                v = int(m.group(1), 16)
                # &H8000 などは VBScript では負の Integer / Long になる
                h = m.group(1)
                if len(h) <= 4 and not m.group(0).endswith("&"):
                    if v >= 0x8000:
                        v -= 0x10000
                elif v >= 0x80000000:
                    v -= 0x100000000
                toks.append(Tok("num", v, line))
                i += m.end()
                continue
        if c.isdigit() or (c == "." and i + 1 < n and code[i + 1].isdigit()
                           and not (prev_sig() and (prev_sig().kind == "id" or prev_sig().val == ")"))):
            m = re.match(r"\d*\.?\d+(?:[eE][+-]?\d+)?|\d+\.", code[i:])
            s = m.group(0)
            v = float(s) if any(ch in s for ch in ".eE") else int(s)
            toks.append(Tok("num", v, line))
            i += len(s)
            continue
        if c.isalpha() and c.isascii():
            m = re.match(r"[A-Za-z][A-Za-z0-9_]*", code[i:])
            s = m.group(0)
            low = s.lower()
            if low == "rem":
                p = prev_sig()
                if p is None or p.kind == "nl" or p.val == ":":
                    while i < n and code[i] != "\n":
                        i += 1
                    continue
            if s == "__WB__" or s == "__wb__":
                pass
            toks.append(Tok("id", s, line))
            i += len(s)
            continue
        if c == "_" and i + 5 < n and code.startswith("__WB__", i):
            m = re.match(r"__WB__ (\d+)", code[i:])
            toks.append(Tok("wb", int(m.group(1)), line))
            i += m.end()
            continue
        if c == "[":
            j = code.find("]", i)
            if j < 0:
                raise LexError(line, "[ が閉じていません")
            toks.append(Tok("id", code[i + 1:j], line))
            i = j + 1
            continue
        two = code[i:i + 2]
        if two in OPS3:
            toks.append(Tok("op", two, line))
            i += 2
            continue
        if c in OPS1:
            toks.append(Tok("op", c, line))
            i += 1
            continue
        raise LexError(line, f"使えない文字があります: {c!r}（全角の記号や空白がコードに混ざっていないか確認）")
    toks.append(Tok("nl", "\n", line))
    toks.append(Tok("eof", "", line))
    return toks
