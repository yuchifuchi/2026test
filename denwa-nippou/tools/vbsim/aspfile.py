"""ASP ファイルの前処理: #include の展開と <% %> の切り分け。

本物の ASP と同じく、#include は字面の差し込みであり、<% と %> は VBScript の
文字列の中にあっても区切りとして扱う（だから "%>" を文字列に書くと壊れる。それも再現する）。
"""
import os
import re

INCLUDE_RE = re.compile(r'<!--\s*#include\s+(file|virtual)\s*=\s*"([^"]*)"\s*-->', re.I)


class AspSourceError(Exception):
    def __init__(self, file, line, msg):
        super().__init__(f"{file}:{line}: {msg}")
        self.file, self.line, self.msg = file, line, msg


class Expanded:
    def __init__(self):
        self.text = []        # 文字
        self.origin = []      # 各文字の (ファイル, 行) 番号
        self.places = []      # (ファイル, 行) の表
        self.includes = []    # (取り込んだ側, 行, 取り込まれた側)


def _read(path):
    raw = open(path, "rb").read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raise AspSourceError(path, 1, "ファイルの先頭に BOM があります（<%@ より前に出力が起きて落ちます）")
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as e:
        raise AspSourceError(path, 1, f"UTF-8 として読めません: {e}")
    return text.replace("\r\n", "\n")


def expand(path, rel_base=None):
    """#include を展開した全文と、各文字の出どころを返す。"""
    ex = Expanded()
    rel_base = rel_base or os.path.dirname(path)

    def disp(p):
        return os.path.relpath(p, rel_base).replace(os.sep, "/")

    def add(path, stack):
        text = _read(path)
        line = 1
        pos = 0
        for m in INCLUDE_RE.finditer(text):
            seg = text[pos:m.start()]
            line = _emit(ex, seg, disp(path), line)
            kind, target = m.groups()
            if kind.lower() == "virtual":
                raise AspSourceError(disp(path), line, f'#include virtual="{target}" は使わない決まりです（置き場所に縛られるため file= を使う）')
            if ".." in target.replace("\\", "/").split("/"):
                raise AspSourceError(disp(path), line, f'#include file="{target}" に .. があります（IIS の既定では親フォルダ参照は禁止）')
            inc = os.path.normpath(os.path.join(os.path.dirname(path), target.replace("\\", "/")))
            if not os.path.isfile(inc):
                raise AspSourceError(disp(path), line, f'#include の取り込み先がありません: {target}')
            if inc in stack:
                raise AspSourceError(disp(path), line, f'#include が循環しています: {target}')
            ex.includes.append((disp(path), line, disp(inc)))
            add(inc, stack + [inc])
            line += m.group(0).count("\n")
            pos = m.end()
        _emit(ex, text[pos:], disp(path), line)

    add(os.path.abspath(path), [os.path.abspath(path)])
    return ex


def _emit(ex, seg, file, line):
    for ch in seg:
        key = (file, line)
        if not ex.places or ex.places[-1] != key:
            ex.places.append(key)
        ex.text.append(ch)
        ex.origin.append(len(ex.places) - 1)
        if ch == "\n":
            line += 1
    return line


class Program:
    """切り分けた結果。code は VBScript のソース、linemap[i] は code の i 行目の出どころ。"""

    def __init__(self):
        self.directive = {}
        self.blocks = []      # 地の文（HTML）
        self.code_lines = []
        self.linemap = []
        self.includes = []

    @property
    def code(self):
        return "\n".join(self.code_lines) + "\n"


def split(ex):
    text = "".join(ex.text)
    prog = Program()
    prog.includes = ex.includes

    def place(i):
        if i >= len(ex.origin):
            return ex.places[-1] if ex.places else ("?", 0)
        return ex.places[ex.origin[i]]

    def add_lines(s, start):
        # s を行に分けて、それぞれの出どころを記録する
        off = start
        for ln in s.split("\n"):
            prog.code_lines.append(ln)
            prog.linemap.append(place(off))
            off += len(ln) + 1

    i, n = 0, len(text)
    first = True
    while i < n:
        j = text.find("<%", i)
        lit = text[i:] if j < 0 else text[i:j]
        if "%>" in lit:
            k = text.find("%>", i)
            f, l = place(k)
            raise AspSourceError(f, l, "%> に対応する <% がありません")
        if lit:
            prog.blocks.append(lit)
            prog.code_lines.append(f"__WB__ {len(prog.blocks) - 1}")
            prog.linemap.append(place(i))
        if j < 0:
            break
        k = text.find("%>", j + 2)
        if k < 0:
            f, l = place(j)
            raise AspSourceError(f, l, "<% が閉じていません（%> がありません）")
        inner = text[j + 2:k]
        if inner.startswith("@"):
            if not first or i != 0 or lit:
                f, l = place(j)
                raise AspSourceError(f, l, "<%@ ... %> はページの一番最初に置く必要があります")
            for m in re.finditer(r'(\w+)\s*=\s*"([^"]*)"', inner):
                prog.directive[m.group(1).upper()] = m.group(2)
            # 指示の直後の改行が出力になるかどうかは確かめていないので、ページは
            # 「<%@ ... %><%」と改行を挟まずに書く決まりにしている（lint で検査）。
            i = k + 2
            first = False
            continue
        if inner.startswith("="):
            expr = inner[1:]
            if "\n" in expr.strip():
                f, l = place(j)
                raise AspSourceError(f, l, "<%= %> の中で改行しないでください")
            prog.code_lines.append("Response.Write " + expr.strip())
            prog.linemap.append(place(j))
        else:
            add_lines(inner, j + 2)
        i = k + 2
        first = False
    return prog


def load_page(path, rel_base=None):
    return split(expand(path, rel_base))
