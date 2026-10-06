"""ASP の静的検査。実行しなくても分かる誤りを、納品前に機械的に見つける。

検査すること（手順書 9-3）:
  - <% %> の対応、#include 先の実在、.. や virtual= を使っていないこと
  - If / For / Do / Select / Function の対応（構文解析で）
  - 定義されていない関数・変数の呼び出し、引数の数の違い、Sub を式で使う誤り
  - 配列の添字あふれ（Dim a(4) と For i = 0 To 7 のような組み合わせ）
  - Response.End 系（Response.End / Redirect / それを呼ぶ自作の手続き）が
    On Error Resume Next の生きている所で呼ばれていないか（7-2）
  - On Error Resume Next を開きっぱなしにしていないか
  - SQL を文字列連結で組み立てていないか（7-8）
  - ページの先頭の決まり（<%@ ... CODEPAGE="65001" %>、Option Explicit、CharSet / CodePage）
"""
import os
import re

from . import aspfile
from . import parser as P
from .builtins import BUILTIN_NAMES
from .interp import Interp, Frame
from .values import VBError

PAGE_HEAD = '<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%\n'
INTRINSICS = {"response", "request", "server", "err"}
FORBIDDEN_NAMES = {"session": "Session は使わない決まりです（7-4）", "application": "Application は使わない決まりです"}
FORBIDDEN_BUILTINS = {"formatdatetime", "formatnumber", "weekdayname", "monthname", "eval", "getref", "isdate",
                      "getobject", "execute", "executeglobal"}
FATAL_MEMBERS = {("response", "end"), ("response", "redirect"), ("server", "transfer"), ("server", "execute")}
ORN_MAX_STMTS = 12


class Finding:
    def __init__(self, where, msg):
        self.where, self.msg = where, msg

    def __str__(self):
        return f"{self.where[0]}:{self.where[1]}: {self.msg}"


class DummyHost:
    def intrinsic(self, low):
        return None


class Linter:
    def __init__(self, page_path, rel_base, is_page=True):
        self.path = page_path
        self.rel_base = rel_base
        self.is_page = is_page
        self.f = []

    def err(self, line, msg):
        if self.prog.linemap and line >= len(self.prog.linemap):
            line = len(self.prog.linemap) - 1
        where = self.prog.linemap[line] if 0 <= line < len(self.prog.linemap) else ("?", 0)
        self.f.append(Finding(where, msg))

    def run(self):
        disp = os.path.relpath(self.path, self.rel_base)
        raw = open(self.path, "rb").read()
        if self.is_page and not raw.decode("utf-8", "replace").replace("\r\n", "\n").startswith(PAGE_HEAD):
            self.f.append(Finding((disp, 1), 'ページの 1 行目は <%@ LANGUAGE="VBScript" CODEPAGE="65001" %><% で始め、改行してから Option Explicit を書く決まりです'))
        try:
            self.prog = aspfile.load_page(self.path, self.rel_base)
        except aspfile.AspSourceError as e:
            self.f.append(Finding((e.file, e.line), e.msg))
            return self.f
        try:
            self.ast = P.parse(self.prog.code)
        except P.ParseError as e:
            self.err(e.line, "構文の誤り: " + e.msg)
            return self.f
        # Interp の準備段階（名前の重複・Const の右辺・Dim の大きさ・Option Explicit の位置）
        it = Interp(self.prog, self.ast, DummyHost(), None)
        try:
            it.prepare()
        except VBError as e:
            self.f.append(Finding(e.where or ("?", 0), e.description))
            return self.f
        self.procs = it.procs
        self.gvars = set(it.g.vars)
        self.gconsts = set(it.g.consts)
        self.garrays = {k: c.v.ubs for k, c in it.g.vars.items() if getattr(c.v, "fixed", False)}
        self.gconstvals = {k: v for k, v in it.g.consts.items() if isinstance(v, int) and not isinstance(v, bool)}
        self.check_head()
        self.check_includes_are_declarations()
        if not self.is_page:
            # include を単独で読むときは、構造（構文・宣言だけか）だけを見る。
            # 名前の解決は、include を順に重ねたページ（probe2～7 と各画面）の検査で行う。
            return self.f
        self.check_names()
        self.check_arrays()
        self.check_fatal_under_orn()
        self.check_sql_safety()
        return self.f

    # ---------- ページの先頭 ----------
    def check_head(self):
        if not self.is_page:
            return
        body = [s for s in self.ast if not isinstance(s, P.FuncDef)]
        if not body or not isinstance(body[0], P.OptionExplicit):
            self.err(body[0].line if body else 0, "Option Explicit がページの最初の文になっていません")
            return
        want = {"charset": False, "codepage": False}
        for s in body[1:4]:
            if isinstance(s, P.Assign) and isinstance(s.target, P.Member) and isinstance(s.target.obj, P.Name) \
                    and s.target.obj.low == "response" and isinstance(s.expr, P.Lit):
                if s.target.low == "charset" and s.expr.v == "utf-8":
                    want["charset"] = True
                if s.target.low == "codepage" and s.expr.v == 65001:
                    want["codepage"] = True
        if not all(want.values()):
            self.err(body[0].line, 'Option Explicit のすぐ後に Response.CharSet = "utf-8" : Response.CodePage = 65001 を書く決まりです（Session.CodePage に頼らない。7-4）')

    def check_includes_are_declarations(self):
        """include ファイルは宣言（Const / Dim / Function / Sub）だけにする。実行順の取り違えを防ぐため。"""
        page_disp = os.path.relpath(self.path, self.rel_base).replace(os.sep, "/")
        for s in self.ast:
            where = self.prog.linemap[s.line]
            if where[0] == page_disp and self.is_page:
                continue
            if isinstance(s, P.WriteBlock):
                text = self.prog.blocks[s.idx]
                if text.strip():
                    self.f.append(Finding(where, "include ファイルに <% %> の外の文字があります"))
                continue
            if not isinstance(s, (P.FuncDef, P.Const, P.Dim)):
                self.f.append(Finding(where, "include ファイルの一番外側には Const / Dim / Function / Sub だけを書く決まりです"))

    # ---------- 名前 ----------
    def scope_names(self, proc):
        names = set()
        consts = set()
        arrays = {}
        if proc is None:
            return names, consts, arrays
        for pn, _, _ in proc.params:
            names.add(pn.lower())
        if proc.kind == "function":
            names.add(proc.low)
        fr = Frame()
        it = Interp(self.prog, [], DummyHost(), None)
        try:
            it._hoist(proc.body, fr)
        except VBError as e:
            self.f.append(Finding(e.where or ("?", 0), e.description))
        for pn in names:
            if pn in fr.vars and pn != proc.low:
                self.err(proc.line, f"引数 '{pn}' と同じ名前を Dim しています")
        names |= set(fr.vars)
        consts |= set(fr.consts)
        arrays = {k: c.v.ubs for k, c in fr.vars.items() if getattr(c.v, "fixed", False)}
        self.lconstvals = {k: v for k, v in fr.consts.items() if isinstance(v, int) and not isinstance(v, bool)}
        return names, consts, arrays

    def scopes(self):
        yield None, [s for s in self.ast if not isinstance(s, P.FuncDef)]
        for p in self.procs.values():
            yield p, p.body

    def check_names(self):
        for proc, body in self.scopes():
            lnames, lconsts, _ = self.scope_names(proc)
            known_vars = lnames | lconsts | self.gvars | self.gconsts

            def resolve(name, line, ctx):
                low = name.lower()
                if low in known_vars:
                    return "var"
                if low in self.procs:
                    return "proc"
                if low in FORBIDDEN_NAMES:
                    self.err(line, FORBIDDEN_NAMES[low])
                    return "bad"
                if low in BUILTIN_NAMES:
                    if low in FORBIDDEN_BUILTINS:
                        self.err(line, f"{name} は使わない決まりです（地域設定で結果が変わる・または危険）")
                    return "builtin"
                if low in INTRINSICS:
                    return "intrinsic"
                self.err(line, f"定義されていない名前 '{name}' を{ctx}（綴りの誤りか、Dim / Function が無い）")
                return "bad"

            def check_expr(e, ctx="使っています"):
                callees = {id(x.callee) for x in P.walk_expr(e) if isinstance(x, P.Call)}
                for x in P.walk_expr(e):
                    if isinstance(x, P.Name) and id(x) in callees:
                        continue
                    if isinstance(x, P.Call) and isinstance(x.callee, P.Name):
                        kind = resolve(x.callee.name, x.line, "呼んでいます")
                        if kind == "proc":
                            p = self.procs[x.callee.low]
                            if p.kind == "sub":
                                self.err(x.line, f"Sub '{p.name}' を式の中で使っています（値を返すなら Function にする）")
                            elif len(p.params) != len(x.args) and not (proc and proc.low == p.low and False):
                                self.err(x.line, f"'{p.name}' の引数の数が違います（{len(p.params)} 個のところ {len(x.args)} 個）")
                    elif isinstance(x, P.Name):
                        kind = resolve(x.name, x.line, ctx)
                        if kind == "proc":
                            p = self.procs[x.low]
                            if p.kind == "sub":
                                self.err(x.line, f"Sub '{p.name}' を値として使っています")
                            elif p.params and not (proc and proc.low == x.low):
                                self.err(x.line, f"'{p.name}' を引数なしで使っています（{len(p.params)} 個必要）")

            stmts = P.walk_stmts(body) if proc else self._walk_top(body)
            for s in stmts:
                if isinstance(s, P.CallStmt):
                    c = s.callee
                    if isinstance(c, P.Name):
                        kind = resolve(c.name, s.line, "呼んでいます")
                        if kind == "proc":
                            p = self.procs[c.low]
                            if len(p.params) != len(s.args):
                                self.err(s.line, f"'{p.name}' の引数の数が違います（{len(p.params)} 個のところ {len(s.args)} 個）")
                        elif kind == "var":
                            self.err(s.line, f"変数 '{c.name}' を Sub のように呼んでいます")
                    else:
                        check_expr(c)
                    for a in s.args:
                        check_expr(a)
                elif isinstance(s, P.Assign):
                    t = s.target
                    if isinstance(t, P.Name):
                        low = t.low
                        if low in lconsts or low in self.gconsts:
                            self.err(s.line, f"Const '{t.name}' に代入しています")
                        elif not (low in known_vars or (proc and low == proc.low)):
                            resolve(t.name, s.line, "代入しています")
                    else:
                        check_expr(t)
                    check_expr(s.expr)
                elif isinstance(s, (P.For, P.ForEach)):
                    resolve(s.var, s.line, "For の変数にしています")
                    for e in P.stmt_exprs(s):
                        check_expr(e)
                elif isinstance(s, (P.ReDim,)):
                    for nm, dims in s.items:
                        resolve(nm, s.line, "ReDim しています")
                        for d in dims:
                            check_expr(d)
                elif isinstance(s, P.Erase):
                    resolve(s.name, s.line, "Erase しています")
                elif isinstance(s, (P.Dim, P.Const)):
                    continue
                else:
                    for e in P.stmt_exprs(s):
                        check_expr(e)

    def _walk_top(self, body):
        yield from P.walk_stmts_local(body)

    # ---------- 配列の添字 ----------
    def check_arrays(self):
        for proc, body in self.scopes():
            self.lconstvals = {}
            _, lconsts, larrays = self.scope_names(proc)
            arrays = dict(self.garrays)
            arrays.update(larrays)
            if proc:
                for pn, _, _ in proc.params:
                    arrays.pop(pn.lower(), None)
            self._arr_block(body, arrays, {})

    def _const_int(self, e, loops, arrays):
        if isinstance(e, P.Lit) and isinstance(e.v, int) and not isinstance(e.v, bool):
            return (e.v, e.v)
        if isinstance(e, P.Paren):
            return self._const_int(e.e, loops, arrays)
        if isinstance(e, P.Name) and e.low in loops:
            return loops[e.low]
        if isinstance(e, P.Name) and e.low in getattr(self, "lconstvals", {}):
            v = self.lconstvals[e.low]
            return (v, v)
        if isinstance(e, P.Name) and e.low in self.gconstvals:
            v = self.gconstvals[e.low]
            return (v, v)
        if isinstance(e, P.Unary) and e.op == "-":
            r = self._const_int(e.e, loops, arrays)
            return (-r[1], -r[0]) if r else None
        if isinstance(e, P.Bin) and e.op in ("+", "-"):
            a = self._const_int(e.l, loops, arrays)
            b = self._const_int(e.r, loops, arrays)
            if a and b:
                return (a[0] + b[0], a[1] + b[1]) if e.op == "+" else (a[0] - b[1], a[1] - b[0])
        if isinstance(e, P.Call) and isinstance(e.callee, P.Name) and e.callee.low == "ubound" and e.args:
            a = e.args[0]
            if isinstance(a, P.Name) and a.low in arrays and len(e.args) == 1:
                u = arrays[a.low][0]
                return (u, u)
        return None

    def _arr_block(self, body, arrays, loops):
        for s in body:
            if isinstance(s, P.FuncDef):
                continue
            for e in P.stmt_exprs(s) if not isinstance(s, (P.Dim, P.ReDim)) else []:
                self._arr_expr(e, arrays, loops, s.line)
            if isinstance(s, P.ReDim):
                for nm, _ in s.items:
                    arrays.pop(nm.lower(), None)
            if isinstance(s, P.For):
                a = self._const_int(s.start, loops, arrays)
                b = self._const_int(s.end, loops, arrays)
                inner = dict(loops)
                step_neg = isinstance(s.step, P.Unary) or (isinstance(s.step, P.Lit) and isinstance(s.step.v, int) and s.step.v < 0)
                if a and b:
                    inner[s.var.lower()] = (min(a[0], b[0]), max(a[1], b[1])) if not step_neg else (min(b[0], a[0]), max(a[1], b[1]))
                else:
                    inner.pop(s.var.lower(), None)
                self._arr_block(s.body, arrays, inner)
            elif isinstance(s, P.If):
                for _, b in s.branches:
                    self._arr_block(b, arrays, loops)
                if s.else_block:
                    self._arr_block(s.else_block, arrays, loops)
            elif isinstance(s, (P.ForEach, P.Do)):
                self._arr_block(s.body, arrays, loops)
            elif isinstance(s, P.Select):
                for _, b in s.cases:
                    self._arr_block(b, arrays, loops)
                if s.else_block:
                    self._arr_block(s.else_block, arrays, loops)

    def _arr_expr(self, e, arrays, loops, line):
        for x in P.walk_expr(e):
            if isinstance(x, P.Call) and isinstance(x.callee, P.Name) and x.callee.low in arrays:
                ubs = arrays[x.callee.low]
                if len(x.args) != len(ubs):
                    self.err(line, f"配列 '{x.callee.name}' は {len(ubs)} 次元なのに添字が {len(x.args)} 個です")
                    continue
                for i, (a, u) in enumerate(zip(x.args, ubs)):
                    r = self._const_int(a, loops, arrays)
                    if r and (r[0] < 0 or r[1] > u):
                        self.err(line, f"配列 '{x.callee.name}' の添字が範囲を超えます（宣言は 0～{u}、ここでは {r[0]}～{r[1]}。Dim a({u}) は要素 {u + 1} 個）")

    # ---------- On Error Resume Next と Response.End 系 ----------
    def fatal_procs(self):
        fatal = set()

        def direct(e):
            for x in P.walk_expr(e):
                if isinstance(x, P.Member) and isinstance(x.obj, P.Name) and (x.obj.low, x.low) in FATAL_MEMBERS:
                    return True
                if isinstance(x, P.Call) and isinstance(x.callee, P.Name) and x.callee.low in fatal:
                    return True
                if isinstance(x, P.Name) and x.low in fatal:
                    return True
            return False

        self._direct = direct
        changed = True
        while changed:
            changed = False
            for p in self.procs.values():
                if p.low in fatal:
                    continue
                for s in P.walk_stmts(p.body):
                    if any(direct(e) for e in P.stmt_exprs(s)):
                        fatal.add(p.low)
                        changed = True
                        break
        return fatal

    def check_fatal_under_orn(self):
        fatal = self.fatal_procs()
        self.fatal = fatal
        for proc, body in self.scopes():
            orn = False
            orn_line = None
            count = 0
            seq = list(P.walk_stmts(body)) if proc else list(self._walk_top(body))
            for s in seq:
                if isinstance(s, P.OnError):
                    if s.resume_next:
                        if orn:
                            self.err(s.line, "On Error Resume Next が二重になっています")
                        orn, orn_line, count = True, s.line, 0
                    else:
                        orn = False
                    continue
                if orn:
                    count += 1
                    if count > ORN_MAX_STMTS:
                        self.err(orn_line, f"On Error Resume Next の範囲が長すぎます（{ORN_MAX_STMTS} 文以内で On Error GoTo 0 に戻す決まり）")
                        orn = False
                        continue
                    if any(self._direct(e) for e in P.stmt_exprs(s)):
                        self.err(s.line, "On Error Resume Next が生きている所で Response.End 系（ページを止める処理）を呼んでいます。"
                                         "Response.End は捕まえられる例外なので、ここではページが止まりません（7-2）。先に On Error GoTo 0 に戻してください")
            if orn:
                self.err(orn_line, "On Error Resume Next を On Error GoTo 0 で戻していません")

    # ---------- SQL ----------
    def check_sql_safety(self):
        sql_takers = {p.low for p in self.procs.values() if p.params and p.params[0][0].lower() == "sql"}
        sql_funcs = {p.low for p in self.procs.values() if p.low.startswith("sql")}

        def safe(e, assigns, seen=frozenset()):
            if isinstance(e, P.Lit) and isinstance(e.v, str):
                return True
            if isinstance(e, P.Paren):
                return safe(e.e, assigns, seen)
            if isinstance(e, P.Bin) and e.op == "&":
                return safe(e.l, assigns, seen) and safe(e.r, assigns, seen)
            if isinstance(e, P.Name):
                if e.low in seen:
                    # s = s & "..." のような自分自身への足し込み（ほかの代入が安全なら安全）
                    return True
                if e.low in assigns:
                    # 変数なら、その手続きの中での代入がすべて安全なときだけ安全
                    return all(safe(x, assigns, seen | {e.low}) for x in assigns[e.low])
                if e.low in sql_funcs or e.low in self.gconsts:
                    return True
                # db.asp の関数の引数 sql はそのまま受け渡すだけなので安全（呼び出し側で検査済み）
                return e.low == "sql"
            if isinstance(e, P.Call) and isinstance(e.callee, P.Name) and e.callee.low in sql_funcs:
                return not e.args
            return False

        for proc, body in self.scopes():
            assigns = {}
            stmts = list(P.walk_stmts(body)) if proc else list(self._walk_top(body))
            for s in stmts:
                if isinstance(s, P.Assign) and isinstance(s.target, P.Name):
                    assigns.setdefault(s.target.low, []).append(s.expr)
            if proc and proc.low in sql_funcs:
                for x in assigns.get(proc.low, []):
                    if not safe(x, assigns, frozenset({proc.low})):
                        self.err(proc.line, f"{proc.name} の SQL に、定数と他の Sql 関数以外のものが混ざっています（値は必ず ? で渡す。7-8）")
            for s in stmts:
                for e in P.stmt_exprs(s):
                    for x in P.walk_expr(e):
                        if isinstance(x, P.Call) and isinstance(x.callee, P.Name) and x.callee.low in sql_takers and x.args:
                            if not safe(x.args[0], assigns):
                                self.err(s.line, f"{x.callee.name} に渡す SQL が文字列連結で組み立てられています（値は ? と引数の配列で渡す。7-8）")
                if isinstance(s, P.CallStmt) and isinstance(s.callee, P.Name) and s.callee.low in sql_takers and s.args:
                    if not safe(s.args[0], assigns):
                        self.err(s.line, f"{s.callee.name} に渡す SQL が文字列連結で組み立てられています（7-8）")
            # ADO を直接さわるのは db.asp と setup_check.asp だけ
            for s in stmts:
                where = self.prog.linemap[s.line][0]
                allowed = where.endswith("include/db.asp") or where.endswith("setup_check.asp")
                for e in P.stmt_exprs(s):
                    for x in P.walk_expr(e):
                        if isinstance(x, P.Member) and x.low in ("commandtext", "execute", "begintrans", "committrans", "rollbacktrans") and not allowed:
                            if x.low == "execute" and isinstance(x.obj, P.Name) and x.obj.low == "server":
                                continue
                            self.err(s.line, f".{x.name} を直接使っています（データベースは include/db.asp の関数を通す決まり）")


def lint_page(path, rel_base, is_page=True):
    return Linter(path, rel_base, is_page).run()
