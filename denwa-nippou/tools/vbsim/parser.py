"""VBScript の構文解析。使ってよい書き方だけを受け付ける（With / Class / GoTo などは使わない決まり）。"""
from .lexer import lex, Tok, LexError
from .values import EMPTY, NULL, NOTHING


class ParseError(Exception):
    def __init__(self, line, msg):
        super().__init__(msg)
        self.line, self.msg = line, msg


# ---------- 式 ----------
class Node:
    line = 0


class Lit(Node):
    def __init__(self, v, line):
        self.v, self.line = v, line


class Name(Node):
    def __init__(self, name, line):
        self.name, self.low, self.line = name, name.lower(), line


class Member(Node):
    def __init__(self, obj, name, line):
        self.obj, self.name, self.low, self.line = obj, name, name.lower(), line


class Call(Node):
    def __init__(self, callee, args, line):
        self.callee, self.args, self.line = callee, args, line


class Paren(Node):
    def __init__(self, e, line):
        self.e, self.line = e, line


class Unary(Node):
    def __init__(self, op, e, line):
        self.op, self.e, self.line = op, e, line


class Bin(Node):
    def __init__(self, op, l, r, line):
        self.op, self.l, self.r, self.line = op, l, r, line


class NewObj(Node):
    def __init__(self, cls, line):
        self.cls, self.line = cls, line


# ---------- 文 ----------
class Stmt:
    line = 0


class WriteBlock(Stmt):
    def __init__(self, idx, line):
        self.idx, self.line = idx, line


class Dim(Stmt):
    def __init__(self, items, line, kind="dim"):
        self.items, self.line, self.kind = items, line, kind   # items: [(name, dims or None)]


class ReDim(Stmt):
    def __init__(self, preserve, items, line):
        self.preserve, self.items, self.line = preserve, items, line


class Const(Stmt):
    def __init__(self, items, line):
        self.items, self.line = items, line


class Assign(Stmt):
    def __init__(self, target, expr, is_set, line):
        self.target, self.expr, self.is_set, self.line = target, expr, is_set, line


class CallStmt(Stmt):
    def __init__(self, callee, args, explicit, line):
        self.callee, self.args, self.explicit, self.line = callee, args, explicit, line


class If(Stmt):
    def __init__(self, branches, else_block, line, single=False):
        self.branches, self.else_block, self.line, self.single = branches, else_block, line, single


class For(Stmt):
    def __init__(self, var, start, end, step, body, line):
        self.var, self.start, self.end, self.step, self.body, self.line = var, start, end, step, body, line


class ForEach(Stmt):
    def __init__(self, var, expr, body, line):
        self.var, self.expr, self.body, self.line = var, expr, body, line


class Do(Stmt):
    def __init__(self, pre, post, body, line):
        self.pre, self.post, self.body, self.line = pre, post, body, line   # pre/post: (kind, expr) or None


class Select(Stmt):
    def __init__(self, expr, cases, else_block, line):
        self.expr, self.cases, self.else_block, self.line = expr, cases, else_block, line


class Exit(Stmt):
    def __init__(self, kind, line):
        self.kind, self.line = kind, line


class OnError(Stmt):
    def __init__(self, resume_next, line):
        self.resume_next, self.line = resume_next, line


class OptionExplicit(Stmt):
    def __init__(self, line):
        self.line = line


class Erase(Stmt):
    def __init__(self, name, line):
        self.name, self.line = name, line


class Randomize(Stmt):
    def __init__(self, line):
        self.line = line


class FuncDef(Stmt):
    def __init__(self, kind, name, params, body, line, end_line):
        self.kind, self.name, self.low = kind, name, name.lower()
        self.params, self.body, self.line, self.end_line = params, body, line, end_line   # params: [(name, byval, is_array)]


KEYWORDS = set("""and byref byval call case class const dim do each else elseif empty end eqv erase error exit
explicit false for function get goto if imp in is let loop mod new next not nothing null on option or
preserve private property public redim rem resume select set step sub then to true until wend while with xor""".split())

FORBIDDEN = {
    "with": "With は使わない決まりです（読み手が対象を見失うため）",
    "class": "Class は使わない決まりです",
    "property": "Property は使わない決まりです",
    "goto": "GoTo は使わない決まりです",
    "resume": "Resume は On Error Resume Next の形でだけ使えます",
    "execute": "Execute は使わない決まりです",
    "executeglobal": "ExecuteGlobal は使わない決まりです",
}


class Parser:
    def __init__(self, toks):
        self.t = toks
        self.p = 0
        self.func_stack = []

    # ---- 道具 ----
    def peek(self, k=0):
        return self.t[self.p + k]

    def next(self):
        tok = self.t[self.p]
        self.p += 1
        return tok

    def is_kw(self, kw, k=0):
        tok = self.t[self.p + k]
        return tok.kind == "id" and tok.low == kw

    def is_op(self, op, k=0):
        tok = self.t[self.p + k]
        return tok.kind == "op" and tok.val == op

    def expect_kw(self, kw):
        tok = self.next()
        if not (tok.kind == "id" and tok.low == kw):
            raise ParseError(tok.line, f"「{kw}」が必要な所に「{tok.val}」があります")
        return tok

    def expect_op(self, op):
        tok = self.next()
        if not (tok.kind == "op" and tok.val == op):
            raise ParseError(tok.line, f"「{op}」が必要な所に「{tok.val}」があります")
        return tok

    def ident(self):
        tok = self.next()
        if tok.kind != "id" or tok.low in KEYWORDS:
            raise ParseError(tok.line, f"名前が必要な所に「{tok.val}」があります")
        return tok

    def at_eos(self):
        tok = self.peek()
        return tok.kind in ("nl", "eof") or (tok.kind == "op" and tok.val == ":")

    def end_stmt(self):
        tok = self.peek()
        if tok.kind in ("nl", "eof") or (tok.kind == "op" and tok.val == ":"):
            if tok.kind != "eof":
                self.p += 1
            return
        raise ParseError(tok.line, f"文の終わりが必要な所に「{tok.val}」があります")

    def skip_seps(self):
        while self.peek().kind == "nl" or self.is_op(":"):
            self.p += 1

    # ---- 全体 ----
    def program(self):
        body = self.block(top=True)
        if self.peek().kind != "eof":
            tok = self.peek()
            raise ParseError(tok.line, f"対応する始まりの無い「{tok.val}」があります")
        return body

    def block(self, terms=(), top=False):
        """terms: 終わりの印（('end','if') や ('next',) など）。見つけたら消費せずに返る。"""
        out = []
        while True:
            self.skip_seps()
            tok = self.peek()
            if tok.kind == "eof":
                if terms:
                    want = " ".join(w.capitalize() for w in terms[-1])
                    raise ParseError(tok.line, f"「{want}」がありません（ブロックが閉じていません）")
                return out
            for term in terms:
                if all(self.is_kw(w, k) for k, w in enumerate(term)):
                    return out
            if not top and tok.kind == "id" and tok.low in ("function", "sub") or \
               (tok.kind == "id" and tok.low in ("public", "private") and self.peek(1).kind == "id" and self.peek(1).low in ("function", "sub") and not top):
                raise ParseError(tok.line, "Function / Sub の中に Function / Sub は書けません")
            self.check_stray(tok, terms)
            out.append(self.statement())

    def check_stray(self, tok, terms):
        if tok.kind != "id":
            return
        stray = {"end": "End", "next": "Next", "loop": "Loop", "wend": "Wend", "else": "Else",
                 "elseif": "ElseIf", "case": "Case"}
        if tok.low in stray:
            want = " / ".join(" ".join(t) for t in terms) if terms else "（なし）"
            nxt = self.peek(1).val if self.peek(1).kind == "id" else ""
            raise ParseError(tok.line, f"「{tok.val} {nxt}」に対応する始まりがありません（ここで待っているのは: {want}）")

    # ---- 文 ----
    def statement(self, single_line=False):
        tok = self.peek()
        line = tok.line
        if tok.kind == "wb":
            self.next()
            self.end_stmt_soft(single_line)
            return WriteBlock(tok.val, line)
        if tok.kind != "id":
            raise ParseError(line, f"文の始まりが読めません: 「{tok.val}」")
        k = tok.low
        if k in FORBIDDEN:
            raise ParseError(line, FORBIDDEN[k])
        if k in ("public", "private"):
            if self.peek(1).kind == "id" and self.peek(1).low in ("function", "sub"):
                self.next()
                return self.funcdef()
            self.next()
            return self.dim(line, k, single_line)
        if k == "dim":
            self.next()
            return self.dim(line, "dim", single_line)
        if k == "redim":
            self.next()
            preserve = False
            if self.is_kw("preserve"):
                self.next()
                preserve = True
            items = []
            while True:
                nm = self.ident()
                self.expect_op("(")
                dims = self.arglist_paren()
                items.append((nm.val, dims))
                if self.is_op(","):
                    self.next()
                    continue
                break
            self.end_stmt_soft(single_line)
            return ReDim(preserve, items, line)
        if k == "const":
            self.next()
            items = []
            while True:
                nm = self.ident()
                self.expect_op("=")
                e = self.expr()
                items.append((nm.val, e))
                if self.is_op(","):
                    self.next()
                    continue
                break
            self.end_stmt_soft(single_line)
            return Const(items, line)
        if k == "set":
            self.next()
            target = self.chain()
            self.expect_op("=")
            e = self.expr()
            self.end_stmt_soft(single_line)
            return Assign(target, e, True, line)
        if k == "let":
            self.next()
        if k == "call":
            self.next()
            target = self.chain()
            if isinstance(target, Call):
                st = CallStmt(target.callee, target.args, True, line)
            else:
                st = CallStmt(target, [], True, line)
            self.end_stmt_soft(single_line)
            return st
        if k == "if":
            return self.if_stmt()
        if k == "for":
            return self.for_stmt()
        if k == "do":
            return self.do_stmt()
        if k == "while":
            self.next()
            cond = self.expr()
            self.end_stmt()
            body = self.block((("wend",),))
            self.expect_kw("wend")
            self.end_stmt_soft(single_line)
            return Do(("while", cond), None, body, line)
        if k == "select":
            return self.select_stmt()
        if k == "exit":
            self.next()
            w = self.next()
            if w.kind != "id" or w.low not in ("for", "do", "function", "sub"):
                raise ParseError(line, "Exit の後ろは For / Do / Function / Sub です")
            if w.low in ("function", "sub"):
                if not self.func_stack:
                    raise ParseError(line, f"Exit {w.val} が Function / Sub の外にあります")
                if self.func_stack[-1] != w.low:
                    raise ParseError(line, f"Exit {w.val} を {self.func_stack[-1]} の中で使っています")
            self.end_stmt_soft(single_line)
            return Exit(w.low, line)
        if k in ("function", "sub"):
            return self.funcdef()
        if k == "on":
            self.next()
            self.expect_kw("error")
            if self.is_kw("resume"):
                self.next()
                self.expect_kw("next")
                self.end_stmt_soft(single_line)
                return OnError(True, line)
            if self.is_kw("goto"):
                self.next()
                z = self.next()
                if z.kind != "num" or z.val != 0:
                    raise ParseError(line, "On Error GoTo は 0 だけ使えます")
                self.end_stmt_soft(single_line)
                return OnError(False, line)
            raise ParseError(line, "On Error の書き方が違います")
        if k == "option":
            self.next()
            self.expect_kw("explicit")
            self.end_stmt_soft(single_line)
            return OptionExplicit(line)
        if k == "erase":
            self.next()
            nm = self.ident()
            self.end_stmt_soft(single_line)
            return Erase(nm.val, line)
        if k == "randomize":
            self.next()
            self.end_stmt_soft(single_line)
            return Randomize(line)
        if k in KEYWORDS and k not in ("let",):
            raise ParseError(line, f"ここに「{tok.val}」は書けません")
        return self.expr_stmt(single_line)

    def end_stmt_soft(self, single_line):
        if single_line:
            tok = self.peek()
            if tok.kind in ("nl", "eof") or self.is_op(":") or self.is_kw("else"):
                return
            raise ParseError(tok.line, f"文の終わりが必要な所に「{tok.val}」があります")
        self.end_stmt()

    def dim(self, line, kind, single_line):
        items = []
        while True:
            nm = self.ident()
            dims = None
            if self.is_op("("):
                self.next()
                if self.is_op(")"):
                    self.next()
                    dims = []
                else:
                    dims = self.arglist_paren()
            items.append((nm.val, dims))
            if self.is_op(","):
                self.next()
                continue
            break
        self.end_stmt_soft(single_line)
        return Dim(items, line, kind)

    def funcdef(self):
        tok = self.next()
        kind = tok.low
        line = tok.line
        nm = self.ident()
        params = []
        if self.is_op("("):
            self.next()
            if not self.is_op(")"):
                while True:
                    byval = False
                    if self.is_kw("byval"):
                        self.next()
                        byval = True
                    elif self.is_kw("byref"):
                        self.next()
                    pn = self.ident()
                    is_arr = False
                    if self.is_op("("):
                        self.next()
                        self.expect_op(")")
                        is_arr = True
                    params.append((pn.val, byval, is_arr))
                    if self.is_op(","):
                        self.next()
                        continue
                    break
            self.expect_op(")")
        self.end_stmt()
        self.func_stack.append(kind)
        body = self.block((("end", kind),))
        self.func_stack.pop()
        end_tok = self.expect_kw("end")
        self.next()
        self.end_stmt()
        return FuncDef(kind, nm.val, params, body, line, end_tok.line)

    def if_stmt(self):
        line = self.next().line
        cond = self.expr()
        self.expect_kw("then")
        if not (self.peek().kind in ("nl", "eof")):
            # 1 行の If
            then_body = []
            while True:
                if self.is_op(":"):
                    self.next()
                    continue
                if self.peek().kind in ("nl", "eof") or self.is_kw("else"):
                    break
                then_body.append(self.statement(single_line=True))
            else_body = None
            if self.is_kw("else"):
                self.next()
                else_body = []
                while True:
                    if self.is_op(":"):
                        self.next()
                        continue
                    if self.peek().kind in ("nl", "eof"):
                        break
                    else_body.append(self.statement(single_line=True))
            if self.is_kw("end") and self.is_kw("if", 1):
                raise ParseError(line, "1 行の If に End If を付けています")
            self.end_stmt()
            return If([(cond, then_body)], else_body, line, single=True)
        self.end_stmt()
        branches = []
        body = self.block((("elseif",), ("else",), ("end", "if")))
        branches.append((cond, body))
        else_body = None
        while True:
            if self.is_kw("elseif"):
                self.next()
                c = self.expr()
                self.expect_kw("then")
                self.end_stmt()
                b = self.block((("elseif",), ("else",), ("end", "if")))
                branches.append((c, b))
                continue
            if self.is_kw("else"):
                self.next()
                self.end_stmt()
                else_body = self.block((("end", "if"),))
                continue
            break
        self.expect_kw("end")
        self.expect_kw("if")
        self.end_stmt()
        return If(branches, else_body, line)

    def for_stmt(self):
        line = self.next().line
        if self.is_kw("each"):
            self.next()
            var = self.ident()
            self.expect_kw("in")
            e = self.expr()
            self.end_stmt()
            body = self.block((("next",),))
            self.expect_kw("next")
            if self.peek().kind == "id" and not self.at_eos():
                self.next()
            self.end_stmt()
            return ForEach(var.val, e, body, line)
        var = self.ident()
        self.expect_op("=")
        start = self.expr()
        self.expect_kw("to")
        end = self.expr()
        step = None
        if self.is_kw("step"):
            self.next()
            step = self.expr()
        self.end_stmt()
        body = self.block((("next",),))
        self.expect_kw("next")
        if self.peek().kind == "id" and not self.at_eos():
            nv = self.next()
            if nv.low != var.low:
                raise ParseError(nv.line, f"Next {nv.val} が For {var.val} と合っていません")
        self.end_stmt()
        return For(var.val, start, end, step, body, line)

    def do_stmt(self):
        line = self.next().line
        pre = None
        if self.is_kw("while") or self.is_kw("until"):
            kind = self.next().low
            pre = (kind, self.expr())
        self.end_stmt()
        body = self.block((("loop",),))
        self.expect_kw("loop")
        post = None
        if self.is_kw("while") or self.is_kw("until"):
            kind = self.next().low
            post = (kind, self.expr())
        if pre and post:
            raise ParseError(line, "Do と Loop の両方に条件があります")
        self.end_stmt()
        return Do(pre, post, body, line)

    def select_stmt(self):
        line = self.next().line
        self.expect_kw("case")
        e = self.expr()
        self.end_stmt()
        cases = []
        else_block = None
        self.skip_seps()
        while True:
            self.skip_seps()
            if self.is_kw("end") and self.is_kw("select", 1):
                break
            if not self.is_kw("case"):
                tok = self.peek()
                raise ParseError(tok.line, f"Select Case の中で Case 以外の「{tok.val}」があります")
            self.next()
            if self.is_kw("else"):
                self.next()
                self.end_stmt_case()
                else_block = self.block((("case",), ("end", "select")))
                continue
            vals = [self.expr()]
            while self.is_op(","):
                self.next()
                vals.append(self.expr())
            self.end_stmt_case()
            b = self.block((("case",), ("end", "select")))
            cases.append((vals, b))
        self.expect_kw("end")
        self.expect_kw("select")
        self.end_stmt()
        return Select(e, cases, else_block, line)

    def end_stmt_case(self):
        # Case x: 文 のように同じ行に続けて書いてもよい
        if self.is_op(":"):
            self.next()
            return
        self.end_stmt()

    def expr_stmt(self, single_line):
        line = self.peek().line
        start = self.p
        target, last_paren = self.chain(track=True)
        if self.is_op("="):
            self.next()
            e = self.expr()
            self.end_stmt_soft(single_line)
            return Assign(target, e, False, line)
        if self.at_eos() or (single_line and self.is_kw("else")):
            self.end_stmt_soft(single_line)
            if isinstance(target, Call):
                if len(target.args) > 1:
                    raise ParseError(line, "Sub を呼ぶときにかっこは使えません（引数が 2 つ以上なら Call を付けるか、かっこを外す）")
                args = [Paren(a, a.line) for a in target.args]
                return CallStmt(target.callee, args, False, line)
            return CallStmt(target, [], False, line)
        # 引数がかっこ無しで続く形: Response.Write ("a") & b など
        if last_paren is not None and isinstance(target, Call):
            self.p = last_paren
            callee = target.callee
        else:
            callee = target
        args = [self.expr()]
        while self.is_op(","):
            self.next()
            args.append(self.expr())
        self.end_stmt_soft(single_line)
        return CallStmt(callee, args, False, line)

    def chain(self, track=False):
        tok = self.next()
        if tok.kind != "id":
            raise ParseError(tok.line, f"名前が必要な所に「{tok.val}」があります")
        if tok.low in KEYWORDS and tok.low not in ("error",):
            raise ParseError(tok.line, f"ここに「{tok.val}」は書けません")
        node = Name(tok.val, tok.line)
        last_paren = None
        while True:
            if self.is_op("."):
                self.next()
                m = self.next()
                if m.kind != "id":
                    raise ParseError(m.line, ". の後ろに名前がありません")
                node = Member(node, m.val, m.line)
                last_paren = None
            elif self.is_op("("):
                p0 = self.p
                self.next()
                args = self.arglist_paren()
                node = Call(node, args, tok.line)
                last_paren = p0
            else:
                break
        if track:
            return node, last_paren
        return node

    def arglist_paren(self):
        """「(」の直後から「)」まで。"""
        args = []
        if self.is_op(")"):
            self.next()
            return args
        while True:
            if self.is_op(",") or self.is_op(")"):
                raise ParseError(self.peek().line, "引数を省略する書き方は使わない決まりです")
            args.append(self.expr())
            if self.is_op(","):
                self.next()
                continue
            self.expect_op(")")
            return args

    # ---- 式 ----
    def expr(self):
        return self.imp()

    def binlevel(self, sub, ops, kw=False):
        node = sub()
        while True:
            tok = self.peek()
            if kw and tok.kind == "id" and tok.low in ops:
                op = tok.low
            elif not kw and tok.kind == "op" and tok.val in ops:
                op = tok.val
            else:
                return node
            self.next()
            node = Bin(op, node, sub(), tok.line)

    def imp(self):
        return self.binlevel(self.eqv, ("imp",), True)

    def eqv(self):
        return self.binlevel(self.xor, ("eqv",), True)

    def xor(self):
        return self.binlevel(self.or_, ("xor",), True)

    def or_(self):
        return self.binlevel(self.and_, ("or",), True)

    def and_(self):
        return self.binlevel(self.not_, ("and",), True)

    def not_(self):
        if self.is_kw("not"):
            tok = self.next()
            return Unary("not", self.not_(), tok.line)
        return self.cmp()

    def cmp(self):
        node = self.concat()
        while True:
            tok = self.peek()
            if tok.kind == "op" and tok.val in ("=", "<>", "<", ">", "<=", ">="):
                op = tok.val
            elif tok.kind == "id" and tok.low == "is":
                op = "is"
            else:
                return node
            self.next()
            node = Bin(op, node, self.concat(), tok.line)

    def concat(self):
        return self.binlevel(self.add, ("&",))

    def add(self):
        return self.binlevel(self.mod, ("+", "-"))

    def mod(self):
        return self.binlevel(self.intdiv, ("mod",), True)

    def intdiv(self):
        return self.binlevel(self.mul, ("\\",))

    def mul(self):
        return self.binlevel(self.unary, ("*", "/"))

    def unary(self):
        tok = self.peek()
        if tok.kind == "op" and tok.val in ("-", "+"):
            self.next()
            return Unary(tok.val, self.unary(), tok.line)
        return self.power()

    def power(self):
        node = self.primary()
        while self.is_op("^"):
            tok = self.next()
            if self.peek().kind == "op" and self.peek().val in ("-", "+"):
                t2 = self.next()
                rhs = Unary(t2.val, self.primary(), t2.line)
            else:
                rhs = self.primary()
            node = Bin("^", node, rhs, tok.line)
        return node

    def primary(self):
        tok = self.next()
        line = tok.line
        if tok.kind in ("num", "str", "date"):
            return Lit(tok.val, line)
        if tok.kind == "op" and tok.val == "(":
            e = self.expr()
            self.expect_op(")")
            node = Paren(e, line)
            return self.postfix(node)
        if tok.kind == "id":
            k = tok.low
            if k == "true":
                return Lit(True, line)
            if k == "false":
                return Lit(False, line)
            if k == "empty":
                return Lit(EMPTY, line)
            if k == "null":
                return Lit(NULL, line)
            if k == "nothing":
                return Lit(NOTHING, line)
            if k == "new":
                c = self.ident()
                return NewObj(c.val, line)
            if k in KEYWORDS:
                raise ParseError(line, f"式の中に「{tok.val}」は書けません")
            if k in FORBIDDEN:
                raise ParseError(line, FORBIDDEN[k])
            return self.postfix(Name(tok.val, line))
        raise ParseError(line, f"式が必要な所に「{tok.val}」があります")

    def postfix(self, node):
        while True:
            if self.is_op("."):
                self.next()
                m = self.next()
                if m.kind != "id":
                    raise ParseError(m.line, ". の後ろに名前がありません")
                node = Member(node, m.val, m.line)
            elif self.is_op("("):
                tok = self.next()
                args = self.arglist_paren()
                node = Call(node, args, tok.line)
            else:
                return node


def parse(code):
    try:
        toks = lex(code)
    except LexError as e:
        raise ParseError(e.line, e.msg)
    return Parser(toks).program()


def walk_stmts(stmts):
    """文を入れ子まで含めて全部たどる。"""
    for s in stmts:
        yield s
        if isinstance(s, If):
            for _, b in s.branches:
                yield from walk_stmts(b)
            if s.else_block:
                yield from walk_stmts(s.else_block)
        elif isinstance(s, (For, ForEach, Do)):
            yield from walk_stmts(s.body)
        elif isinstance(s, Select):
            for _, b in s.cases:
                yield from walk_stmts(b)
            if s.else_block:
                yield from walk_stmts(s.else_block)
        elif isinstance(s, FuncDef):
            yield from walk_stmts(s.body)


def walk_stmts_local(stmts):
    """文を入れ子までたどるが、Function / Sub の中には入らない（ページの一番外側の宣言を集めるため）。"""
    for s in stmts:
        if isinstance(s, FuncDef):
            continue
        yield s
        if isinstance(s, If):
            for _, b in s.branches:
                yield from walk_stmts_local(b)
            if s.else_block:
                yield from walk_stmts_local(s.else_block)
        elif isinstance(s, (For, ForEach, Do)):
            yield from walk_stmts_local(s.body)
        elif isinstance(s, Select):
            for _, b in s.cases:
                yield from walk_stmts_local(b)
            if s.else_block:
                yield from walk_stmts_local(s.else_block)


def walk_expr(e):
    if e is None:
        return
    yield e
    if isinstance(e, Member):
        yield from walk_expr(e.obj)
    elif isinstance(e, Call):
        yield from walk_expr(e.callee)
        for a in e.args:
            yield from walk_expr(a)
    elif isinstance(e, (Paren,)):
        yield from walk_expr(e.e)
    elif isinstance(e, Unary):
        yield from walk_expr(e.e)
    elif isinstance(e, Bin):
        yield from walk_expr(e.l)
        yield from walk_expr(e.r)


def stmt_exprs(s):
    """文が直接持つ式。"""
    if isinstance(s, (Dim,)):
        for _, dims in s.items:
            for d in dims or []:
                yield d
    elif isinstance(s, ReDim):
        for _, dims in s.items:
            yield from dims
    elif isinstance(s, Const):
        for _, e in s.items:
            yield e
    elif isinstance(s, Assign):
        yield s.target
        yield s.expr
    elif isinstance(s, CallStmt):
        yield s.callee
        yield from s.args
    elif isinstance(s, If):
        for c, _ in s.branches:
            yield c
    elif isinstance(s, For):
        yield s.start
        yield s.end
        if s.step is not None:
            yield s.step
    elif isinstance(s, ForEach):
        yield s.expr
    elif isinstance(s, Do):
        if s.pre:
            yield s.pre[1]
        if s.post:
            yield s.post[1]
    elif isinstance(s, Select):
        yield s.expr
        for vals, _ in s.cases:
            yield from vals
