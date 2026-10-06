"""VBScript の実行（厳格モード）。"""
import datetime

from . import parser as P
from .builtins import VB_CONSTS
from .values import (VBError, ResponseEnd, strict_error, EMPTY, NULL, NOTHING, VDate, VArray, VBObject,
                     to_value, to_str, to_num, to_bool, norm_int, copy_val, is_obj, vb_round, vartype)


class ExitLoop(Exception):
    def __init__(self, kind):
        self.kind = kind


class ExitProc(Exception):
    pass


class Cell:
    __slots__ = ("v",)

    def __init__(self, v=EMPTY):
        self.v = v

    def get(self):
        return self.v

    def set(self, v):
        self.v = v


class ElemCell(Cell):
    """配列の要素を ByRef で渡したときの参照。"""
    __slots__ = ("arr", "idx")

    def __init__(self, arr, idx):
        self.arr, self.idx = arr, idx

    def get(self):
        return self.arr.get(self.idx)

    def set(self, v):
        self.arr.set(self.idx, v)


class Frame:
    def __init__(self, proc=None):
        self.proc = proc
        self.vars = {}
        self.consts = {}
        self.orn = False


class ErrObject(VBObject):
    def __init__(self):
        self.clear()

    def clear(self):
        self.number = 0
        self.description = ""
        self.source = ""

    def type_name(self):
        return "ErrObject"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "number":
            return self.number
        if n == "description":
            return self.description
        if n == "source":
            return self.source
        if n == "clear":
            self.clear()
            return EMPTY
        if n == "raise":
            num = to_num(args[0])
            src = to_str(args[1]) if len(args) > 1 else "Microsoft VBScript 実行時エラー"
            desc = to_str(args[2]) if len(args) > 2 else f"エラー {num}"
            raise VBError(num, desc, src)
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        return self.number


def _is_proc_name(stmt):
    return isinstance(stmt, P.FuncDef)


class Interp:
    """1 ページぶんの実行。host は ASP の組み込みオブジェクトを提供する（host.py）。"""

    def __init__(self, prog, ast, host, builtins):
        self.prog = prog
        self.ast = ast
        self.host = host
        self.bi = builtins
        self.g = Frame()
        self.procs = {}
        self.err = ErrObject()
        self.explicit = False
        self.cur_line = 0
        self.warnings = []
        self.call_depth = 0

    # ---------- 位置 ----------
    def where(self, line):
        if 0 <= line < len(self.prog.linemap):
            return self.prog.linemap[line]
        return ("?", 0)

    def warn(self, msg, line=None):
        f, l = self.where(self.cur_line if line is None else line)
        self.warnings.append(f"{f}:{l}: {msg}")

    # ---------- 準備 ----------
    def prepare(self):
        body = self.ast
        # Option Explicit は最初の文でなければならない
        for i, s in enumerate(body):
            if isinstance(s, P.OptionExplicit):
                if i != 0:
                    raise VBError(1002, "Option Explicit はページの最初の文でなければなりません（前に出力や文があります）", "Microsoft VBScript コンパイル エラー")
                self.explicit = True
        names = {}
        for s in body:
            if isinstance(s, P.FuncDef):
                if s.low in self.procs:
                    raise self._redef(s.name, s.line)
                self.procs[s.low] = s
        self._hoist(body, self.g, top=True)

    def _redef(self, name, line):
        e = VBError(1041, f"名前が再定義されています: '{name}'", "Microsoft VBScript コンパイル エラー")
        e.where = self.where(line)
        return e

    def _hoist(self, body, frame, top=False):
        for s in P.walk_stmts_local(body):
            if isinstance(s, P.Dim):
                for name, dims in s.items:
                    low = name.lower()
                    if low in frame.vars or low in frame.consts or (frame is self.g and low in self.procs):
                        raise self._redef(name, s.line)
                    if dims:
                        ubs = []
                        for d in dims:
                            if not isinstance(d, P.Lit) or not isinstance(d.v, int):
                                e = VBError(1028, "Dim の配列の大きさは数字で書いてください（変数で決めるなら ReDim）", "Microsoft VBScript コンパイル エラー")
                                e.where = self.where(s.line)
                                raise e
                            ubs.append(d.v)
                        frame.vars[low] = Cell(VArray(ubs, fixed=True))
                    elif dims == []:
                        frame.vars[low] = Cell(VArray([-1]))
                    else:
                        frame.vars[low] = Cell(EMPTY)
            elif isinstance(s, P.Const):
                for name, e in s.items:
                    low = name.lower()
                    if low in frame.vars or low in frame.consts or (frame is self.g and low in self.procs):
                        raise self._redef(name, s.line)
                    v = self._const_value(e, s.line)
                    frame.consts[low] = v

    def _const_value(self, e, line):
        # VBScript の Const は「定数（数字・文字列・日付・True/False）」しか書けない
        if isinstance(e, P.Lit):
            return e.v
        if isinstance(e, P.Unary) and e.op == "-" and isinstance(e.e, P.Lit) and isinstance(e.e.v, (int, float)):
            return -e.e.v
        err = VBError(1045, "Const には定数しか書けません（式や他の Const は使えません）", "Microsoft VBScript コンパイル エラー")
        err.where = self.where(line)
        raise err

    # ---------- 実行 ----------
    def run(self):
        self.prepare()
        self.exec_block(self.ast, self.g)

    def exec_block(self, body, frame):
        for s in body:
            if isinstance(s, P.FuncDef):
                continue
            self.cur_line = s.line
            try:
                self.exec_stmt(s, frame)
            except VBError as e:
                if e.where is None:
                    e.where = self.where(s.line)
                if frame.orn and not getattr(e, "fatal", False):
                    if isinstance(e, ResponseEnd):
                        self.warn("Response.End が On Error Resume Next に握りつぶされました（ページが止まらずに続きます）", s.line)
                    self.err.number = e.number
                    self.err.description = e.description
                    self.err.source = e.source
                    if e.strict:
                        # 厳格モードの検出は握りつぶさせない（検査が目的なので）
                        raise
                    continue
                raise

    def exec_stmt(self, s, f):
        t = type(s)
        if t is P.WriteBlock:
            self.host.write_block(self.prog.blocks[s.idx])
        elif t is P.Assign:
            self.exec_assign(s, f)
        elif t is P.CallStmt:
            self.exec_call_stmt(s, f)
        elif t is P.If:
            self.exec_if(s, f)
        elif t is P.For:
            self.exec_for(s, f)
        elif t is P.ForEach:
            self.exec_foreach(s, f)
        elif t is P.Do:
            self.exec_do(s, f)
        elif t is P.Select:
            self.exec_select(s, f)
        elif t is P.Exit:
            if s.kind in ("function", "sub"):
                raise ExitProc()
            raise ExitLoop(s.kind)
        elif t is P.OnError:
            f.orn = s.resume_next
            self.err.clear()
        elif t is P.Dim or t is P.Const or t is P.OptionExplicit or t is P.Randomize:
            pass
        elif t is P.ReDim:
            for name, dims in s.items:
                cell = self.lookup_cell(name, f)
                if cell is None:
                    if self.explicit:
                        raise VBError(500, f"変数が定義されていません: '{name}'")
                    cell = f.vars.setdefault(name.lower(), Cell(VArray([-1])))
                ubs = [int(to_num(self.eval(d, f))) for d in dims]
                cur = cell.get()
                if isinstance(cur, VArray):
                    if cur.fixed:
                        raise VBError(10, "この配列は固定されているか、または一時的にロックされています")
                    cur.redim(ubs, s.preserve)
                else:
                    if s.preserve:
                        raise VBError(9, "インデックスが有効範囲にありません（配列でない変数に ReDim Preserve）")
                    cell.set(VArray(ubs))
        elif t is P.Erase:
            cell = self.lookup_cell(s.name, f)
            v = cell.get() if cell else None
            if isinstance(v, VArray):
                if v.fixed:
                    v.data = [EMPTY] * len(v.data)
                else:
                    cell.set(VArray([-1]))
        else:
            raise VBError(5, f"この文は実行できません: {t.__name__}")

    # ---------- 変数 ----------
    def lookup_cell(self, name, f):
        low = name.lower()
        if low in f.vars:
            return f.vars[low]
        if f is not self.g and low in self.g.vars:
            return self.g.vars[low]
        return None

    def lookup_const(self, low, f):
        if low in VB_CONSTS and low not in f.vars and low not in self.g.vars:
            return True, VB_CONSTS[low]
        if low in f.consts:
            return True, f.consts[low]
        if low in self.g.consts:
            return True, self.g.consts[low]
        return False, None

    # ---------- 代入 ----------
    def exec_assign(self, s, f):
        if s.is_set:
            v = self.eval(s.expr, f)
            if not is_obj(v):
                raise VBError(424, "オブジェクトが必要です（Set の右辺がオブジェクトではありません）")
        else:
            v = self.eval(s.expr, f)
            if isinstance(v, VBObject):
                v = to_value(v)
            v = copy_val(v)
        self.assign_to(s.target, v, f, s.is_set)

    def assign_to(self, target, v, f, is_set=False):
        if isinstance(target, P.Name):
            low = target.low
            isc, _ = self.lookup_const(low, f)
            if isc:
                raise VBError(501, f"代入できません（'{target.name}' は Const です）")
            cell = self.lookup_cell(low, f)
            if cell is None:
                if low in self.procs and not (f.proc and f.proc.low == low):
                    raise VBError(501, f"代入できません（'{target.name}' は Function / Sub の名前です）")
                if self.explicit:
                    raise VBError(500, f"変数が定義されていません: '{target.name}'")
                cell = f.vars.setdefault(low, Cell())
            cur = cell.get()
            if isinstance(cur, VArray) and cur.fixed:
                raise VBError(10, "この配列は固定されているか、または一時的にロックされています")
            cell.set(v)
            return
        if isinstance(target, P.Call):
            callee = target.callee
            if isinstance(callee, P.Name):
                cell = self.lookup_cell(callee.low, f)
                if cell is not None:
                    cur = cell.get()
                    if isinstance(cur, VArray):
                        idx = [self.eval_value(a, f) for a in target.args]
                        cur.set(idx, v)
                        return
                    if isinstance(cur, VBObject):
                        cur.vb_default_set([self.eval(a, f) for a in target.args], v)
                        return
                    raise VBError(13, f"型が一致しません: '{callee.name}'")
                obj = self.host.intrinsic(callee.low)
                if obj is not None:
                    obj.vb_default_set([self.eval(a, f) for a in target.args], v)
                    return
                if self.explicit:
                    raise VBError(500, f"変数が定義されていません: '{callee.name}'")
                raise VBError(13, f"型が一致しません: '{callee.name}'")
            if isinstance(callee, P.Member):
                o = self.eval_obj(callee.obj, f)
                o.vb_set(callee.name, [self.eval(a, f) for a in target.args], v)
                return
            base = self.eval(callee, f)
            if isinstance(base, VArray):
                base.set([self.eval_value(a, f) for a in target.args], v)
                return
            if isinstance(base, VBObject):
                base.vb_default_set([self.eval(a, f) for a in target.args], v)
                return
            raise VBError(13, "型が一致しません")
        if isinstance(target, P.Member):
            o = self.eval_obj(target.obj, f)
            o.vb_set(target.name, [], v)
            return
        raise VBError(501, "代入できません")

    # ---------- 呼び出し ----------
    def exec_call_stmt(self, s, f):
        callee = s.callee
        if isinstance(callee, P.Name):
            low = callee.low
            cell = self.lookup_cell(low, f)
            if cell is not None and not (low in self.procs and not s.args and cell.get() is EMPTY and False):
                cur = cell.get()
                if isinstance(cur, VBObject) and not s.args:
                    return
                if isinstance(cur, VBObject):
                    cur.vb_default_get([self.eval(a, f) for a in s.args])
                    return
                if low not in self.procs:
                    raise VBError(13, f"型が一致しません: '{callee.name}'（変数を Sub のように呼んでいます）")
            if low in self.procs:
                self.invoke(self.procs[low], s.args, f)
                return
            if self.bi.has(low):
                self.bi.call(self, low, [self.eval(a, f) for a in s.args])
                return
            obj = self.host.intrinsic(low)
            if obj is not None:
                if s.args:
                    obj.vb_default_get([self.eval(a, f) for a in s.args])
                return
            raise VBError(13, f"型が一致しません: '{callee.name}'（その名前の Function / Sub がありません）")
        if isinstance(callee, P.Member):
            o = self.eval_obj(callee.obj, f)
            o.vb_call(callee.name, self.host_args(s.args, f))
            return
        self.eval(P.Call(callee, s.args, s.line) if s.args else callee, f)

    def host_args(self, args, f):
        """COM のメソッドに渡す引数。変数は参照で渡す（ADO の RecordsAffected などのため）。"""
        out = []
        for a in args:
            if isinstance(a, P.Name):
                cell = self.lookup_cell(a.low, f)
                if cell is not None:
                    out.append(cell)
                    continue
            out.append(self.eval(a, f))
        return out

    def invoke(self, proc, arg_nodes, caller, line=None):
        if len(arg_nodes) != len(proc.params):
            raise VBError(450, f"引数の数が一致していません: '{proc.name}'（{len(proc.params)} 個のところ {len(arg_nodes)} 個）")
        fr = Frame(proc)
        for (pname, byval, is_arr), a in zip(proc.params, arg_nodes):
            low = pname.lower()
            cell = None
            if not byval:
                cell = self.ref_cell(a, caller)
            if cell is None:
                v = self.eval(a, caller)
                cell = Cell(copy_val(v) if not isinstance(v, VBObject) else v)
            if low in fr.vars:
                raise self._redef(pname, proc.line)
            fr.vars[low] = cell
        if proc.kind == "function":
            if proc.low in fr.vars:
                raise self._redef(proc.name, proc.line)
            fr.vars[proc.low] = Cell(EMPTY)
        self._hoist(proc.body, fr)
        self.call_depth += 1
        if self.call_depth > 200:
            raise VBError(28, "スタック領域が不足しています（呼び出しが深すぎます）")
        saved = self.cur_line
        try:
            self.exec_block(proc.body, fr)
        except ExitProc:
            pass
        finally:
            self.call_depth -= 1
            self.cur_line = saved
        if proc.kind == "function":
            return fr.vars[proc.low].get()
        return EMPTY

    def ref_cell(self, a, f):
        if isinstance(a, P.Name):
            isc, _ = self.lookup_const(a.low, f)
            if isc:
                return None
            cell = self.lookup_cell(a.low, f)
            if cell is not None and not (f.proc and f.proc.low == a.low and False):
                return cell
            return None
        if isinstance(a, P.Call) and isinstance(a.callee, P.Name):
            cell = self.lookup_cell(a.callee.low, f)
            if cell is not None and isinstance(cell.get(), VArray):
                idx = [self.eval_value(x, f) for x in a.args]
                arr = cell.get()
                arr._flat(idx)
                return ElemCell(arr, idx)
        return None

    # ---------- 式 ----------
    def eval_value(self, e, f):
        return to_value(self.eval(e, f))

    def eval_obj(self, e, f):
        o = self.eval(e, f)
        if o is NOTHING:
            raise VBError(91, "オブジェクト変数が設定されていません（Nothing のまま使っています）")
        if not isinstance(o, VBObject):
            raise VBError(424, "オブジェクトが必要です")
        return o

    def eval(self, e, f):
        t = type(e)
        if t is P.Lit:
            return e.v
        if t is P.Name:
            return self.eval_name(e, f)
        if t is P.Member:
            o = self.eval_obj(e.obj, f)
            return o.vb_get(e.name, [])
        if t is P.Call:
            return self.eval_call(e, f)
        if t is P.Paren:
            return self.eval(e.e, f)
        if t is P.Bin:
            return self.eval_bin(e, f)
        if t is P.Unary:
            return self.eval_unary(e, f)
        if t is P.NewObj:
            raise strict_error(f"New {e.cls} は使わない決まりです")
        raise VBError(5, "式を評価できません")

    def eval_name(self, e, f):
        low = e.low
        cell = f.vars.get(low)
        if cell is not None:
            return cell.get()
        if low in f.consts:
            return f.consts[low]
        if f is not self.g:
            cell = self.g.vars.get(low)
            if cell is not None:
                return cell.get()
        if low in self.g.consts:
            return self.g.consts[low]
        if low in self.procs:
            p = self.procs[low]
            if p.kind == "sub":
                raise VBError(13, f"型が一致しません: Sub '{p.name}' を式の中で使っています")
            return self.invoke(p, [], f)
        if low == "err":
            return self.err
        if low in VB_CONSTS:
            return VB_CONSTS[low]
        obj = self.host.intrinsic(low)
        if obj is not None:
            return obj
        if self.bi.has(low):
            return self.bi.call(self, low, [])
        if self.explicit:
            raise VBError(500, f"変数が定義されていません: '{e.name}'")
        return EMPTY

    def eval_call(self, e, f):
        callee = e.callee
        if isinstance(callee, P.Name):
            low = callee.low
            cell = f.vars.get(low)
            if cell is None and f is not self.g:
                cell = self.g.vars.get(low)
            is_self_ret = f.proc is not None and f.proc.low == low
            if cell is not None and not is_self_ret:
                cur = cell.get()
                if isinstance(cur, VArray):
                    return cur.get([self.eval_value(a, f) for a in e.args])
                if isinstance(cur, VBObject):
                    return cur.vb_default_get([self.eval(a, f) for a in e.args])
                if cur is NOTHING:
                    raise VBError(91, f"オブジェクト変数が設定されていません: '{callee.name}'")
                raise VBError(13, f"型が一致しません: '{callee.name}'（配列でもオブジェクトでもない変数にかっこを付けています）")
            isc, _ = self.lookup_const(low, f)
            if isc:
                raise VBError(13, f"型が一致しません: '{callee.name}'")
            if low in self.procs:
                p = self.procs[low]
                if p.kind == "sub":
                    raise VBError(13, f"型が一致しません: Sub '{p.name}' を式の中で使っています")
                return self.invoke(p, e.args, f)
            if low == "err":
                raise VBError(13, "型が一致しません: 'Err'")
            if self.bi.has(low):
                return self.bi.call(self, low, [self.eval(a, f) for a in e.args])
            obj = self.host.intrinsic(low)
            if obj is not None:
                return obj.vb_default_get([self.eval(a, f) for a in e.args])
            raise VBError(13, f"型が一致しません: '{callee.name}'（その名前の Function / 配列がありません）")
        if isinstance(callee, P.Member):
            o = self.eval_obj(callee.obj, f)
            return o.vb_call(callee.name, self.host_args(e.args, f))
        base = self.eval(callee, f)
        if isinstance(base, VArray):
            return base.get([self.eval_value(a, f) for a in e.args])
        if isinstance(base, VBObject):
            return base.vb_default_get([self.eval(a, f) for a in e.args])
        raise VBError(13, "型が一致しません")

    def eval_unary(self, e, f):
        v = self.eval_value(e.e, f)
        if e.op == "not":
            if v is NULL:
                return NULL
            if isinstance(v, bool):
                return not v
            if isinstance(v, int):
                return ~v
            if v is EMPTY:
                return -1
            raise strict_error("Not を数値・真偽値以外に使っています")
        if v is NULL:
            return NULL
        n = to_num(v)
        return norm_int(-n) if e.op == "-" else n

    def eval_bin(self, e, f):
        op = e.op
        # VBScript の And / Or は「短絡評価しない」。両辺を必ず評価する（本物と同じ）。
        a = self.eval(e.l, f)
        b = self.eval(e.r, f)
        if op == "is":
            if not is_obj(a) or not is_obj(b):
                raise VBError(424, "オブジェクトが必要です（Is の両辺はオブジェクト）")
            return a is b
        a = to_value(a)
        b = to_value(b)
        if op == "&":
            if a is NULL and b is NULL:
                return NULL
            return (to_str(a, "文字列連結") if a is not NULL else "") + (to_str(b, "文字列連結") if b is not NULL else "")
        if op in ("=", "<>", "<", ">", "<=", ">="):
            c = compare(a, b)
            if c is NULL:
                return NULL
            return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
        if op in ("and", "or", "xor", "eqv", "imp"):
            return logic(op, a, b)
        if a is NULL or b is NULL:
            return NULL
        if op == "+":
            if isinstance(a, str) and isinstance(b, str):
                raise strict_error("文字列どうしを + でつないでいます（& を使ってください）")
            if isinstance(a, VDate) and not isinstance(b, VDate):
                return VDate.from_serial(a.serial() + to_num(b))
            if isinstance(b, VDate) and not isinstance(a, VDate):
                return VDate.from_serial(b.serial() + to_num(a))
            return norm_int(to_num(a) + to_num(b))
        if op == "-":
            if isinstance(a, VDate) and isinstance(b, VDate):
                return a.serial() - b.serial()
            if isinstance(a, VDate):
                return VDate.from_serial(a.serial() - to_num(b))
            return norm_int(to_num(a) - to_num(b))
        if op == "*":
            return norm_int(to_num(a) * to_num(b))
        if op == "/":
            d = to_num(b)
            if d == 0:
                raise VBError(11, "0 で除算しました")
            return to_num(a) / d
        if op == "\\":
            x, y = vb_round(to_num(a)), vb_round(to_num(b))
            if y == 0:
                raise VBError(11, "0 で除算しました")
            q = abs(x) // abs(y)
            return q if (x >= 0) == (y >= 0) else -q
        if op == "mod":
            x, y = vb_round(to_num(a)), vb_round(to_num(b))
            if y == 0:
                raise VBError(11, "0 で除算しました")
            r = abs(x) % abs(y)
            return r if x >= 0 else -r
        if op == "^":
            return float(to_num(a)) ** float(to_num(b))
        raise VBError(5, f"演算子 {op}")

    def exec_if(self, s, f):
        for cond, body in s.branches:
            try:
                ok = to_bool(self.eval(cond, f))
            except VBError as e:
                if f.orn and not e.strict and not isinstance(e, ResponseEnd):
                    # 本物の VBScript は条件式でエラーが起きると「次の文」＝ Then の中へ進む
                    self.warn("If の条件式でエラーが起き、On Error Resume Next により Then の中へ進みました")
                    self.err.number, self.err.description = e.number, e.description
                    ok = True
                else:
                    raise
            if ok:
                self.exec_block(body, f)
                return
        if s.else_block is not None:
            self.exec_block(s.else_block, f)

    def exec_for(self, s, f):
        start = to_num(self.eval_value(s.start, f), "For の開始値")
        end = to_num(self.eval_value(s.end, f), "For の終了値")
        step = to_num(self.eval_value(s.step, f), "For の Step") if s.step is not None else 1
        var = P.Name(s.var, s.line)
        i = start
        self.assign_to(var, i, f)
        while (step >= 0 and i <= end) or (step < 0 and i >= end):
            try:
                self.exec_block(s.body, f)
            except ExitLoop as x:
                if x.kind == "for":
                    return
                raise
            i = to_num(self.eval_name(var, f), "For の変数") + step
            self.assign_to(var, norm_int(i), f)

    def exec_foreach(self, s, f):
        coll = self.eval(s.expr, f)
        if isinstance(coll, VArray):
            items = coll.items()
        elif isinstance(coll, VBObject):
            items = coll.vb_enum()
        else:
            raise VBError(451, "For Each の対象がコレクションでも配列でもありません")
        var = P.Name(s.var, s.line)
        for it in items:
            cell = self.lookup_cell(s.var, f)
            if cell is None:
                if self.explicit:
                    raise VBError(500, f"変数が定義されていません: '{s.var}'")
                cell = f.vars.setdefault(s.var.lower(), Cell())
            cell.set(it)
            try:
                self.exec_block(s.body, f)
            except ExitLoop as x:
                if x.kind == "for":
                    return
                raise

    def exec_do(self, s, f):
        guard = 0
        while True:
            if s.pre:
                c = to_bool(self.eval(s.pre[1], f))
                if (s.pre[0] == "while" and not c) or (s.pre[0] == "until" and c):
                    return
            try:
                self.exec_block(s.body, f)
            except ExitLoop as x:
                if x.kind == "do":
                    return
                raise
            if s.post:
                c = to_bool(self.eval(s.post[1], f))
                if (s.post[0] == "while" and not c) or (s.post[0] == "until" and c):
                    return
            guard += 1
            if guard > 1000000:
                raise strict_error("Do ループが 100 万回を超えました（終わらないループの疑い）")

    def exec_select(self, s, f):
        v = self.eval_value(s.expr, f)
        for vals, body in s.cases:
            for ce in vals:
                c = compare(v, self.eval_value(ce, f))
                if c is not NULL and c == 0:
                    self.exec_block(body, f)
                    return
        if s.else_block is not None:
            self.exec_block(s.else_block, f)


def compare(a, b):
    """比較。文字列と数値を比べるような「本物では結果が場面で変わる」比較は厳格モードで止める。"""
    if a is NULL or b is NULL:
        return NULL
    if isinstance(a, VArray) or isinstance(b, VArray):
        raise VBError(13, "型が一致しません（配列を比較しています）")
    if a is EMPTY and b is EMPTY:
        return 0
    if a is EMPTY:
        a = "" if isinstance(b, str) else 0
    if b is EMPTY:
        b = "" if isinstance(a, str) else 0
    if isinstance(a, str) and isinstance(b, str):
        return (a > b) - (a < b)
    if isinstance(a, VDate) and isinstance(b, VDate):
        x, y = a.serial(), b.serial()
        return (x > y) - (x < y)
    if isinstance(a, bool) and isinstance(b, bool):
        x, y = (-1 if a else 0), (-1 if b else 0)
        return (x > y) - (x < y)
    if isinstance(a, bool) or isinstance(b, bool):
        raise strict_error("真偽値と数値・文字列を比べています")
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return (a > b) - (a < b)
    if isinstance(a, str) or isinstance(b, str):
        raise strict_error(f"文字列と数値（または日付）を比べています: {a!r} と {b!r}（本物の VBScript では書き方次第で結果が変わります。CLng / CStr で揃えてください）")
    if isinstance(a, VDate) or isinstance(b, VDate):
        raise strict_error("日付と数値を比べています")
    raise VBError(13, "型が一致しません")


def logic(op, a, b):
    def tb(v):
        if v is NULL:
            return NULL
        if v is EMPTY:
            return False
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return int(vb_round(v))
        raise strict_error(f"{op} を文字列などに使っています")
    x, y = tb(a), tb(b)
    if isinstance(x, bool) and isinstance(y, bool):
        return {"and": x and y, "or": x or y, "xor": x != y, "eqv": x == y, "imp": (not x) or y}[op]
    if x is NULL or y is NULL:
        if op == "and":
            if x is False or y is False:
                return False
            return NULL
        if op == "or":
            if x is True or y is True:
                return True
            return NULL
        return NULL
    xi = -1 if x is True else (0 if x is False else x)
    yi = -1 if y is True else (0 if y is False else y)
    return {"and": xi & yi, "or": xi | yi, "xor": xi ^ yi, "eqv": ~(xi ^ yi), "imp": (~xi) | yi}[op]
