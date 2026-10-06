"""VBScript の値と型変換。"""
import datetime
import math


class VBError(Exception):
    """VBScript の実行時エラー。number は VBScript の番号（ADO のものは負の HRESULT）。"""

    def __init__(self, number, description, source="Microsoft VBScript 実行時エラー", strict=False):
        super().__init__(description)
        self.number = number
        self.description = description
        self.source = source
        self.strict = strict
        self.where = None   # (ファイル, 行)

    def __str__(self):
        w = f" [{self.where[0]}:{self.where[1]}]" if self.where else ""
        return f"{self.source} {self.number}: {self.description}{w}"


class ResponseEnd(VBError):
    """Response.End。本物の ASP では「捕まえられる例外」として実装されているため、
    On Error Resume Next が生きていると握りつぶされる（手順書 7-2）。ここでも同じにする。"""

    def __init__(self):
        super().__init__(-2147467260, "Response.End", "Response")


def strict_error(msg):
    return VBError(9000, "【検査】" + msg, "厳格モード", strict=True)


class _Special:
    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return self.name


EMPTY = _Special("Empty")
NULL = _Special("Null")
NOTHING = _Special("Nothing")


class VDate:
    """VBScript の Date。内部は datetime。"""
    __slots__ = ("dt",)

    def __init__(self, dt):
        if isinstance(dt, datetime.date) and not isinstance(dt, datetime.datetime):
            dt = datetime.datetime(dt.year, dt.month, dt.day)
        self.dt = dt

    def __repr__(self):
        return f"#{self.dt.isoformat()}#"

    def serial(self):
        base = datetime.datetime(1899, 12, 30)
        d = self.dt - base
        return d.days + d.seconds / 86400.0

    @staticmethod
    def from_serial(x):
        base = datetime.datetime(1899, 12, 30)
        return VDate(base + datetime.timedelta(days=x))


class VArray:
    """VBScript の配列（下限は常に 0）。"""

    def __init__(self, ubs, fixed=False, data=None):
        self.ubs = list(ubs)
        self.fixed = fixed
        n = 1
        for u in self.ubs:
            n *= (u + 1) if u >= 0 else 0
        self.data = data if data is not None else [EMPTY] * n

    def _flat(self, idxs):
        if len(idxs) != len(self.ubs):
            raise VBError(9, "インデックスが有効範囲にありません（次元の数が違います）")
        flat = 0
        for i, u in zip(idxs, self.ubs):
            i = to_index(i)
            if i < 0 or i > u:
                raise VBError(9, f"インデックスが有効範囲にありません（添字 {i}、上限 {u}）")
            flat = flat * (u + 1) + i
        return flat

    def get(self, idxs):
        return self.data[self._flat(idxs)]

    def set(self, idxs, v):
        self.data[self._flat(idxs)] = v

    def copy(self):
        return VArray(self.ubs, self.fixed, list(self.data))

    def redim(self, ubs, preserve):
        if self.fixed:
            raise VBError(10, "この配列は固定されているか、または一時的にロックされています")
        if preserve and self.ubs:
            if len(ubs) != len(self.ubs) or ubs[:-1] != self.ubs[:-1]:
                raise VBError(9, "インデックスが有効範囲にありません（Preserve は最後の次元だけ変えられます）")
            new = VArray(ubs)
            # 最後の次元だけ伸縮する
            old = self
            import itertools
            ranges = [range(min(a, b) + 1) for a, b in zip(old.ubs, ubs)]
            for idx in itertools.product(*ranges):
                new.set(list(idx), old.get(list(idx)))
            self.ubs, self.data = new.ubs, new.data
        else:
            new = VArray(ubs)
            self.ubs, self.data = new.ubs, new.data

    def items(self):
        return list(self.data)


class VBinary:
    """バイト配列（ADODB.Stream.Read などが返す）。VarType は 8209（vbArray + vbByte）。"""

    def __init__(self, b):
        self.b = bytes(b)


class Ref:
    """ByRef で渡す変数の参照。"""
    __slots__ = ("get", "set")

    def __init__(self, getter, setter):
        self.get = getter
        self.set = setter


class VBObject:
    """COM オブジェクトの模擬の基底。"""
    progid = "Object"

    def vb_get(self, name, args):
        raise VBError(438, f"オブジェクトでサポートされていないプロパティまたはメソッドです: '{name}'")

    def vb_set(self, name, args, value):
        raise VBError(438, f"オブジェクトでサポートされていないプロパティまたはメソッドです: '{name}'")

    def vb_call(self, name, args):
        return self.vb_get(name, args)

    def vb_default_get(self, args):
        raise VBError(438, "オブジェクトでサポートされていないプロパティまたはメソッドです（既定のプロパティがありません）")

    def vb_default_set(self, args, value):
        raise VBError(438, "オブジェクトでサポートされていないプロパティまたはメソッドです（既定のプロパティがありません）")

    def vb_enum(self):
        raise VBError(451, "オブジェクトはコレクションではありません")

    def type_name(self):
        return type(self).__name__


def is_obj(v):
    return isinstance(v, VBObject) or v is NOTHING


def to_value(v):
    """式の中で使われたオブジェクトは既定のプロパティの値になる。"""
    depth = 0
    while isinstance(v, VBObject):
        v = v.vb_default_get([])
        depth += 1
        if depth > 5:
            break
    if v is NOTHING:
        raise VBError(91, "オブジェクト変数または With ブロック変数が設定されていません")
    return v


def vartype(v):
    if v is EMPTY:
        return 0
    if v is NULL:
        return 1
    if isinstance(v, bool):
        return 11
    if isinstance(v, int):
        return 2 if -32768 <= v <= 32767 else 3
    if isinstance(v, float):
        return 5
    if isinstance(v, VDate):
        return 7
    if isinstance(v, str):
        return 8
    if isinstance(v, VArray):
        return 8204
    if isinstance(v, VBinary):
        return 8209
    if is_obj(v):
        return 9
    return 9


def type_name(v):
    t = vartype(v)
    if t == 9:
        if v is NOTHING:
            return "Nothing"
        return v.type_name()
    return {0: "Empty", 1: "Null", 2: "Integer", 3: "Long", 5: "Double", 7: "Date", 8: "String",
            11: "Boolean", 8204: "Variant()"}[t]


def is_numeric_str(s):
    t = s.strip(" ")
    if t == "":
        return False
    try:
        float(t.replace(",", ""))
        return not any(c in t for c in "nNiI_")  # inf / nan / 1_000 は数値とみなさない
    except ValueError:
        pass
    if t[:2].lower() == "&h":
        try:
            int(t[2:], 16)
            return True
        except ValueError:
            return False
    return False


def num_from_str(s):
    t = s.strip(" ")
    if t[:2].lower() == "&h":
        return int(t[2:], 16)
    t = t.replace(",", "")
    f = float(t)
    if f == int(f) and "." not in t and "e" not in t.lower():
        return int(f)
    return f


def to_index(v):
    v = to_value(v)
    if isinstance(v, bool):
        raise strict_error("真偽値を添字に使っています")
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return vb_round(v)
    if v is EMPTY:
        return 0
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です")
    if isinstance(v, str):
        raise strict_error(f"文字列 \"{v}\" を添字に使っています（CLng で変換してください）")
    raise VBError(13, "型が一致しません")


def vb_round(x):
    """VBScript の CLng / CInt は銀行型丸め（0.5 は偶数へ）。"""
    f = math.floor(x)
    d = x - f
    if d > 0.5:
        return int(f) + 1
    if d < 0.5:
        return int(f)
    return int(f) if int(f) % 2 == 0 else int(f) + 1


def fmt_double(x):
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = repr(float(f"{x:.15g}"))
    if "e" in s:
        m, e = s.split("e")
        e = int(e)
        return f"{m}E{'+' if e > 0 else '-'}{abs(e):02d}"
    return s


def to_str(v, what="文字列化"):
    """CStr 相当。日付・真偽値の暗黙の文字列化は地域設定に依存するので厳格モードで止める。"""
    v = to_value(v)
    if isinstance(v, str):
        return v
    if v is EMPTY:
        return ""
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です")
    if isinstance(v, bool):
        raise strict_error(f"真偽値を{what}しています（\"True\"/\"False\" か \"-1\"/\"0\" かは場面で変わります。明示してください）")
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return fmt_double(v)
    if isinstance(v, VDate):
        raise strict_error(f"日付を{what}しています（書式がサーバーの地域設定で変わります。自前の書式関数を使ってください）")
    if isinstance(v, VArray):
        raise VBError(13, "型が一致しません（配列を文字列にしようとしました）")
    raise VBError(13, "型が一致しません")


def to_num(v, what="計算"):
    v = to_value(v)
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, (int, float)):
        return v
    if v is EMPTY:
        return 0
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です")
    if isinstance(v, str):
        raise strict_error(f"文字列 \"{v[:20]}\" をそのまま{what}に使っています（CLng などで明示的に変換してください）")
    if isinstance(v, VDate):
        raise strict_error(f"日付をそのまま{what}に使っています")
    raise VBError(13, "型が一致しません")


def to_bool(v, what="条件"):
    v = to_value(v)
    if isinstance(v, bool):
        return v
    if v is NULL or v is EMPTY:
        return False
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        raise strict_error(f"文字列 \"{v[:20]}\" を{what}に使っています")
    raise VBError(13, "型が一致しません")


def norm_int(x):
    """整数演算の結果。Long の範囲を超えたら Double（VBScript の Variant と同じ）。"""
    if isinstance(x, int) and not isinstance(x, bool):
        if -2147483648 <= x <= 2147483647:
            return x
        return float(x)
    return x


def copy_val(v):
    if isinstance(v, VArray):
        return v.copy()
    return v
