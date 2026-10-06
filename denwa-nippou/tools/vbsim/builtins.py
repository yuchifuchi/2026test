"""VBScript の組み込み関数。本物の癖（AscW が負の値を返す、Trim は半角空白だけ、など）を再現する。"""
import calendar
import datetime
import math
import random
import time

from .values import (VBError, strict_error, EMPTY, NULL, NOTHING, VDate, VArray, VBObject, VBinary, to_value, to_str,
                     to_num, to_bool, vb_round, vartype, type_name, is_numeric_str, num_from_str, norm_int, is_obj)

FUNCS = {}
ARITY = {}
NOW_OVERRIDE = [None]   # 検査で「今日」を固定するため


def now():
    return NOW_OVERRIDE[0] or datetime.datetime.now()


def bi(name, lo, hi=None):
    def deco(fn):
        FUNCS[name.lower()] = fn
        ARITY[name.lower()] = (lo, lo if hi is None else hi)
        return fn
    return deco


def V(a):
    return to_value(a)


def null_pass(fn):
    def w(args):
        if args and V(args[0]) is NULL:
            return NULL
        return fn(args)
    return w


def as_str_arg(v, fname):
    v = V(v)
    if isinstance(v, str):
        return v
    if v is EMPTY:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return to_str(v)
    if v is NULL:
        raise VBError(94, f"Null の使い方が不正です: {fname}")
    return to_str(v, f"{fname} の引数として文字列化")


def as_int_arg(v, fname):
    v = V(v)
    if v is NULL:
        raise VBError(94, f"Null の使い方が不正です: {fname}")
    if isinstance(v, str):
        raise strict_error(f"{fname} の数値の引数に文字列 \"{v}\" を渡しています")
    return vb_round(to_num(v))


# ---------- 文字列 ----------
@bi("Len", 1)
def _len(a):
    v = V(a[0])
    if v is NULL:
        return NULL
    return len(as_str_arg(v, "Len"))


@bi("Left", 2)
def _left(a):
    if V(a[0]) is NULL:
        return NULL
    s = as_str_arg(a[0], "Left")
    n = as_int_arg(a[1], "Left")
    if n < 0:
        raise VBError(5, "プロシージャの呼び出し、または引数が不正です: Left")
    return s[:n]


@bi("Right", 2)
def _right(a):
    if V(a[0]) is NULL:
        return NULL
    s = as_str_arg(a[0], "Right")
    n = as_int_arg(a[1], "Right")
    if n < 0:
        raise VBError(5, "プロシージャの呼び出し、または引数が不正です: Right")
    return s[len(s) - n:] if n else ""


@bi("Mid", 2, 3)
def _mid(a):
    if V(a[0]) is NULL:
        return NULL
    s = as_str_arg(a[0], "Mid")
    st = as_int_arg(a[1], "Mid")
    if st < 1:
        raise VBError(5, "プロシージャの呼び出し、または引数が不正です: Mid の開始位置は 1 以上")
    if len(a) > 2:
        n = as_int_arg(a[2], "Mid")
        if n < 0:
            raise VBError(5, "プロシージャの呼び出し、または引数が不正です: Mid")
        return s[st - 1:st - 1 + n]
    return s[st - 1:]


@bi("InStr", 2, 4)
def _instr(a):
    if len(a) >= 3:
        start = as_int_arg(a[0], "InStr")
        s1, s2 = V(a[1]), V(a[2])
        mode = as_int_arg(a[3], "InStr") if len(a) > 3 else 0
    else:
        start, s1, s2, mode = 1, V(a[0]), V(a[1]), 0
    if s1 is NULL or s2 is NULL:
        return NULL
    if start < 1:
        raise VBError(5, "InStr の開始位置は 1 以上")
    s1, s2 = as_str_arg(s1, "InStr"), as_str_arg(s2, "InStr")
    if mode == 1:
        s1, s2 = s1.lower(), s2.lower()
    if s2 == "":
        return start if start <= len(s1) + 1 else 0
    return s1.find(s2, start - 1) + 1


@bi("InStrRev", 2, 4)
def _instrrev(a):
    s1, s2 = V(a[0]), V(a[1])
    if s1 is NULL or s2 is NULL:
        return NULL
    s1, s2 = as_str_arg(s1, "InStrRev"), as_str_arg(s2, "InStrRev")
    start = as_int_arg(a[2], "InStrRev") if len(a) > 2 else -1
    if start == -1:
        start = len(s1)
    return s1.rfind(s2, 0, start) + 1


@bi("Replace", 3, 6)
def _replace(a):
    if V(a[0]) is NULL:
        return NULL
    s, fnd, rep = as_str_arg(a[0], "Replace"), as_str_arg(a[1], "Replace"), as_str_arg(a[2], "Replace")
    start = as_int_arg(a[3], "Replace") if len(a) > 3 else 1
    count = as_int_arg(a[4], "Replace") if len(a) > 4 else -1
    mode = as_int_arg(a[5], "Replace") if len(a) > 5 else 0
    s = s[start - 1:]
    if fnd == "":
        return s
    if mode == 1:
        out, i, n = [], 0, 0
        ls, lf = s.lower(), fnd.lower()
        while True:
            j = ls.find(lf, i)
            if j < 0 or (count != -1 and n >= count):
                out.append(s[i:])
                break
            out.append(s[i:j] + rep)
            i = j + len(fnd)
            n += 1
        return "".join(out)
    return s.replace(fnd, rep, count if count != -1 else -1)


def _trim_sp(s, left, right):
    # VBScript の Trim は半角空白（&H20）だけを取る。全角空白・タブは残る。
    if left:
        s = s.lstrip(" ")
    if right:
        s = s.rstrip(" ")
    return s


@bi("Trim", 1)
def _trim(a):
    return NULL if V(a[0]) is NULL else _trim_sp(as_str_arg(a[0], "Trim"), True, True)


@bi("LTrim", 1)
def _ltrim(a):
    return NULL if V(a[0]) is NULL else _trim_sp(as_str_arg(a[0], "LTrim"), True, False)


@bi("RTrim", 1)
def _rtrim(a):
    return NULL if V(a[0]) is NULL else _trim_sp(as_str_arg(a[0], "RTrim"), False, True)


@bi("UCase", 1)
def _ucase(a):
    return NULL if V(a[0]) is NULL else as_str_arg(a[0], "UCase").upper()


@bi("LCase", 1)
def _lcase(a):
    return NULL if V(a[0]) is NULL else as_str_arg(a[0], "LCase").lower()


@bi("Space", 1)
def _space(a):
    return " " * as_int_arg(a[0], "Space")


@bi("String", 2)
def _string(a):
    n = as_int_arg(a[0], "String")
    c = V(a[1])
    ch = chr(c) if isinstance(c, int) else as_str_arg(c, "String")[:1]
    return ch * n


@bi("StrReverse", 1)
def _strrev(a):
    return as_str_arg(a[0], "StrReverse")[::-1]


@bi("StrComp", 2, 3)
def _strcomp(a):
    x, y = V(a[0]), V(a[1])
    if x is NULL or y is NULL:
        return NULL
    x, y = as_str_arg(x, "StrComp"), as_str_arg(y, "StrComp")
    if len(a) > 2 and as_int_arg(a[2], "StrComp") == 1:
        x, y = x.lower(), y.lower()
    return (x > y) - (x < y)


@bi("Split", 1, 4)
def _split(a):
    s = V(a[0])
    if s is NULL:
        raise VBError(94, "Null の使い方が不正です: Split")
    s = as_str_arg(s, "Split")
    d = as_str_arg(a[1], "Split") if len(a) > 1 else " "
    if s == "":
        return VArray([-1])
    parts = s.split(d) if d != "" else [s]
    return VArray([len(parts) - 1], data=list(parts))


@bi("Join", 1, 2)
def _join(a):
    arr = V(a[0])
    if not isinstance(arr, VArray):
        raise VBError(13, "型が一致しません: Join")
    d = as_str_arg(a[1], "Join") if len(a) > 1 else " "
    return d.join(to_str(x, "Join で文字列化") for x in arr.items())


@bi("Chr", 1)
def _chr(a):
    n = as_int_arg(a[0], "Chr")
    if n < 0 or n > 255:
        raise strict_error("Chr に 0～255 以外を渡しています（日本語は ChrW を使う）")
    return chr(n)


@bi("ChrW", 1)
def _chrw(a):
    n = as_int_arg(a[0], "ChrW")
    if n < -32768 or n > 65535:
        raise VBError(5, "ChrW の引数が範囲外です")
    return chr(n & 0xFFFF)


@bi("Asc", 1)
def _asc(a):
    s = as_str_arg(a[0], "Asc")
    if s == "":
        raise VBError(5, "Asc に空文字を渡しています")
    c = ord(s[0])
    if c > 127:
        raise strict_error("Asc は日本語で地域設定依存の値を返します（AscW を使う）")
    return c


@bi("AscW", 1)
def _ascw(a):
    s = as_str_arg(a[0], "AscW")
    if s == "":
        raise VBError(5, "AscW に空文字を渡しています")
    c = ord(s[0])
    # 本物の AscW は Integer（16 ビット符号付き）を返す。&H8000 以上は負になる。
    return c - 0x10000 if c >= 0x8000 else c


@bi("Hex", 1)
def _hex(a):
    v = V(a[0])
    if v is NULL:
        return NULL
    n = vb_round(to_num(v))
    if n < 0:
        n &= 0xFFFFFFFF if n < -32768 else 0xFFFF
    return format(n, "X")


@bi("Escape", 1)
def _escape(a):
    s = as_str_arg(a[0], "Escape")
    out = []
    for ch in s:
        c = ord(ch)
        if ch.isascii() and (ch.isalnum() or ch in "@*_+-./"):
            out.append(ch)
        elif c < 256:
            out.append("%%%02X" % c)
        else:
            out.append("%%u%04X" % c)
    return "".join(out)


# ---------- 判定 ----------
@bi("IsNull", 1)
def _isnull(a):
    return V(a[0]) is NULL


@bi("IsEmpty", 1)
def _isempty(a):
    return V(a[0]) is EMPTY if not isinstance(a[0], VBObject) else False


@bi("IsArray", 1)
def _isarray(a):
    return isinstance(a[0], VArray)


@bi("IsObject", 1)
def _isobject(a):
    return is_obj(a[0])


@bi("IsNumeric", 1)
def _isnumeric(a):
    v = V(a[0])
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return is_numeric_str(v)
    if v is EMPTY:
        return True
    return False


@bi("IsDate", 1)
def _isdate(a):
    v = V(a[0])
    if isinstance(v, VDate):
        return True
    if isinstance(v, str):
        raise strict_error("IsDate に文字列を渡しています（判定が地域設定で変わります。自前で年・月・日に分けて調べてください）")
    return False


@bi("VarType", 1)
def _vartype(a):
    return vartype(a[0] if not isinstance(a[0], VBObject) else a[0])


@bi("TypeName", 1)
def _typename(a):
    return type_name(a[0])


# ---------- 変換 ----------
def _to_number_strict(v, fname):
    v = V(v)
    if v is NULL:
        raise VBError(94, f"Null の使い方が不正です: {fname}")
    if isinstance(v, str):
        if not is_numeric_str(v):
            raise VBError(13, f"型が一致しません: {fname}(\"{v}\")")
        if "," in v:
            raise strict_error(f"{fname} に桁区切りの , を含む文字列を渡しています（\"{v}\"）")
        return num_from_str(v)
    if isinstance(v, VDate):
        raise strict_error(f"{fname} に日付を渡しています")
    return to_num(v)


@bi("CLng", 1)
def _clng(a):
    n = vb_round(_to_number_strict(a[0], "CLng"))
    if not -2147483648 <= n <= 2147483647:
        raise VBError(6, "オーバーフローしました: CLng")
    return n


@bi("CInt", 1)
def _cint(a):
    n = vb_round(_to_number_strict(a[0], "CInt"))
    if not -32768 <= n <= 32767:
        raise VBError(6, "オーバーフローしました: CInt")
    return n


@bi("CDbl", 1)
def _cdbl(a):
    return float(_to_number_strict(a[0], "CDbl"))


@bi("CSng", 1)
def _csng(a):
    return float(_to_number_strict(a[0], "CSng"))


@bi("CCur", 1)
def _ccur(a):
    return float(_to_number_strict(a[0], "CCur"))


@bi("CByte", 1)
def _cbyte(a):
    n = vb_round(_to_number_strict(a[0], "CByte"))
    if not 0 <= n <= 255:
        raise VBError(6, "オーバーフローしました: CByte")
    return n


@bi("CBool", 1)
def _cbool(a):
    v = V(a[0])
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です: CBool")
    if isinstance(v, str):
        raise strict_error("CBool に文字列を渡しています（\"True\" / \"真\" など地域設定で変わる）")
    return to_bool(v)


@bi("CStr", 1)
def _cstr(a):
    v = V(a[0])
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です: CStr")
    return to_str(v, "CStr で文字列化")


@bi("CDate", 1)
def _cdate(a):
    v = V(a[0])
    if isinstance(v, VDate):
        return v
    if isinstance(v, str):
        raise strict_error("CDate に文字列を渡しています（解釈が地域設定で変わります。DateSerial を使う）")
    if v is NULL:
        raise VBError(94, "Null の使い方が不正です: CDate")
    return VDate.from_serial(to_num(v))


@bi("Int", 1)
def _int(a):
    v = V(a[0])
    if v is NULL:
        return NULL
    n = to_num(v)
    return int(math.floor(n)) if isinstance(n, float) else n


@bi("Fix", 1)
def _fix(a):
    v = V(a[0])
    if v is NULL:
        return NULL
    n = to_num(v)
    return int(n) if isinstance(n, float) else n


@bi("Abs", 1)
def _abs(a):
    v = V(a[0])
    return NULL if v is NULL else abs(to_num(v))


@bi("Sgn", 1)
def _sgn(a):
    n = to_num(V(a[0]))
    return (n > 0) - (n < 0)


@bi("Round", 1, 2)
def _round(a):
    v = V(a[0])
    if v is NULL:
        return NULL
    n = to_num(v)
    d = as_int_arg(a[1], "Round") if len(a) > 1 else 0
    m = 10 ** d
    r = vb_round(n * m) / m
    return int(r) if d == 0 else r


@bi("Rnd", 0, 1)
def _rnd(a):
    return random.random()


# ---------- 配列 ----------
@bi("Array", 0, 9999)
def _array(a):
    vals = [x for x in a]
    return VArray([len(vals) - 1], data=vals)


@bi("UBound", 1, 2)
def _ubound(a):
    arr = a[0]
    if not isinstance(arr, VArray):
        raise VBError(13, "型が一致しません: UBound の引数が配列ではありません")
    d = as_int_arg(a[1], "UBound") if len(a) > 1 else 1
    if d < 1 or d > len(arr.ubs):
        raise VBError(9, "インデックスが有効範囲にありません: UBound の次元")
    return arr.ubs[d - 1]


@bi("LBound", 1, 2)
def _lbound(a):
    if not isinstance(a[0], VArray):
        raise VBError(13, "型が一致しません: LBound の引数が配列ではありません")
    return 0


# ---------- 日付 ----------
def as_date(v, fname):
    v = V(v)
    if isinstance(v, VDate):
        return v.dt
    if v is NULL:
        raise VBError(94, f"Null の使い方が不正です: {fname}")
    if isinstance(v, str):
        raise strict_error(f"{fname} に文字列を渡しています（日付として解釈されるかは地域設定次第）")
    if isinstance(v, (int, float)):
        raise strict_error(f"{fname} に数値を渡しています")
    raise VBError(13, f"型が一致しません: {fname}")


@bi("Date", 0)
def _date(a):
    n = now()
    return VDate(datetime.datetime(n.year, n.month, n.day))


@bi("Now", 0)
def _now(a):
    n = now()
    return VDate(n.replace(microsecond=0))


@bi("Time", 0)
def _time(a):
    n = now()
    return VDate(datetime.datetime(1899, 12, 30, n.hour, n.minute, n.second))


@bi("Timer", 0)
def _timer(a):
    n = datetime.datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second + n.microsecond / 1e6


@bi("Year", 1)
def _year(a):
    return NULL if V(a[0]) is NULL else as_date(a[0], "Year").year


@bi("Month", 1)
def _month(a):
    return NULL if V(a[0]) is NULL else as_date(a[0], "Month").month


@bi("Day", 1)
def _day(a):
    return NULL if V(a[0]) is NULL else as_date(a[0], "Day").day


@bi("Hour", 1)
def _hour(a):
    return as_date(a[0], "Hour").hour


@bi("Minute", 1)
def _minute(a):
    return as_date(a[0], "Minute").minute


@bi("Second", 1)
def _second(a):
    return as_date(a[0], "Second").second


@bi("Weekday", 1, 2)
def _weekday(a):
    d = as_date(a[0], "Weekday")
    first = as_int_arg(a[1], "Weekday") if len(a) > 1 else 1
    w = (d.isoweekday() % 7) + 1    # 日曜 = 1
    return (w - first) % 7 + 1


@bi("DateSerial", 3)
def _dateserial(a):
    y, m, d = (as_int_arg(x, "DateSerial") for x in a)
    if 0 <= y <= 29:
        y += 2000
    elif 30 <= y <= 99:
        y += 1900
    # 月・日のはみ出しは繰り上げる（本物と同じ）
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    base = datetime.datetime(y, m, 1)
    if not 100 <= y <= 9999:
        raise VBError(5, "DateSerial の年が範囲外です")
    return VDate(base + datetime.timedelta(days=d - 1))


@bi("TimeSerial", 3)
def _timeserial(a):
    h, mi, s = (as_int_arg(x, "TimeSerial") for x in a)
    return VDate(datetime.datetime(1899, 12, 30) + datetime.timedelta(hours=h, minutes=mi, seconds=s))


def _add_months(dt, n):
    m = dt.month - 1 + n
    y = dt.year + m // 12
    m = m % 12 + 1
    d = min(dt.day, calendar.monthrange(y, m)[1])
    return dt.replace(year=y, month=m, day=d)


@bi("DateAdd", 3)
def _dateadd(a):
    unit = as_str_arg(a[0], "DateAdd").lower()
    n = as_int_arg(a[1], "DateAdd")
    dt = as_date(a[2], "DateAdd")
    if unit in ("d", "y", "w"):
        return VDate(dt + datetime.timedelta(days=n))
    if unit == "ww":
        return VDate(dt + datetime.timedelta(days=7 * n))
    if unit == "m":
        return VDate(_add_months(dt, n))
    if unit == "q":
        return VDate(_add_months(dt, 3 * n))
    if unit == "yyyy":
        return VDate(_add_months(dt, 12 * n))
    if unit == "h":
        return VDate(dt + datetime.timedelta(hours=n))
    if unit == "n":
        return VDate(dt + datetime.timedelta(minutes=n))
    if unit == "s":
        return VDate(dt + datetime.timedelta(seconds=n))
    raise VBError(5, f"DateAdd の単位が不正です: {unit}")


@bi("DateDiff", 3, 5)
def _datediff(a):
    unit = as_str_arg(a[0], "DateDiff").lower()
    d1 = as_date(a[1], "DateDiff")
    d2 = as_date(a[2], "DateDiff")
    if unit in ("d", "y"):
        return (d2.date() - d1.date()).days
    if unit == "h":
        return int((d2.replace(minute=0, second=0) - d1.replace(minute=0, second=0)).total_seconds() // 3600)
    if unit == "n":
        return int((d2.replace(second=0) - d1.replace(second=0)).total_seconds() // 60)
    if unit == "s":
        return int((d2 - d1).total_seconds())
    if unit == "m":
        return (d2.year - d1.year) * 12 + d2.month - d1.month
    if unit == "yyyy":
        return d2.year - d1.year
    if unit == "ww":
        return ((d2.date() - d1.date()).days) // 7
    raise VBError(5, f"DateDiff の単位が不正です: {unit}")


@bi("FormatDateTime", 1, 2)
def _fdt(a):
    raise strict_error("FormatDateTime は地域設定で書式が変わるので使わない決まりです")


@bi("FormatNumber", 1, 5)
def _fn(a):
    raise strict_error("FormatNumber は地域設定で書式が変わるので使わない決まりです")


@bi("WeekdayName", 1, 3)
def _wdn(a):
    raise strict_error("WeekdayName は地域設定で文字が変わるので使わない決まりです")


@bi("MonthName", 1, 2)
def _mn(a):
    raise strict_error("MonthName は地域設定で文字が変わるので使わない決まりです")


@bi("Eval", 1)
def _eval(a):
    raise strict_error("Eval は使わない決まりです")


@bi("GetRef", 1)
def _getref(a):
    raise strict_error("GetRef は使わない決まりです")


@bi("ScriptEngineMajorVersion", 0)
def _sev(a):
    return 5


@bi("ScriptEngineMinorVersion", 0)
def _sevm(a):
    return 8


@bi("ScriptEngineBuildVersion", 0)
def _sevb(a):
    return 16384


@bi("ScriptEngine", 0)
def _se(a):
    return "VBScript"


@bi("RGB", 3)
def _rgb(a):
    r, g, b = (as_int_arg(x, "RGB") for x in a)
    return r + g * 256 + b * 65536


# ---------- バイト列 ----------
@bi("LenB", 1)
def _lenb(a):
    v = a[0].get() if hasattr(a[0], "get") and not isinstance(a[0], VBObject) else a[0]
    if isinstance(v, VBinary):
        return len(v.b)
    return 2 * len(as_str_arg(v, "LenB"))


@bi("MidB", 2, 3)
def _midb(a):
    v = a[0]
    if not isinstance(v, VBinary):
        raise strict_error("MidB はバイト配列（ADODB.Stream の Read の結果）にだけ使う決まりです")
    st = as_int_arg(a[1], "MidB")
    n = as_int_arg(a[2], "MidB") if len(a) > 2 else len(v.b)
    return VBinary(v.b[st - 1:st - 1 + n])


@bi("AscB", 1)
def _ascb(a):
    v = a[0]
    if not isinstance(v, VBinary):
        raise strict_error("AscB はバイト配列にだけ使う決まりです")
    if not v.b:
        raise VBError(5, "AscB に空のバイト列を渡しています")
    return v.b[0]


class Builtins:
    def __init__(self, host):
        self.host = host

    def has(self, low):
        return low in FUNCS or low == "createobject" or low == "getobject"

    def call(self, interp, low, args):
        if low == "createobject":
            if len(args) != 1:
                raise VBError(450, "引数の数が一致していません: CreateObject")
            return self.host.create_object(to_str(V(args[0])))
        if low == "getobject":
            raise strict_error("GetObject は使わない決まりです")
        lo, hi = ARITY[low]
        if not lo <= len(args) <= hi:
            raise VBError(450, f"引数の数が一致していません: '{low}'")
        return FUNCS[low](args)


BUILTIN_NAMES = set(FUNCS) | {"createobject", "getobject"}


# VBScript に最初から入っている定数（ADO の定数は入っていないので、使う側で Const を書く）
VB_CONSTS = {
    "vbcrlf": "\r\n", "vbcr": "\r", "vblf": "\n", "vbtab": "\t", "vbnullstring": "", "vbnewline": "\r\n",
    "vbnullchar": "\x00",
    "vbempty": 0, "vbnull": 1, "vbinteger": 2, "vblong": 3, "vbsingle": 4, "vbdouble": 5, "vbcurrency": 6,
    "vbdate": 7, "vbstring": 8, "vbobject": 9, "vberror": 10, "vbboolean": 11, "vbvariant": 12, "vbbyte": 17,
    "vbarray": 8192,
    "vbbinarycompare": 0, "vbtextcompare": 1,
    "vbsunday": 1, "vbmonday": 2, "vbtuesday": 3, "vbwednesday": 4, "vbthursday": 5, "vbfriday": 6, "vbsaturday": 7,
    "vbusesystemdayofweek": 0,
    "vbobjecterror": -2147221504, "vbtrue": -1, "vbfalse": 0, "vbusedefault": -2,
}
BUILTIN_NAMES |= set(VB_CONSTS)
