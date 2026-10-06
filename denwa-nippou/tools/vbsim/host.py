"""ASP の組み込みオブジェクトと COM オブジェクトの模擬。"""
import datetime
import os
import re
import shutil
import subprocess
import threading
import time
import urllib.parse

from .values import (VBError, ResponseEnd, strict_error, EMPTY, NULL, NOTHING, VDate, VArray, VBObject, VBinary,
                     to_value, to_str, to_num, to_bool, vb_round, is_obj)
from .interp import Cell
from .sqlbridge import BridgeError


def raw(a):
    if isinstance(a, Cell):
        return a.get()
    return a


def val(a):
    return to_value(raw(a))


def sval(a, what):
    v = val(a)
    if v is NULL:
        raise VBError(94, f"Null の使い方が不正です: {what}")
    return to_str(v, what)


def ival(a, what):
    v = val(a)
    if isinstance(v, str):
        raise strict_error(f"{what} に文字列 \"{v}\" を渡しています")
    return vb_round(to_num(v))


def argn(args, n, what):
    if len(args) < n:
        raise VBError(450, f"引数の数が一致していません: {what}")


def ado_error(number, desc, source="Microsoft Access Database Engine"):
    return VBError(number, desc, source)


# ======================================================================
# Response
# ======================================================================
class CookieOut(VBObject):
    def __init__(self, name):
        self.name = name
        self.value = ""
        self.expires = None
        self.path = "/"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "expires":
            return self.expires or EMPTY
        if n == "path":
            return self.path
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        if n == "expires":
            v = to_value(value)
            if not isinstance(v, VDate):
                raise strict_error("Cookies().Expires には日付を入れてください")
            self.expires = v
            return
        if n == "path":
            self.path = to_str(to_value(value))
            return
        super().vb_set(name, args, value)

    def vb_default_get(self, args):
        return self.value

    def vb_default_set(self, args, value):
        self.value = to_str(to_value(value), "Cookie に保存")


class CookiesOut(VBObject):
    def __init__(self, host):
        self.host = host

    def vb_default_get(self, args):
        argn(args, 1, "Response.Cookies")
        return self.host.cookie_out(sval(args[0], "Cookie の名前"))

    def vb_default_set(self, args, value):
        argn(args, 1, "Response.Cookies")
        self.host.check_header("Response.Cookies")
        self.host.cookie_out(sval(args[0], "Cookie の名前")).vb_default_set([], value)


class ResponseObj(VBObject):
    def __init__(self, host):
        self.h = host

    def type_name(self):
        return "IResponse"

    def vb_get(self, name, args):
        h = self.h
        n = name.lower()
        if n == "write":
            argn(args, 1, "Response.Write")
            v = val(args[0])
            if v is NULL:
                raise strict_error("Response.Write に Null を渡しています（データベースの空欄をそのまま出していないか）")
            h.out_text(to_str(v, "Response.Write で出力"))
            return EMPTY
        if n == "binarywrite":
            v = raw(args[0])
            if not isinstance(v, VBinary):
                raise VBError(13, "型が一致しません: BinaryWrite にはバイト配列を渡します")
            h.out_bin(v.b)
            return EMPTY
        if n == "clear":
            if not h.buffer:
                raise VBError(-2147467259, "Response.Clear はバッファが無効のとき使えません", "Response オブジェクト")
            h.body_text = []
            h.body_bin = []
            h.parts = []
            return EMPTY
        if n == "end":
            raise ResponseEnd()
        if n == "flush":
            h.flushed = True
            return EMPTY
        if n == "redirect":
            argn(args, 1, "Response.Redirect")
            url = sval(args[0], "Response.Redirect")
            h.check_header("Response.Redirect")
            h.status = "302 Object moved"
            h.headers["Location"] = url
            h.parts = []
            h.out_text('<head><title>Object moved</title></head><body><h1>Object Moved</h1></body>')
            raise ResponseEnd()
        if n == "addheader":
            argn(args, 2, "Response.AddHeader")
            h.check_header("Response.AddHeader")
            h.headers[sval(args[0], "AddHeader")] = sval(args[1], "AddHeader")
            return EMPTY
        if n == "appendtolog":
            return EMPTY
        if n == "isclientconnected":
            return True
        if n == "cookies":
            if args:
                return h.cookie_out(sval(args[0], "Cookie の名前"))
            return CookiesOut(h)
        if n in ("buffer", "contenttype", "charset", "codepage", "status", "expires", "cachecontrol", "expiresabsolute"):
            return h.props.get(n, EMPTY)
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        h = self.h
        n = name.lower()
        if n == "cookies":
            argn(args, 1, "Response.Cookies")
            h.check_header("Response.Cookies")
            h.cookie_out(sval(args[0], "Cookie の名前")).vb_default_set([], value)
            return
        v = to_value(value)
        if n == "codepage":
            cp = vb_round(to_num(v))
            if cp != 65001:
                raise strict_error(f"Response.CodePage = {cp}（このシステムは 65001 だけを使う決まり）")
            h.codepage = cp
            h.props[n] = cp
            return
        if n == "charset":
            h.check_header("Response.CharSet")
            h.props[n] = to_str(v)
            return
        if n == "contenttype":
            h.check_header("Response.ContentType")
            h.props[n] = to_str(v)
            return
        if n == "status":
            h.check_header("Response.Status")
            s = to_str(v)
            if not re.match(r"^\d{3} \S", s):
                raise VBError(-2147467259, f"Response.Status の書き方が不正です: \"{s}\"（\"200 OK\" の形）", "Response オブジェクト")
            h.status = s
            return
        if n == "buffer":
            if h.parts:
                raise VBError(-2147467259, "出力の後で Buffer は変えられません", "Response オブジェクト")
            h.buffer = to_bool(v)
            return
        if n in ("expires", "cachecontrol", "expiresabsolute"):
            h.check_header("Response." + name)
            h.props[n] = v
            return
        super().vb_set(name, args, value)


# ======================================================================
# Request
# ======================================================================
class RequestItem(VBObject):
    def __init__(self, values):
        self.values = values

    def type_name(self):
        return "IStringList"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "count":
            return len(self.values)
        if n == "item":
            if not args:
                return self.vb_default_get([])
            i = ival(args[0], "Item")
            if i < 1 or i > len(self.values):
                raise VBError(9, "インデックスが有効範囲にありません")
            return self.values[i - 1]
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        if args:
            return self.vb_get("item", args)
        if not self.values:
            return EMPTY
        return ", ".join(self.values)


class RequestColl(VBObject):
    def __init__(self, pairs, name):
        self.pairs = pairs
        self.name = name

    def type_name(self):
        return "IRequestDictionary"

    def _vals(self, key):
        k = key.lower()
        return [v for (kk, v) in self.pairs if kk.lower() == k]

    def vb_default_get(self, args):
        if not args:
            raise strict_error(f"Request.{self.name} をそのまま文字列として使っています（名前を指定してください）")
        k = val(args[0])
        if isinstance(k, int):
            keys = self._keys()
            if k < 1 or k > len(keys):
                raise VBError(9, "インデックスが有効範囲にありません")
            return RequestItem(self._vals(keys[k - 1]))
        return RequestItem(self._vals(to_str(k)))

    def _keys(self):
        seen = []
        for k, _ in self.pairs:
            if k not in seen:
                seen.append(k)
        return seen

    def vb_get(self, name, args):
        n = name.lower()
        if n == "count":
            return len(self._keys())
        if n == "item":
            return self.vb_default_get(args)
        if n == "key":
            return self._keys()[ival(args[0], "Key") - 1]
        return super().vb_get(name, args)

    def vb_enum(self):
        return list(self._keys())


class SimpleColl(VBObject):
    """ServerVariables / Cookies（値はただの文字列）。"""

    def __init__(self, d, name):
        self.d = {k.upper(): v for k, v in d.items()}
        self.name = name

    def vb_default_get(self, args):
        argn(args, 1, f"Request.{self.name}")
        k = sval(args[0], self.name).upper()
        return self.d.get(k, "")

    def vb_get(self, name, args):
        if name.lower() == "item":
            return self.vb_default_get(args)
        if name.lower() == "count":
            return len(self.d)
        return super().vb_get(name, args)

    def vb_enum(self):
        return list(self.d)


class RequestObj(VBObject):
    def __init__(self, host):
        self.h = host
        r = host.req
        self.form = RequestColl(r["form"], "Form")
        self.qs = RequestColl(r["query"], "QueryString")
        self.cookies = SimpleColl(r["cookies"], "Cookies")
        self.sv = SimpleColl(r["server_vars"], "ServerVariables")

    def type_name(self):
        return "IRequest"

    def vb_get(self, name, args):
        n = name.lower()
        m = {"form": self.form, "querystring": self.qs, "cookies": self.cookies, "servervariables": self.sv}
        if n in m:
            return m[n].vb_default_get(args) if args else m[n]
        if n == "totalbytes":
            return len(self.h.req.get("body", b""))
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        raise strict_error("Request(\"名前\") は QueryString・Form・Cookies などを順に探すので使わない決まりです（Request.Form / Request.QueryString と明示する）")


# ======================================================================
# Server / ASPError / Session
# ======================================================================
class ASPErrorObj(VBObject):
    def __init__(self, e, vpath_of):
        self.e = e
        self.vpath_of = vpath_of

    def type_name(self):
        return "IASPError"

    def vb_get(self, name, args):
        n = name.lower()
        e = self.e
        if e is None:
            return {"number": 0, "line": 0, "column": -1}.get(n, "")
        if n == "number":
            # ASPError.Number は HRESULT（VBScript の実行時エラー 9 なら 800A0009）。Err.Number（9）とは違う。
            if 0 < e.number < 65536:
                return -2146828288 + e.number
            return e.number
        if n == "description":
            return e.description
        if n == "aspdescription":
            return ""
        if n == "category":
            return e.source
        if n == "source":
            return ""
        if n == "file":
            return self.vpath_of(e.where[0]) if e.where else ""
        if n == "line":
            return e.where[1] if e.where else 0
        if n == "column":
            return -1
        if n == "aspcode":
            return ""
        return super().vb_get(name, args)


class ServerObj(VBObject):
    def __init__(self, host):
        self.h = host
        self.timeout = 90

    def type_name(self):
        return "IServer"

    def vb_get(self, name, args):
        h = self.h
        n = name.lower()
        if n == "mappath":
            argn(args, 1, "Server.MapPath")
            return h.sim.map_path(sval(args[0], "MapPath"), h.req["path"])
        if n == "htmlencode":
            argn(args, 1, "Server.HTMLEncode")
            v = val(args[0])
            if v is NULL:
                raise strict_error("Server.HTMLEncode に Null を渡しています")
            s = to_str(v, "HTMLEncode")
            out = []
            for ch in s:
                if ch == "&":
                    out.append("&amp;")
                elif ch == "<":
                    out.append("&lt;")
                elif ch == ">":
                    out.append("&gt;")
                elif ch == '"':
                    out.append("&quot;")
                elif 0xA0 <= ord(ch) <= 0xFF:
                    out.append(f"&#{ord(ch)};")
                else:
                    out.append(ch)
            return "".join(out)
        if n == "urlencode":
            argn(args, 1, "Server.URLEncode")
            s = sval(args[0], "URLEncode")
            return urllib.parse.quote_plus(s.encode("utf-8"), safe="")
        if n == "createobject":
            argn(args, 1, "Server.CreateObject")
            return h.create_object(sval(args[0], "CreateObject"))
        if n == "getlasterror":
            return ASPErrorObj(h.last_error, h.sim.vpath_of_display)
        if n == "scripttimeout":
            return self.timeout
        if n in ("execute", "transfer"):
            raise strict_error(f"Server.{name} は使わない決まりです")
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        if name.lower() == "scripttimeout":
            self.timeout = vb_round(to_num(to_value(value)))
            return
        super().vb_set(name, args, value)


class Forbidden(VBObject):
    def __init__(self, what, why):
        self.what, self.why = what, why

    def _no(self):
        raise strict_error(f"{self.what} は使わない決まりです（{self.why}）")

    def vb_get(self, name, args):
        self._no()

    def vb_set(self, name, args, value):
        self._no()

    def vb_default_get(self, args):
        self._no()

    def vb_default_set(self, args, value):
        self._no()


# ======================================================================
# Scripting.Dictionary
# ======================================================================
class Dictionary(VBObject):
    def __init__(self):
        self.d = {}
        self.mode = 0

    def type_name(self):
        return "Dictionary"

    def _key(self, k):
        k = raw(k)
        if isinstance(k, VBObject):
            raise strict_error("Dictionary のキーにオブジェクトを渡しています（rs(\"列\") ではなく値を渡す）")
        if not isinstance(k, str):
            raise strict_error(f"Dictionary のキーは文字列にする決まりです（{k!r}。数値と文字列の 1 と \"1\" は別のキーになるため CStr で揃える）")
        return k.lower() if self.mode == 1 else k

    def vb_get(self, name, args):
        n = name.lower()
        if n == "add":
            argn(args, 2, "Dictionary.Add")
            k = self._key(args[0])
            if k in self.d:
                raise VBError(457, f"このキーは既にこのコレクションの要素に割り当てられています: {k}")
            v = raw(args[1])
            if isinstance(v, VBObject) and getattr(v, "is_live_field", False):
                raise strict_error("Dictionary.Add に rs(\"列\") をそのまま渡しています（Field オブジェクトが入り、次の行に進むと値が変わります）")
            self.d[k] = v
            return EMPTY
        if n == "exists":
            return self._key(args[0]) in self.d
        if n == "item":
            return self.vb_default_get(args)
        if n == "keys":
            ks = list(self.d)
            return VArray([len(ks) - 1], data=ks)
        if n == "items":
            vs = list(self.d.values())
            return VArray([len(vs) - 1], data=vs)
        if n == "count":
            return len(self.d)
        if n == "remove":
            k = self._key(args[0])
            if k not in self.d:
                raise VBError(32811, f"要素が見つかりません: {k}")
            del self.d[k]
            return EMPTY
        if n == "removeall":
            self.d.clear()
            return EMPTY
        if n == "comparemode":
            return self.mode
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        if n == "comparemode":
            if self.d:
                raise VBError(5, "要素がある Dictionary の CompareMode は変えられません")
            self.mode = vb_round(to_num(to_value(value)))
            return
        if n == "item":
            self.vb_default_set(args, value)
            return
        super().vb_set(name, args, value)

    def vb_default_get(self, args):
        argn(args, 1, "Dictionary.Item")
        k = self._key(args[0])
        if k not in self.d:
            # 本物の Dictionary は「無いキーを読むと空の要素を足してしまう」
            self.d[k] = EMPTY
        return self.d[k]

    def vb_default_set(self, args, value):
        argn(args, 1, "Dictionary.Item")
        self.d[self._key(args[0])] = value

    def vb_enum(self):
        return list(self.d)


# ======================================================================
# Scripting.FileSystemObject
# ======================================================================
def fs_date(ts):
    return VDate(datetime.datetime.fromtimestamp(ts).replace(microsecond=0))


class FileObj(VBObject):
    def __init__(self, fso, local):
        self.fso, self.local = fso, local

    def type_name(self):
        return "File"

    def vb_get(self, name, args):
        n = name.lower()
        if not os.path.exists(self.local):
            raise VBError(53, "ファイルが見つかりません")
        st = os.stat(self.local)
        if n == "name":
            return os.path.basename(self.local)
        if n == "path":
            return self.fso.w.to_win(self.local)
        if n == "size":
            return st.st_size
        if n == "datelastmodified":
            return fs_date(st.st_mtime)
        if n == "datecreated":
            return fs_date(getattr(st, "st_ctime", st.st_mtime))
        if n == "delete":
            self.fso.check_write(self.local)
            os.remove(self.local)
            return EMPTY
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        return self.vb_get("path", [])


class ItemsColl(VBObject):
    def __init__(self, items):
        self.items = items

    def vb_get(self, name, args):
        if name.lower() == "count":
            return len(self.items)
        return super().vb_get(name, args)

    def vb_enum(self):
        return list(self.items)


class FolderObj(VBObject):
    def __init__(self, fso, local):
        self.fso, self.local = fso, local

    def type_name(self):
        return "Folder"

    def vb_get(self, name, args):
        n = name.lower()
        if not os.path.isdir(self.local):
            raise VBError(76, "パスが見つかりません")
        if n == "name":
            return os.path.basename(self.local)
        if n == "path":
            return self.fso.w.to_win(self.local)
        if n == "files":
            return ItemsColl([FileObj(self.fso, os.path.join(self.local, x)) for x in sorted(os.listdir(self.local))
                              if os.path.isfile(os.path.join(self.local, x))])
        if n == "subfolders":
            return ItemsColl([FolderObj(self.fso, os.path.join(self.local, x)) for x in sorted(os.listdir(self.local))
                              if os.path.isdir(os.path.join(self.local, x))])
        if n == "datelastmodified":
            return fs_date(os.stat(self.local).st_mtime)
        if n == "datecreated":
            return fs_date(os.stat(self.local).st_ctime)
        if n == "delete":
            self.fso.check_write(self.local)
            shutil.rmtree(self.local)
            return EMPTY
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        return self.vb_get("path", [])


class TextStream(VBObject):
    def __init__(self, fso, local, mode, enc):
        self.fso, self.local, self.mode, self.enc = fso, local, mode, enc
        self.buf = []
        self.data = None
        self.pos = 0
        if mode == 1:
            raw_ = open(local, "rb").read()
            # 本物と同じく、指定された文字コードで（合っていなくても）そのまま読む
            if enc == "utf-16-le" and raw_.startswith(b"\xff\xfe"):
                raw_ = raw_[2:]
            self.data = raw_.decode(enc, errors="replace")

    def type_name(self):
        return "TextStream"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "readall":
            s = self.data[self.pos:]
            self.pos = len(self.data)
            return s
        if n == "readline":
            j = self.data.find("\n", self.pos)
            if j < 0:
                s = self.data[self.pos:]
                self.pos = len(self.data)
            else:
                s = self.data[self.pos:j]
                self.pos = j + 1
            return s.rstrip("\r")
        if n == "atendofstream":
            return self.pos >= len(self.data)
        if n in ("write", "writeline"):
            s = sval(args[0], "TextStream.Write") if args else ""
            self.buf.append(s + ("\r\n" if n == "writeline" else ""))
            return EMPTY
        if n == "close":
            if self.mode in (2, 8):
                text = "".join(self.buf)
                try:
                    b = text.encode(self.enc)
                except UnicodeEncodeError:
                    raise strict_error("FileSystemObject の文字コード（Shift_JIS）で表せない文字を書き込もうとしました（本物では ? に化けます）")
                if self.enc == "utf-16-le":
                    b = b"\xff\xfe" + b
                with open(self.local, "ab" if self.mode == 8 else "wb") as fh:
                    fh.write(b)
            return EMPTY
        return super().vb_get(name, args)


class FSO(VBObject):
    def __init__(self, host):
        self.h = host
        self.w = host.sim.winfs

    def type_name(self):
        return "FileSystemObject"

    def L(self, a, what):
        p = sval(a, what)
        if not self.w.is_abs(p):
            raise strict_error(f"{what} に相対パス \"{p}\" を渡しています（IIS の中では基準の場所が C:\\Windows\\System32\\inetsrv になります）")
        return self.w.to_local(p)

    def check_write(self, local):
        for ro in self.h.sim.readonly_dirs:
            if os.path.abspath(local).startswith(os.path.abspath(ro)):
                raise VBError(70, "書き込みできません（アクセス許可がありません）", "Microsoft VBScript 実行時エラー")

    def vb_get(self, name, args):
        n = name.lower()
        if n == "fileexists":
            return os.path.isfile(self.L(args[0], "FileExists"))
        if n == "folderexists":
            return os.path.isdir(self.L(args[0], "FolderExists"))
        if n == "createfolder":
            p = self.L(args[0], "CreateFolder")
            if os.path.exists(p):
                raise VBError(58, "既に同名のファイルが存在しています")
            if not os.path.isdir(os.path.dirname(p)):
                raise VBError(76, "パスが見つかりません")
            self.check_write(p)
            os.mkdir(p)
            return FolderObj(self, p)
        if n == "getfolder":
            p = self.L(args[0], "GetFolder")
            if not os.path.isdir(p):
                raise VBError(76, "パスが見つかりません")
            return FolderObj(self, p)
        if n == "getfile":
            p = self.L(args[0], "GetFile")
            if not os.path.isfile(p):
                raise VBError(53, "ファイルが見つかりません")
            return FileObj(self, p)
        if n == "deletefile":
            p = self.L(args[0], "DeleteFile")
            if not os.path.isfile(p):
                raise VBError(53, "ファイルが見つかりません")
            self.check_write(p)
            os.remove(p)
            return EMPTY
        if n == "deletefolder":
            p = self.L(args[0], "DeleteFolder")
            if not os.path.isdir(p):
                raise VBError(76, "パスが見つかりません")
            self.check_write(p)
            shutil.rmtree(p)
            return EMPTY
        if n == "buildpath":
            a, b = sval(args[0], "BuildPath"), sval(args[1], "BuildPath")
            if a.endswith("\\") or a.endswith(":"):
                return a + b
            return a + "\\" + b
        if n == "getparentfoldername":
            p = sval(args[0], "GetParentFolderName").rstrip("\\")
            i = p.rfind("\\")
            return p[:i] if i > 0 else ""
        if n == "getfilename":
            p = sval(args[0], "GetFileName").rstrip("\\")
            return p[p.rfind("\\") + 1:]
        if n == "getbasename":
            p = sval(args[0], "GetBaseName").rstrip("\\")
            b = p[p.rfind("\\") + 1:]
            return b[:b.rfind(".")] if "." in b else b
        if n == "getextensionname":
            p = sval(args[0], "GetExtensionName")
            b = p[p.rfind("\\") + 1:]
            return b[b.rfind(".") + 1:] if "." in b else ""
        if n == "getspecialfolder":
            k = ival(args[0], "GetSpecialFolder")
            if k == 2:
                return FolderObj(self, self.w.to_local(self.h.sim.env["TEMP"]))
            raise strict_error("GetSpecialFolder は 2（一時フォルダ）だけ使う決まりです")
        if n == "gettempname":
            import random
            return "rad%05X.tmp" % random.randint(0, 0xFFFFF)
        if n == "createtextfile":
            p = self.L(args[0], "CreateTextFile")
            over = to_bool(val(args[1])) if len(args) > 1 else True
            uni = to_bool(val(args[2])) if len(args) > 2 else False
            if os.path.exists(p) and not over:
                raise VBError(58, "既に同名のファイルが存在しています")
            self.check_write(p)
            open(p, "wb").close()
            return TextStream(self, p, 2, "utf-16-le" if uni else "cp932")
        if n == "opentextfile":
            p = self.L(args[0], "OpenTextFile")
            mode = ival(args[1], "OpenTextFile") if len(args) > 1 else 1
            create = to_bool(val(args[2])) if len(args) > 2 else False
            fmt = ival(args[3], "OpenTextFile") if len(args) > 3 else 0
            enc = {0: "cp932", -1: "utf-16-le", -2: "cp932"}[fmt]
            if not os.path.exists(p):
                if mode == 1 or not create:
                    raise VBError(53, "ファイルが見つかりません")
                open(p, "wb").close()
            if mode != 1:
                self.check_write(p)
            return TextStream(self, p, mode, enc)
        return super().vb_get(name, args)


# ======================================================================
# ADODB.Stream
# ======================================================================
CHARSETS = {"utf-8": "utf-8", "unicode": "utf-16-le", "shift_jis": "cp932", "us-ascii": "ascii", "x-sjis": "cp932"}


class AdoStream(VBObject):
    def __init__(self, host):
        self.h = host
        self.type = 2
        self.charset = "unicode"
        self.buf = bytearray()
        self.pos = 0
        self.open_ = False

    def type_name(self):
        return "Stream"

    def _need_open(self):
        if not self.open_:
            raise ado_error(-2146824584, "オブジェクトが閉じている場合は、操作は許可されません。", "ADODB.Stream")

    def _enc(self):
        cs = self.charset.lower()
        if cs not in CHARSETS:
            raise strict_error(f"ADODB.Stream の Charset \"{self.charset}\" は想定外です")
        return CHARSETS[cs]

    def vb_get(self, name, args):
        n = name.lower()
        if n == "open":
            self.open_ = True
            self.buf = bytearray()
            self.pos = 0
            return EMPTY
        if n == "close":
            self.open_ = False
            return EMPTY
        if n == "type":
            return self.type
        if n == "charset":
            return self.charset
        if n == "size":
            return len(self.buf)
        if n == "position":
            return self.pos
        if n == "state":
            return 1 if self.open_ else 0
        self._need_open()
        if n == "loadfromfile":
            p = self.h.fso().L(args[0], "ADODB.Stream.LoadFromFile")
            if not os.path.isfile(p):
                raise ado_error(-2146825286, "ファイルを開けませんでした。", "ADODB.Stream")
            self.buf = bytearray(open(p, "rb").read())
            self.pos = 0
            return EMPTY
        if n == "savetofile":
            p = self.h.fso().L(args[0], "ADODB.Stream.SaveToFile")
            opt = ival(args[1], "SaveToFile") if len(args) > 1 else 1
            if opt == 1 and os.path.exists(p):
                raise ado_error(-2146825284, "ファイルへの書き込みに失敗しました。（既にあります）", "ADODB.Stream")
            if not os.path.isdir(os.path.dirname(p)):
                raise ado_error(-2146825284, "ファイルへの書き込みに失敗しました。（フォルダがありません）", "ADODB.Stream")
            self.h.fso().check_write(p)
            open(p, "wb").write(bytes(self.buf))
            return EMPTY
        if n == "readtext":
            if self.type != 2:
                raise ado_error(-2146825287, "操作は、この状況では許可されていません。（Type が文字ではありません）", "ADODB.Stream")
            enc = self._enc()
            data = bytes(self.buf[self.pos:])
            if self.pos == 0:
                if enc == "utf-8" and data.startswith(b"\xef\xbb\xbf"):
                    data = data[3:]
                if enc == "utf-16-le" and data.startswith(b"\xff\xfe"):
                    data = data[2:]
            self.pos = len(self.buf)
            return data.decode(enc, errors="replace")
        if n == "writetext":
            if self.type != 2:
                raise ado_error(-2146825287, "操作は、この状況では許可されていません。（Type が文字ではありません）", "ADODB.Stream")
            s = sval(args[0], "WriteText")
            if len(args) > 1 and ival(args[1], "WriteText") == 1:
                s += "\r\n"
            enc = self._enc()
            try:
                b = s.encode(enc)
            except UnicodeEncodeError:
                raise strict_error(f"ADODB.Stream の文字コード {self.charset} で表せない文字があります")
            if self.pos == 0 and len(self.buf) == 0:
                # 本物の ADODB.Stream は utf-8 / unicode で書くと先頭に BOM を付ける
                if enc == "utf-8":
                    b = b"\xef\xbb\xbf" + b
                elif enc == "utf-16-le":
                    b = b"\xff\xfe" + b
            self.buf[self.pos:self.pos + len(b)] = b
            self.pos += len(b)
            return EMPTY
        if n == "read":
            if self.type != 1:
                raise ado_error(-2146825287, "操作は、この状況では許可されていません。（Type がバイナリではありません）", "ADODB.Stream")
            cnt = ival(args[0], "Read") if args else -1
            end = len(self.buf) if cnt < 0 else self.pos + cnt
            out = bytes(self.buf[self.pos:end])
            self.pos = min(end, len(self.buf))
            return VBinary(out)
        if n == "write":
            v = raw(args[0])
            if not isinstance(v, VBinary):
                raise VBError(13, "型が一致しません: Stream.Write")
            self.buf[self.pos:self.pos + len(v.b)] = v.b
            self.pos += len(v.b)
            return EMPTY
        if n == "seteos":
            del self.buf[self.pos:]
            return EMPTY
        if n == "flush":
            return EMPTY
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        v = to_value(value)
        if n == "type":
            if self.open_ and self.pos != 0:
                raise ado_error(-2146825287, "Position が 0 でないときは Type を変えられません", "ADODB.Stream")
            self.type = vb_round(to_num(v))
            return
        if n == "charset":
            if self.open_ and self.pos != 0:
                raise ado_error(-2146825287, "Position が 0 でないときは Charset を変えられません", "ADODB.Stream")
            self.charset = to_str(v)
            return
        if n == "position":
            self.pos = vb_round(to_num(v))
            return
        if n in ("mode", "lineseparator"):
            return
        super().vb_set(name, args, value)


# ======================================================================
# ADO（データベース）
# ======================================================================
AD_TYPES_STR = {8, 129, 130, 200, 201, 202, 203}


def count_placeholders(sql):
    n, i, q = 0, 0, None
    while i < len(sql):
        c = sql[i]
        if q:
            if c == q:
                q = None
        elif c in "'\"":
            q = c
        elif c == "[":
            j = sql.find("]", i)
            i = j if j > 0 else i
        elif c == "?":
            n += 1
        i += 1
    return n


class AdoParameter(VBObject):
    def __init__(self, name, typ, direction, size, value):
        self.name, self.type, self.direction, self.size, self.value = name, typ, direction, size, value

    def type_name(self):
        return "Parameter"

    def vb_get(self, name, args):
        n = name.lower()
        if n in ("name", "type", "direction", "size", "value"):
            return getattr(self, n)
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        if n in ("name", "type", "direction", "size", "value"):
            setattr(self, n, to_value(value))
            return
        super().vb_set(name, args, value)

    def vb_default_get(self, args):
        return self.value


class AdoParameters(VBObject):
    def __init__(self):
        self.items = []

    def vb_get(self, name, args):
        n = name.lower()
        if n == "append":
            p = raw(args[0])
            if not isinstance(p, AdoParameter):
                raise VBError(13, "Parameters.Append にはパラメータを渡します")
            self.items.append(p)
            return EMPTY
        if n == "count":
            return len(self.items)
        if n == "item":
            return self.vb_default_get(args)
        if n == "refresh":
            raise strict_error("Parameters.Refresh は使わない決まりです")
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        k = val(args[0])
        if isinstance(k, int):
            return self.items[k]
        for p in self.items:
            if p.name.lower() == to_str(k).lower():
                return p
        raise ado_error(-2146825023, "要求された名前、または序数に対応する項目がコレクションに見つかりません。", "ADODB.Parameters")

    def vb_enum(self):
        return list(self.items)


class AdoField(VBObject):
    is_live_field = True

    def __init__(self, rs, i):
        self.rs, self.i = rs, i

    def type_name(self):
        return "Field"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "value":
            return self.rs.cur(self.i)
        if n == "name":
            return self.rs.cols[self.i]["name"]
        if n == "type":
            return {"int": 3, "bigint": 5, "dbl": 5, "bool": 11, "date": 7, "str": 202}.get(self.rs.cols[self.i]["t"], 202)
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        if args:
            raise VBError(450, "Field の既定のプロパティに引数は付きません")
        return self.rs.cur(self.i)


class AdoFields(VBObject):
    def __init__(self, rs):
        self.rs = rs

    def vb_get(self, name, args):
        n = name.lower()
        if n == "count":
            self.rs.check_open()
            return len(self.rs.cols)
        if n == "item":
            return self.vb_default_get(args)
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        argn(args, 1, "Fields")
        return self.rs.field(val(args[0]))

    def vb_enum(self):
        return [AdoField(self.rs, i) for i in range(len(self.rs.cols))]


def conv_cell(v, t):
    if v is None:
        return NULL
    if t == "int":
        return int(v)
    if t in ("bigint", "dbl"):
        # ACE が集計結果を何型で返すかは確かめていない。Double として返し、ASP 側に CLng 等での変換を強制する。
        return float(v)
    if t == "bool":
        return bool(v)
    if t == "date":
        return VDate(datetime.datetime.fromisoformat(v))
    return str(v)


class AdoRecordset(VBObject):
    def __init__(self, cols, rows):
        self.cols = cols
        self.rows = rows
        self.pos = 0
        self.open_ = True

    def type_name(self):
        return "Recordset"

    def check_open(self):
        if not self.open_:
            raise ado_error(-2146824584, "オブジェクトが閉じている場合は、操作は許可されません。", "ADODB.Recordset")

    def cur(self, i):
        self.check_open()
        if self.pos >= len(self.rows):
            raise ado_error(-2146825267, "BOF と EOF のいずれかが True になっているか、または現在のレコードが削除されています。要求された操作には、現在のレコードが必要です。", "ADODB.Field")
        return conv_cell(self.rows[self.pos][i], self.cols[i]["t"])

    def field(self, k):
        self.check_open()
        if isinstance(k, int):
            if 0 <= k < len(self.cols):
                return AdoField(self, k)
        elif isinstance(k, str):
            for i, c in enumerate(self.cols):
                if c["name"].lower() == k.lower():
                    return AdoField(self, i)
        raise ado_error(-2146825023, f"要求された名前、または序数に対応する項目がコレクションに見つかりません。（{k}）", "ADODB.Recordset")

    def vb_get(self, name, args):
        n = name.lower()
        if n == "state":
            return 1 if self.open_ else 0
        if n == "close":
            self.check_open()
            self.open_ = False
            return EMPTY
        self.check_open()
        if n == "eof":
            return self.pos >= len(self.rows)
        if n == "bof":
            return len(self.rows) == 0
        if n == "movenext":
            if self.pos >= len(self.rows):
                raise ado_error(-2146825267, "BOF と EOF のいずれかが True になっています（EOF で MoveNext）。", "ADODB.Recordset")
            self.pos += 1
            return EMPTY
        if n == "movefirst":
            if self.pos != 0:
                raise ado_error(-2147217884, "行セットは逆方向のフェッチをサポートしていません。（前向き専用カーソル）", "ADODB.Recordset")
            return EMPTY
        if n == "fields":
            if args:
                return self.field(val(args[0]))
            return AdoFields(self)
        if n == "recordcount":
            return -1
        if n == "getrows":
            if self.pos >= len(self.rows):
                raise ado_error(-2146825267, "BOF と EOF のいずれかが True になっています（GetRows）。", "ADODB.Recordset")
            rest = self.rows[self.pos:]
            arr = VArray([len(self.cols) - 1, len(rest) - 1])
            for r, row in enumerate(rest):
                for c in range(len(self.cols)):
                    arr.set([c, r], conv_cell(row[c], self.cols[c]["t"]))
            self.pos = len(self.rows)
            return arr
        return super().vb_get(name, args)

    def vb_default_get(self, args):
        argn(args, 1, "Recordset の既定（Fields）")
        return self.field(val(args[0]))


class AdoConnection(VBObject):
    def __init__(self, host):
        self.h = host
        self.state = 0
        self.connstr = ""
        self.tx = 0
        self.timeout = 15

    def type_name(self):
        return "Connection"

    def bridge(self):
        if self.state != 1:
            raise ado_error(-2146824584, "オブジェクトが閉じている場合は、操作は許可されません。", "ADODB.Connection")
        return self.h.sim.bridge

    def vb_get(self, name, args):
        n = name.lower()
        if n == "open":
            if args:
                self.connstr = sval(args[0], "Connection.Open")
            self._open()
            return EMPTY
        if n == "close":
            if self.state != 1:
                raise ado_error(-2146824584, "オブジェクトが閉じている場合は、操作は許可されません。", "ADODB.Connection")
            if self.tx:
                self.h.sim.bridge.rollback()
                self.tx = 0
            self.state = 0
            return EMPTY
        if n == "state":
            return self.state
        if n == "connectionstring":
            return self.connstr
        if n == "begintrans":
            b = self.bridge()
            if self.tx:
                raise ado_error(-2147168227, "入れ子のトランザクションは使わない決まりです", "ADODB.Connection")
            b.begin()
            self.tx = 1
            return 1
        if n in ("committrans", "rollbacktrans"):
            b = self.bridge()
            if not self.tx:
                raise ado_error(-2147168242, "トランザクションが開始されていません。", "ADODB.Connection")
            b.commit() if n == "committrans" else b.rollback()
            self.tx = 0
            return EMPTY
        if n == "execute":
            sql = sval(args[0], "Connection.Execute")
            return run_sql(self, sql, [], args[1] if len(args) > 1 else None)
        if n == "openschema":
            kind = ival(args[0], "OpenSchema")
            if kind != 4:
                raise strict_error("OpenSchema は 4（列の一覧）だけ使う決まりです")
            crit = raw(args[1]) if len(args) > 1 else None
            table = None
            col = None
            if isinstance(crit, VArray):
                items = crit.items()
                if len(items) > 2 and isinstance(items[2], str):
                    table = items[2]
                if len(items) > 3 and isinstance(items[3], str):
                    col = items[3]
            if not table:
                raise strict_error("OpenSchema(4) には表の名前を指定する決まりです")
            b = self.bridge()
            try:
                cols, _ = b.query(f"SELECT * FROM [{table}] WHERE 1=0", [])
                names = [c["name"] for c in cols]
            except BridgeError:
                names = []
            rows = [[table, nm] for nm in names if col is None or nm.lower() == col.lower()]
            return AdoRecordset([{"name": "TABLE_NAME", "t": "str"}, {"name": "COLUMN_NAME", "t": "str"}], rows)
        if n == "commandtimeout":
            return self.timeout
        if n == "provider":
            m = re.search(r"provider\s*=\s*([^;]+)", self.connstr, re.I)
            return m.group(1).strip() if m else "MSDASQL"
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        if n == "connectionstring":
            self.connstr = to_str(to_value(value))
            return
        if n == "commandtimeout":
            self.timeout = vb_round(to_num(to_value(value)))
            return
        if n in ("mode", "cursorlocation"):
            return
        super().vb_set(name, args, value)

    def _open(self):
        if self.state == 1:
            raise ado_error(-2146824583, "オブジェクトが開いている場合は、操作は許可されません。", "ADODB.Connection")
        kv = {}
        for part in self.connstr.split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                kv[k.strip().lower()] = v.strip()
        prov = kv.get("provider", "").lower()
        if prov not in self.h.sim.providers:
            raise ado_error(-2146824582, "プロバイダーが見つかりません。正しくインストールされていない可能性があります。", "ADODB.Connection")
        ds = kv.get("data source", "")
        w = self.h.sim.winfs
        if not w.is_abs(ds):
            raise ado_error(-2147467259, f"ファイル 'C:\\Windows\\System32\\inetsrv\\{ds}' が見つかりませんでした。", "Microsoft Access Database Engine")
        local = w.to_local(ds)
        if not os.path.isfile(local):
            raise ado_error(-2147467259, f"ファイル '{ds}' が見つかりませんでした。", "Microsoft Access Database Engine")
        folder = os.path.dirname(local)
        for ro in self.h.sim.readonly_dirs:
            if os.path.abspath(folder).startswith(os.path.abspath(ro)):
                self.h.sim.db_readonly = True
        self.h.sim.bridge.open(local)
        self.state = 1


def bridge_params(params):
    out = []
    for p in params:
        v = to_value(p.value)
        t = p.type
        if v is NULL or v is EMPTY:
            out.append({"t": "null", "v": None})
            continue
        if t == 3:
            if not isinstance(v, int) or isinstance(v, bool):
                raise strict_error(f"adInteger のパラメータに {type(v).__name__} の値 {v!r} を渡しています")
            if not -2147483648 <= v <= 2147483647:
                raise VBError(6, "オーバーフローしました（adInteger）")
            out.append({"t": "int", "v": v})
        elif t == 5:
            out.append({"t": "dbl", "v": float(to_num(v))})
        elif t in (7, 133, 135):
            if not isinstance(v, VDate):
                raise strict_error(f"adDate のパラメータに日付以外 {v!r} を渡しています")
            out.append({"t": "date", "v": v.dt.replace(microsecond=0).isoformat()})
        elif t == 11:
            if not isinstance(v, bool):
                raise strict_error(f"adBoolean のパラメータに真偽値以外 {v!r} を渡しています")
            out.append({"t": "bool", "v": v})
        elif t in AD_TYPES_STR:
            if not isinstance(v, str):
                raise strict_error(f"文字列型のパラメータに {type(v).__name__} の値を渡しています")
            if t in (200, 202, 129, 130) and len(v) > (p.size or 0):
                raise ado_error(-2146824867, f"パラメータの Size（{p.size}）より長い文字列（{len(v)} 文字）を渡しています。", "ADODB.Command")
            if v == "":
                raise strict_error("空文字をデータベースに保存しようとしています（空欄は Null で保存する決まり）")
            out.append({"t": "str", "v": v})
        else:
            raise strict_error(f"想定外のパラメータ型 {t}")
    return out


def run_sql(conn, sql, params, affected_ref):
    b = conn.bridge()
    sql_check = conn.h.sim.sql_hook
    if sql_check:
        sql_check(sql)
    need = count_placeholders(sql)
    if need != len(params):
        raise ado_error(-2147217904, f"1 つ以上の必要なパラメーターの値が設定されていません。（? が {need} 個、渡した値が {len(params)} 個）")
    bp = bridge_params(params)
    head = sql.lstrip().split(None, 1)[0].upper() if sql.strip() else ""
    try:
        if head == "SELECT":
            cols, rows = b.query(sql, bp)
            conn.h.sim.stats["select"] += 1
            return AdoRecordset(cols, rows)
        if conn.h.sim.db_readonly:
            raise ado_error(-2147467259, "更新可能なクエリであることが必要です。（データベースのフォルダに書き込めません）")
        n = b.execute(sql, bp)
        conn.h.sim.stats["write"] += 1
        if isinstance(affected_ref, Cell):
            affected_ref.set(n)
        rs = AdoRecordset([], [])
        rs.open_ = False
        return rs
    except BridgeError as e:
        raise ado_error(-2147217900, str(e))


class AdoCommand(VBObject):
    def __init__(self, host):
        self.h = host
        self.conn = None
        self.text = ""
        self.ctype = 1
        self.params = AdoParameters()

    def type_name(self):
        return "Command"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "createparameter":
            nm = sval(args[0], "CreateParameter") if args else ""
            typ = ival(args[1], "CreateParameter") if len(args) > 1 else 0
            d = ival(args[2], "CreateParameter") if len(args) > 2 else 1
            size = ival(args[3], "CreateParameter") if len(args) > 3 else 0
            v = raw(args[4]) if len(args) > 4 else EMPTY
            if isinstance(v, VBObject):
                v = to_value(v)
            if typ in (200, 202, 129, 130) and size <= 0:
                raise ado_error(-2146824867, "文字列パラメータの Size が 0 です（ADO の「パラメータ オブジェクトの定義が不正」）", "ADODB.Command")
            return AdoParameter(nm, typ, d, size, v)
        if n == "parameters":
            if args:
                return self.params.vb_default_get(args)
            return self.params
        if n == "execute":
            if self.conn is None:
                raise ado_error(-2146824594, "接続が閉じられているか無効です（ActiveConnection が未設定）。", "ADODB.Command")
            if self.ctype != 1:
                raise strict_error("CommandType は 1（adCmdText）だけ使う決まりです")
            if len(args) > 1 and not (raw(args[1]) is EMPTY):
                raise strict_error("Execute の第 2 引数でパラメータを渡さない決まりです（Parameters.Append を使う）")
            return run_sql(self.conn, self.text, self.params.items, args[0] if args else None)
        if n == "commandtext":
            return self.text
        if n == "activeconnection":
            return self.conn or NOTHING
        return super().vb_get(name, args)

    def vb_set(self, name, args, value):
        n = name.lower()
        if n == "activeconnection":
            if not isinstance(value, AdoConnection):
                raise strict_error("ActiveConnection には Set で接続オブジェクトを入れる決まりです")
            self.conn = value
            return
        v = to_value(value)
        if n == "commandtext":
            self.text = to_str(v)
            return
        if n == "commandtype":
            self.ctype = vb_round(to_num(v))
            return
        if n in ("prepared", "commandtimeout", "namedparameters"):
            return
        super().vb_set(name, args, value)


# ======================================================================
# WScript.Shell / WScript.Network
# ======================================================================
def win_argv(cmd):
    """Windows の CommandLineToArgvW と同じ規則でコマンド行を分ける（\\" の扱いも同じ）。"""
    args, cur, i, n, inq, has = [], [], 0, len(cmd), False, False
    while i < n:
        c = cmd[i]
        if c == "\\":
            j = i
            while j < n and cmd[j] == "\\":
                j += 1
            nb = j - i
            if j < n and cmd[j] == '"':
                cur.append("\\" * (nb // 2))
                if nb % 2 == 1:
                    cur.append('"')
                    i = j + 1
                else:
                    i = j
                has = True
                continue
            cur.append("\\" * nb)
            i = j
            has = True
            continue
        if c == '"':
            if inq and i + 1 < n and cmd[i + 1] == '"':
                cur.append('"')
                i += 2
                continue
            inq = not inq
            has = True
            i += 1
            continue
        if c in " \t" and not inq:
            if has or cur:
                args.append("".join(cur))
                cur, has = [], False
            i += 1
            continue
        cur.append(c)
        has = True
        i += 1
    if has or cur:
        args.append("".join(cur))
    return args


class ShellExec(VBObject):
    def __init__(self, proc, out_lines, err_lines):
        self.proc = proc
        self.out_lines, self.err_lines = out_lines, err_lines

    def type_name(self):
        return "WshScriptExec"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "status":
            return 0 if self.proc.poll() is None else 1
        if n == "exitcode":
            rc = self.proc.poll()
            return rc if rc is not None else 0
        if n == "processid":
            return self.proc.pid
        if n == "terminate":
            if self.proc.poll() is None:
                self.proc.kill()
            return EMPTY
        if n in ("stdout", "stderr"):
            return PipeText(self.proc, self.out_lines if n == "stdout" else self.err_lines)
        return super().vb_get(name, args)


class PipeText(VBObject):
    def __init__(self, proc, lines):
        self.proc, self.lines = proc, lines

    def vb_get(self, name, args):
        n = name.lower()
        if n == "readall":
            self.proc.wait()
            time.sleep(0.05)
            s = "".join(self.lines)
            self.lines.clear()
            return s
        if n == "atendofstream":
            return self.proc.poll() is not None and not self.lines
        return super().vb_get(name, args)


class EnvColl(VBObject):
    def __init__(self, env):
        self.env = env

    def vb_default_get(self, args):
        argn(args, 1, "Environment")
        return self.env.get(sval(args[0], "Environment").upper(), "")


class Shell(VBObject):
    def __init__(self, host):
        self.h = host

    def type_name(self):
        return "IWshShell3"

    def expand(self, s):
        env = self.h.sim.env
        return re.sub(r"%([^%]+)%", lambda m: env.get(m.group(1).upper(), m.group(0)), s)

    def vb_get(self, name, args):
        n = name.lower()
        sim = self.h.sim
        if n == "environment":
            return EnvColl(sim.env)
        if n == "expandenvironmentstrings":
            return self.expand(sval(args[0], "ExpandEnvironmentStrings"))
        if n == "currentdirectory":
            return "C:\\Windows\\System32\\inetsrv"
        if n == "exec":
            cmd = self.expand(sval(args[0], "Exec"))
            argv = win_argv(cmd)
            return sim.spawn(argv, cmd)
        if n == "run":
            cmd = self.expand(sval(args[0], "Run"))
            wait = to_bool(val(args[2])) if len(args) > 2 else False
            return sim.run_cmd(win_argv(cmd), cmd, wait)
        return super().vb_get(name, args)


class Network(VBObject):
    def __init__(self, host):
        self.h = host

    def vb_get(self, name, args):
        n = name.lower()
        if n == "username":
            return self.h.sim.impersonated_user
        if n == "computername":
            return self.h.sim.server_name
        if n == "userdomain":
            return self.h.sim.server_name
        return super().vb_get(name, args)


# ======================================================================
# MSXML2.ServerXMLHTTP（設置チェックが自分のサーバーの画面を叩くため）
# ======================================================================
class XmlHttp(VBObject):
    def __init__(self, host):
        self.h = host
        self.method = None
        self.url = None
        self.res = None
        self.headers = {}

    def type_name(self):
        return "IServerXMLHTTPRequest2"

    def vb_get(self, name, args):
        n = name.lower()
        if n == "open":
            self.method = sval(args[0], "open").upper()
            self.url = sval(args[1], "open")
            if len(args) > 2 and to_bool(val(args[2])):
                raise strict_error("ServerXMLHTTP は非同期（第 3 引数 True）で使わない決まりです")
            return EMPTY
        if n in ("settimeouts", "setproxy", "setoption"):
            return EMPTY
        if n == "setrequestheader":
            self.headers[sval(args[0], "setRequestHeader")] = sval(args[1], "setRequestHeader")
            return EMPTY
        if n == "send":
            body = sval(args[0], "send") if args and val(args[0]) is not EMPTY else ""
            self.res = self.h.sim.http(self.method, self.url, body, self.headers)
            return EMPTY
        if n == "readystate":
            return 4 if self.res else 1
        if self.res is None:
            raise VBError(-1072896750, "send が呼ばれていません", "msxml6.dll")
        if n == "status":
            return self.res.code
        if n == "statustext":
            return self.res.status.split(" ", 1)[1] if " " in self.res.status else ""
        if n == "responsetext":
            return self.res.body.decode("utf-8", errors="replace")
        if n == "responsebody":
            return VBinary(self.res.body)
        if n == "getresponseheader":
            k = sval(args[0], "getResponseHeader").lower()
            for hk, hv in self.res.headers.items():
                if hk.lower() == k:
                    return hv
            return ""
        if n == "getallresponseheaders":
            return "\r\n".join(f"{k}: {v}" for k, v in self.res.headers.items())
        return super().vb_get(name, args)


# ======================================================================
# 1 回の要求ぶんの文脈
# ======================================================================
class AspHost:
    def __init__(self, sim, req):
        self.sim = sim
        self.req = req
        self.status = "200 OK"
        self.headers = {}
        self.props = {"buffer": True}
        self.buffer = True
        self.flushed = False
        self.parts = []          # ("t", 文字列) / ("b", バイト)
        self.cookies = {}
        self.codepage = None
        self.directive_codepage = None
        self.last_error = None
        self._fso = None
        self.objs = {
            "response": ResponseObj(self),
            "request": RequestObj(self),
            "server": ServerObj(self),
            "session": Forbidden("Session", "セッションが無効な環境で落ちるため（7-4）"),
            "application": Forbidden("Application", "状態をサーバーに持たない決まり"),
        }

    @property
    def body_text(self):
        return self.parts

    @body_text.setter
    def body_text(self, v):
        self.parts = []

    @property
    def body_bin(self):
        return self.parts

    @body_bin.setter
    def body_bin(self, v):
        self.parts = []

    def fso(self):
        if self._fso is None:
            self._fso = FSO(self)
        return self._fso

    def intrinsic(self, low):
        return self.objs.get(low)

    def check_header(self, what):
        if self.flushed or (not self.buffer and self.parts):
            raise VBError(-2147467259, f"{what}: 出力した後でヘッダーは変えられません（ASP 0156）", "Response オブジェクト")

    def cookie_out(self, name):
        if name not in self.cookies:
            self.cookies[name] = CookieOut(name)
        return self.cookies[name]

    def out_text(self, s):
        if s:
            self.parts.append(("t", s))

    def out_bin(self, b):
        self.parts.append(("b", b))

    def write_block(self, s):
        self.parts.append(("t", s))

    def create_object(self, progid):
        p = progid.lower()
        if p == "adodb.connection":
            return AdoConnection(self)
        if p == "adodb.command":
            return AdoCommand(self)
        if p == "adodb.stream":
            return AdoStream(self)
        if p == "adodb.recordset":
            raise strict_error("ADODB.Recordset を直接作らない決まりです（Command.Execute の結果を使う）")
        if p == "scripting.dictionary":
            return Dictionary()
        if p == "scripting.filesystemobject":
            return self.fso()
        if p == "wscript.shell":
            return Shell(self)
        if p == "wscript.network":
            return Network(self)
        if p in ("msxml2.serverxmlhttp.6.0", "msxml2.serverxmlhttp"):
            return XmlHttp(self)
        raise VBError(429, f"ActiveX コンポーネントはオブジェクトを作成できません: '{progid}'")

    def body_bytes(self):
        out = bytearray()
        for kind, x in self.parts:
            if kind == "t":
                cp = self.codepage or self.directive_codepage
                if cp == 65001:
                    out += x.encode("utf-8")
                else:
                    try:
                        out += x.encode("cp932")
                    except UnicodeEncodeError:
                        out += x.encode("cp932", errors="replace")
            else:
                out += x
        return bytes(out)
