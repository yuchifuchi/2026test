"""IIS + ASP の模擬。URL から .asp を探して実行し、web.config の規則も（検査に必要な範囲で）再現する。"""
import os
import re
import shutil
import signal
import subprocess
import threading
import time
import urllib.parse
import xml.etree.ElementTree as ET

from . import aspfile
from .builtins import Builtins
from .host import AspHost, ShellExec
from .interp import Interp, ExitLoop, ExitProc
from .parser import parse, ParseError
from .values import VBError, ResponseEnd, strict_error
from .winfs import WinFS

# web.config に書くとサイト全体が 500.19 になる区画（IIS の既定でロックされているもの）
LOCKED_SECTIONS = [
    ("system.webServer", "asp"),
    ("system.webServer/security", "authentication"),
    ("system.webServer/security", "ipSecurity"),
    ("system.webServer", "handlers"),
    ("system.webServer", "modules"),
    ("system.webServer", "serverRuntime"),
    ("system.webServer", "httpLogging"),
    ("system.webServer", "isapiFilters"),
    ("system.webServer", "cgi"),
]

MIME = {".css": "text/css", ".html": "text/html", ".htm": "text/html", ".txt": "text/plain", ".png": "image/png",
        ".gif": "image/gif", ".jpg": "image/jpeg", ".js": "application/javascript", ".pdf": "application/pdf",
        ".ico": "image/x-icon", ".svg": "image/svg+xml"}

EDGE_WIN = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"


class Result:
    def __init__(self, status, headers=None, body=b"", url=""):
        self.status = status
        self.code = int(status.split()[0])
        self.headers = headers or {}
        self.body = body
        self.url = url
        self.error = None
        self.warnings = []
        self.strict = []
        self.handled_by = None

    @property
    def text(self):
        return self.body.decode("utf-8", errors="replace")

    def __repr__(self):
        return f"<Result {self.status} {self.url} {len(self.body)}B>"


class AspSim:
    def __init__(self, workdir, nippou_dir, bridge, server_name="NIPPOU-SV", edge_present=True):
        if os.path.exists(workdir):
            shutil.rmtree(workdir)
        os.makedirs(workdir)
        self.winfs = WinFS(workdir)
        self.webroot_win = "C:\\inetpub\\wwwroot"
        self.webroot = self.winfs.to_local(self.webroot_win)
        os.makedirs(self.webroot)
        shutil.copytree(nippou_dir, os.path.join(self.webroot, "nippou"))
        os.makedirs(self.winfs.to_local("C:\\Windows\\Temp"))
        if edge_present:
            ep = self.winfs.to_local(EDGE_WIN)
            os.makedirs(os.path.dirname(ep))
            open(ep, "wb").close()
        self.server_name = server_name
        self.providers = {"microsoft.ace.oledb.12.0", "microsoft.ace.oledb.16.0"}
        self.env = {"TEMP": "C:\\Windows\\Temp", "TMP": "C:\\Windows\\Temp", "PROCESSOR_ARCHITECTURE": "AMD64",
                    "USERNAME": "DefaultAppPool", "COMSPEC": "C:\\Windows\\system32\\cmd.exe",
                    "SYSTEMROOT": "C:\\Windows", "WINDIR": "C:\\Windows"}
        self.impersonated_user = "IUSR"
        self.readonly_dirs = []
        self.db_readonly = False
        self.bridge = bridge
        self.stats = {"select": 0, "write": 0}
        self.sql_hook = None
        self.chromium = "/opt/pw-browsers/chromium"
        self.ping_scale = 0.25
        self.depth = 0
        self.procs = {}
        self.log = []

    # ---------- パス ----------
    def map_path(self, path, current_url):
        p = path.replace("\\", "/")
        if ".." in p.split("/"):
            raise VBError(-2147467259, "Server.MapPath() に '..' は使えません（親パスが無効のため。ASP 0175）", "Server.MapPath()")
        if re.match(r"^[A-Za-z]:", path) or path.startswith("\\\\"):
            raise VBError(-2147467259, "Server.MapPath() には物理パスではなく仮想パスを渡します（ASP 0173）", "Server.MapPath()")
        if p.startswith("/"):
            v = p
        else:
            base = current_url.rsplit("/", 1)[0]
            v = base + "/" + p
        return self.webroot_win + v.replace("/", "\\").rstrip("\\") if v != "/" else self.webroot_win

    def vpath_of_display(self, disp):
        return "/" + disp

    # ---------- web.config ----------
    def configs_for(self, local_dir):
        """サイトの先頭からそのフォルダまでの web.config を順に読む。"""
        rel = os.path.relpath(local_dir, self.webroot)
        dirs = [self.webroot]
        if rel != ".":
            cur = self.webroot
            for part in rel.split(os.sep):
                cur = os.path.join(cur, part)
                dirs.append(cur)
        out = []
        for d in dirs:
            wc = os.path.join(d, "web.config")
            if os.path.isfile(wc):
                out.append((d, wc))
        return out

    def check_config(self, wc):
        try:
            raw = open(wc, "rb").read()
            root = ET.fromstring(raw)
        except ET.ParseError as e:
            return f"web.config が XML として読めません: {e}"
        for parent, name in LOCKED_SECTIONS:
            node = root
            for part in parent.split("/"):
                node = node.find(part) if node is not None else None
            if node is not None and node.find(name) is not None:
                return f"web.config に <{name}> があります（IIS の既定でロックされた区画。サイト全体が 500.19 になります）"
        return None

    def rules(self, local_dir):
        r = {"errors": {}, "existing": "Auto", "allow_unlisted": True, "allowed": set(), "denied": {".config", ".mdb", ".accdb", ".inc"}}
        for d, wc in self.configs_for(local_dir):
            root = ET.fromstring(open(wc, "rb").read())
            sws = root.find("system.webServer")
            if sws is None:
                continue
            he = sws.find("httpErrors")
            if he is not None:
                r["existing"] = he.get("existingResponse", r["existing"])
                for el in he:
                    if el.tag == "remove":
                        r["errors"].pop((el.get("statusCode"), el.get("subStatusCode", "-1")), None)
                    elif el.tag == "error":
                        r["errors"][(el.get("statusCode"), el.get("subStatusCode", "-1"))] = (el.get("path"), el.get("responseMode", "File"))
            fe = sws.find("security/requestFiltering/fileExtensions")
            if fe is not None:
                if fe.get("allowUnlisted") is not None:
                    r["allow_unlisted"] = fe.get("allowUnlisted").lower() == "true"
                for el in fe:
                    if el.tag == "add":
                        ext = el.get("fileExtension", "").lower()
                        if el.get("allowed", "true").lower() == "true":
                            r["allowed"].add(ext)
                            r["denied"].discard(ext)
                        else:
                            r["denied"].add(ext)
        return r

    # ---------- 要求 ----------
    def request(self, method, url, form=None, cookies=None, sv=None, headers=None, _err=None):
        u = urllib.parse.urlsplit(url)
        path = urllib.parse.unquote(u.path)
        query = urllib.parse.parse_qsl(u.query, keep_blank_values=True)
        form = list(form.items()) if isinstance(form, dict) else (form or [])
        req = {"method": method.upper(), "path": path, "query": query, "form": form, "cookies": dict(cookies or {}),
               "url": url}
        # ローカルのファイルを探す
        local = self.winfs.to_local(self.webroot_win + path.replace("/", "\\"))
        local_dir = local if os.path.isdir(local) else os.path.dirname(local)
        for d, wc in self.configs_for(local_dir):
            msg = self.check_config(wc)
            if msg:
                res = Result("500 Internal Server Error", {"Content-Type": "text/html"},
                             ("HTTP Error 500.19 - Internal Server Error " + msg).encode("utf-8"), url)
                res.error = msg
                return res
        if os.path.isdir(local):
            if not path.endswith("/"):
                return Result("301 Moved Permanently", {"Location": path + "/"}, b"", url)
            dflt = os.path.join(local, "default.asp")
            hit = [x for x in os.listdir(local) if x.lower() == "default.asp"]
            if not hit:
                return Result("403 Forbidden", {}, b"", url)
            local = os.path.join(local, hit[0])
            path = path + hit[0]
            req["path"] = path
        if not os.path.isfile(local):
            return Result("404 Not Found", {}, b"", url)
        rules = self.rules(os.path.dirname(local))
        ext = os.path.splitext(local)[1].lower()
        if ext in rules["denied"] or (not rules["allow_unlisted"] and ext not in rules["allowed"]):
            return Result("404 Not Found", {"X-Sim": "404.7 requestFiltering"}, b"", url)
        if ext != ".asp":
            if ext not in MIME:
                return Result("404 Not Found", {"X-Sim": "404.3 MIME"}, b"", url)
            return Result("200 OK", {"Content-Type": MIME[ext]}, open(local, "rb").read(), url)
        req["server_vars"] = self.server_vars(req, sv, headers)
        return self.run_asp(local, req, rules, _err)

    def server_vars(self, req, sv, headers):
        d = {
            "REQUEST_METHOD": req["method"], "URL": req["path"], "SCRIPT_NAME": req["path"], "PATH_INFO": req["path"],
            "QUERY_STRING": urllib.parse.urlencode(req["query"]), "REMOTE_ADDR": "192.168.10.21",
            "LOCAL_ADDR": "192.168.10.5", "SERVER_NAME": self.server_name, "SERVER_PORT": "80",
            "HTTP_HOST": self.server_name, "HTTPS": "off", "LOGON_USER": "", "AUTH_USER": "",
            "APP_POOL_ID": "DefaultAppPool", "SERVER_SOFTWARE": "Microsoft-IIS/10.0", "INSTANCE_ID": "1",
            "APPL_PHYSICAL_PATH": self.webroot_win + "\\", "PATH_TRANSLATED": self.map_path(req["path"], req["path"]),
            "HTTP_USER_AGENT": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Edg/141.0",
        }
        if headers:
            for k, v in headers.items():
                d["HTTP_" + k.upper().replace("-", "_")] = v
        if sv:
            d.update(sv)
        return d

    def run_asp(self, local, req, rules, last_error=None):
        host = AspHost(self, req)
        host.last_error = last_error
        res = Result("200 OK", url=req["url"])
        interp = None
        err = None
        try:
            prog = aspfile.load_page(local, rel_base=self.webroot)
            cp = prog.directive.get("CODEPAGE")
            host.directive_codepage = int(cp) if cp and cp.isdigit() else None
            try:
                ast = parse(prog.code)
            except ParseError as e:
                ve = VBError(1002, e.msg, "Microsoft VBScript コンパイル エラー")
                ve.where = prog.linemap[e.line] if 0 <= e.line < len(prog.linemap) else ("?", 0)
                raise ve
            interp = Interp(prog, ast, host, Builtins(host))
            interp.run()
        except aspfile.AspSourceError as e:
            err = VBError(-2147467259, e.msg, "Active Server Pages エラー")
            err.where = (e.file, e.line)
        except ResponseEnd:
            pass
        except VBError as e:
            err = e
        except (ExitLoop, ExitProc):
            err = VBError(5, "Exit がループや Function の外に出ました")
        if interp is not None:
            res.warnings = list(interp.warnings)
        if err is not None:
            res.error = err
            if err.strict:
                res.strict.append(str(err))
            # ASP の実行時エラー = 500.100。web.config の指定があればそのページを実行する。
            target = rules["errors"].get(("500", "100"))
            if target and target[1] == "ExecuteURL" and last_error is None and self.depth < 3:
                self.depth += 1
                try:
                    sub = self.request(req["method"], target[0], form=req["form"], cookies=req["cookies"],
                                       sv={"URL": req["path"], "QUERY_STRING": urllib.parse.urlencode(req["query"])}, _err=err)
                finally:
                    self.depth -= 1
                sub.error = err
                sub.handled_by = target[0]
                sub.strict = res.strict + sub.strict
                sub.warnings = res.warnings + sub.warnings
                return sub
            res.status = "500 Internal Server Error"
            res.code = 500
            res.body = b"<h1>An error occurred on the server when processing the URL.</h1>"
            res.headers = {"Content-Type": "text/html"}
            return res
        res.status = host.status
        res.code = int(host.status.split()[0])
        ctype = host.props.get("contenttype", "text/html")
        cs = host.props.get("charset")
        res.headers = dict(host.headers)
        res.headers["Content-Type"] = ctype + (f"; charset={cs}" if cs else "")
        for c in host.cookies.values():
            res.headers.setdefault("Set-Cookie", "")
            res.headers["Set-Cookie"] += f"{c.name}={urllib.parse.quote(c.value)}; path={c.path};"
        res.body = host.body_bytes()
        if res.code >= 400 and rules["existing"] != "PassThrough":
            # IIS が本文を捨てて既定の画面に差し替える（7-3）
            res.body = b"<h1>An error occurred on the server when processing the URL.</h1>"
            res.warnings.append("自前の本文が IIS に差し替えられました（existingResponse が PassThrough ではありません）")
        return res

    def http(self, method, url, body, headers):
        u = urllib.parse.urlsplit(url)
        host = (u.hostname or "").lower()
        if host not in (self.server_name.lower(), "localhost", "127.0.0.1"):
            raise VBError(-2147012889, f"サーバー名またはアドレスを解決できませんでした: {host}", "msxml6.dll")
        if self.depth >= 3:
            raise strict_error("ServerXMLHTTP の呼び出しが入れ子になりすぎています")
        form = urllib.parse.parse_qsl(body, keep_blank_values=True) if body else []
        self.depth += 1
        try:
            path = u.path + ("?" + u.query if u.query else "")
            res = self.request(method, path, form=form, cookies={})
            hops = 0
            # ServerXMLHTTP（WinHTTP）は 301 / 302 を自動でたどる
            while res.code in (301, 302) and hops < 5:
                loc = res.headers.get("Location", "")
                if not loc.startswith("/"):
                    loc = u.path.rsplit("/", 1)[0] + "/" + loc
                res = self.request("GET", loc, cookies={})
                hops += 1
            return res
        finally:
            self.depth -= 1

    # ---------- 外部プログラム ----------
    def spawn(self, argv, cmd):
        exe = argv[0]
        if exe.lower() == EDGE_WIN.lower() and os.path.isfile(self.winfs.to_local(EDGE_WIN)):
            args = []
            for a in argv[1:]:
                m = re.match(r"^(--print-to-pdf|--user-data-dir)=(.*)$", a)
                if m:
                    if not self.winfs.is_abs(m.group(2)):
                        raise strict_error(f"{m.group(1)} に絶対パスでない値: {m.group(2)}")
                    args.append(f"{m.group(1)}={self.winfs.to_local(m.group(2))}")
                    continue
                if a.lower().startswith("file:///"):
                    wp = urllib.parse.unquote(a[8:]).replace("/", "\\")
                    args.append("file://" + urllib.parse.quote(self.winfs.to_local(wp)))
                    continue
                if a.startswith("--") or a.startswith("-"):
                    args.append(a)
                    continue
                raise strict_error(f"Edge に渡す引数が想定外です: {a!r}（引用符の付け方を確認）")
            if "--no-sandbox" not in args:
                args.append("--no-sandbox")
            self.log.append(("edge", argv))
            p = subprocess.Popen([self.chromium] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                                 start_new_session=True)
            out_lines, err_lines = [], []
            threading.Thread(target=lambda: [out_lines.append(x) for x in p.stdout], daemon=True).start()
            threading.Thread(target=lambda: [err_lines.append(x) for x in p.stderr], daemon=True).start()
            self.procs[p.pid] = p
            return ShellExec(p, out_lines, err_lines)
        raise strict_error(f"想定外の外部プログラムを起動しようとしています: {cmd}")

    def run_cmd(self, argv, cmd, wait):
        a = [x for x in argv]
        if a and os.path.basename(a[0].replace("\\", "/")).lower() in ("cmd.exe", "cmd") and len(a) > 2 and a[1].lower() == "/c":
            rest = a[2:]
            if rest[0].lower() == "ping":
                n = 1
                if "-n" in rest:
                    n = int(rest[rest.index("-n") + 1])
                time.sleep(max(0, n - 1) * self.ping_scale)
                return 0
            if rest[0].lower() == "taskkill":
                if "/PID" in [x.upper() for x in rest]:
                    i = [x.upper() for x in rest].index("/PID")
                    pid = int(rest[i + 1])
                    p = self.procs.get(pid)
                    if p and p.poll() is None:
                        os.killpg(p.pid, signal.SIGKILL)
                return 0
        raise strict_error(f"想定外のコマンドを実行しようとしています: {cmd}")
