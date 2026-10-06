"""Java の SqlBridge（UCanAccess）を呼ぶ。"""
import json
import os
import subprocess
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))          # denwa-nippou
JAVA_DIR = os.path.join(ROOT, "tools", "java")
LIB = os.path.join(ROOT, "tools", "lib")
CLASSES = os.path.join(ROOT, "build", "classes")


class BridgeError(Exception):
    pass


def compile_java(name, libset):
    """必要なら javac で作り直す（毎回ソースから起動すると遅いため）。"""
    out = os.path.join(CLASSES, libset)
    src = os.path.join(JAVA_DIR, name + ".java")
    cls = os.path.join(out, name + ".class")
    if not os.path.exists(cls) or os.path.getmtime(cls) < os.path.getmtime(src):
        os.makedirs(out, exist_ok=True)
        cp = os.path.join(LIB, libset, "*")
        r = subprocess.run(["javac", "-encoding", "UTF-8", "-nowarn", "-cp", cp, "-d", out, src],
                           capture_output=True, text=True)
        if r.returncode != 0:
            raise BridgeError("javac に失敗: " + r.stderr)
    return out


def java_cmd(name, libset):
    out = compile_java(name, libset)
    cp = out + os.pathsep + os.path.join(LIB, libset, "*")
    return ["java", "-Dfile.encoding=UTF-8", "-Dstdout.encoding=UTF-8", "-Dstderr.encoding=UTF-8", "-cp", cp, name]


class Bridge:
    def __init__(self):
        if not os.path.isdir(os.path.join(LIB, "ucanaccess")):
            raise BridgeError("tools/lib/ucanaccess がありません。python3 tools/fetch_deps.py を先に実行してください")
        self.p = subprocess.Popen(java_cmd("SqlBridge", "ucanaccess"), stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                  stderr=subprocess.PIPE, text=True, encoding="utf-8", bufsize=1)
        self.stderr = []
        threading.Thread(target=self._drain, daemon=True).start()
        self.path = None
        self.in_tx = False

    def _drain(self):
        for line in self.p.stderr:
            if "JAVA_TOOL_OPTIONS" not in line:
                self.stderr.append(line)

    def req(self, **kw):
        self.p.stdin.write(json.dumps(kw, ensure_ascii=False) + "\n")
        self.p.stdin.flush()
        while True:
            line = self.p.stdout.readline()
            if not line:
                raise BridgeError("SqlBridge が終了しました: " + "".join(self.stderr[-20:]))
            line = line.strip()
            if line.startswith("{"):
                break
        r = json.loads(line)
        if not r.get("ok"):
            raise BridgeError(r.get("error", "不明なエラー"))
        return r

    def open(self, path):
        path = os.path.abspath(path)
        if self.path == path:
            return
        if self.path:
            self.req(op="close")
        self.req(op="open", path=path)
        self.path = path

    def close(self):
        if self.path:
            self.req(op="close")
            self.path = None

    def query(self, sql, params):
        r = self.req(op="query", sql=sql, params=params)
        return r["columns"], r["rows"]

    def execute(self, sql, params):
        return self.req(op="exec", sql=sql, params=params)["affected"]

    def begin(self):
        self.req(op="begin")
        self.in_tx = True

    def commit(self):
        self.req(op="commit")
        self.in_tx = False

    def rollback(self):
        self.req(op="rollback")
        self.in_tx = False

    def shutdown(self):
        try:
            self.close()
            self.p.stdin.close()
            self.p.wait(timeout=10)
        except Exception:
            self.p.kill()
