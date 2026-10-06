"""Windows のパス（C:\\... / \\\\サーバー\\共有\\...）を検査用の作業フォルダに対応づける。
Windows のファイル名は大文字・小文字を区別しないので、それも再現する。"""
import os
import re


class WinFS:
    def __init__(self, base):
        self.base = os.path.abspath(base)

    def _split(self, win):
        p = win.replace("/", "\\")
        m = re.match(r"^([A-Za-z]):\\?(.*)$", p)
        if m:
            return os.path.join(self.base, "drive_" + m.group(1).lower()), m.group(2)
        m = re.match(r"^\\\\([^\\]+)\\([^\\]+)\\?(.*)$", p)
        if m:
            return os.path.join(self.base, "unc", m.group(1).lower(), m.group(2).lower()), m.group(3)
        return None, None

    def is_abs(self, win):
        return self._split(win)[0] is not None

    def to_local(self, win):
        root, rest = self._split(win)
        if root is None:
            raise ValueError(f"Windows の絶対パスではありません: {win}")
        cur = root
        for part in [x for x in rest.split("\\") if x not in ("", ".")]:
            if part == "..":
                cur = os.path.dirname(cur)
                continue
            # 大文字・小文字を区別せずに既存のものを探す
            hit = None
            if os.path.isdir(cur):
                low = part.lower()
                for name in os.listdir(cur):
                    if name.lower() == low:
                        hit = name
                        break
            cur = os.path.join(cur, hit or part)
        return cur

    def to_win(self, local):
        local = os.path.abspath(local)
        rel = os.path.relpath(local, self.base)
        parts = rel.split(os.sep)
        if parts[0].startswith("drive_"):
            return parts[0][6:].upper() + ":\\" + "\\".join(parts[1:])
        if parts[0] == "unc":
            return "\\\\" + "\\".join(parts[1:])
        raise ValueError(local)
