# -*- coding: utf-8 -*-
"""パート職員用と職員用の、2 つのフォルダを書き出す。

なぜ分けるか
    職員だけが使う画面を、置き場所ごと分けてしまう。
    職員用フォルダのアドレスにだけ IIS で IP 制限をかければ、
    Windows 認証が無くても「職員の席からしか開けない」状態にできる。

    URL だけ分けて中身を同じにする手もあるが、
    「このフォルダには職員用の画面が無い」ほうが確かめやすい。

置くもの
    nippou        … 受付入力・その他業務だけ (config.asp の ROLE = "part")
    nippou-staff  … 上に加えて 日報・帳票・集計・入力もれ・マスタ保守
                    (ROLE = "staff")

    共通のファイル (include / css / エラー画面 / 設置チェック) は両方に入れる。
    include を親フォルダから読むことは IIS が既定で禁じているため、
    それぞれのフォルダが自分のぶんを持つ形にしている。

    python3 tools/make_two_folders.py [出力先フォルダ]
"""
import io
import os
import shutil
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(HERE, "web")

# 両方に入れるもの
COMMON = ["error.asp", "setup_check.asp", "web.config",
          "probe1.asp", "probe2.asp", "probe3.asp", "probe4.asp",
          "probe5.asp", "probe6.asp", "probe7.asp"]
COMMON_DIRS = ["include", "css"]

PART = ["default.asp", "entry.asp", "tasks.asp"]
STAFF = ["default.asp", "staff.asp", "entry.asp", "tasks.asp", "daily.asp",
         "report.asp", "printpdf.asp", "summary.asp", "check.asp", "master.asp"]


def copy(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)


def build(out, name, pages, role):
    base = os.path.join(out, name)
    shutil.rmtree(base, ignore_errors=True)
    for f in COMMON + pages:
        copy(os.path.join(WEB, f), os.path.join(base, f))
    for d in COMMON_DIRS:
        for root, _dirs, files in os.walk(os.path.join(WEB, d)):
            for f in files:
                p = os.path.join(root, f)
                copy(p, os.path.join(base, os.path.relpath(p, WEB)))

    # config.asp の ROLE を、このフォルダ用に書き換える
    cfg = os.path.join(base, "include", "config.asp")
    s = io.open(cfg, encoding="utf-8").read()
    old = 'Const ROLE = "part"'
    if old not in s:
        raise SystemExit("config.asp の Const ROLE が見つかりません")
    s = s.replace(old, 'Const ROLE = "%s"' % role, 1)
    io.open(cfg, "w", encoding="utf-8", newline="\n").write(s)

    n = sum(len(f) for _r, _d, f in os.walk(base))
    print("%-14s %2d ファイル  (ROLE = %s)" % (name, n, role))
    return base


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "dist_web")
    os.makedirs(out, exist_ok=True)
    build(out, "nippou", PART, "part")
    build(out, "nippou-staff", STAFF, "staff")
    print("出力先:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
