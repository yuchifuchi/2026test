"""納品物を組み立てる（開発者の作業環境用）。

  python3 tools/build.py accdb      … db/schema.sql から .accdb を作り、形式と中身を検査する
  python3 tools/build.py assemble   … src/ から build/dist/nippou（part / staff / data）を組み立てる
  python3 tools/build.py            … 上の 2 つ

part と staff には、include と css を「同じもの」を置く（IIS の既定では #include の .. が使えないため）。
違うのは config.asp の ROLE と、web.config のエラー画面の場所だけ。
"""
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")
BUILD = os.path.join(ROOT, "build")
DIST = os.path.join(BUILD, "dist", "nippou")
DB_NAME = "日報集計_be.accdb"

sys.path.insert(0, HERE)
from vbsim.sqlbridge import java_cmd  # noqa: E402

COMMON_PAGES = ["default.asp", "entry.asp", "tasks.asp", "error.asp", "setup_check.asp"] + \
               [f"probe{i}.asp" for i in range(1, 8)]
STAFF_PAGES = ["staff.asp", "daily.asp", "report.asp", "printpdf.asp", "summary.asp", "check.asp", "master.asp"]
INCLUDES = ["auth.asp", "config.asp", "db.asp", "layout.asp", "sheet.asp", "sql.asp", "pdf.asp"]


def build_accdb(out_dir=None):
    out_dir = out_dir or os.path.join(BUILD, "db")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, DB_NAME)
    dump = os.path.join(out_dir, "accdb_dump.json")
    r = subprocess.run(java_cmd("BuildAccdb", "builder") + [os.path.join(ROOT, "db", "schema.sql"), out, dump],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stdout)
        print(r.stderr)
        raise SystemExit("NG: .accdb を作れませんでした")
    print(r.stdout.strip())
    return out, dump


def assemble(accdb):
    if os.path.exists(DIST):
        shutil.rmtree(DIST)
    for role in ("part", "staff"):
        d = os.path.join(DIST, role)
        os.makedirs(os.path.join(d, "include"))
        os.makedirs(os.path.join(d, "css"))
        pages = COMMON_PAGES + (STAFF_PAGES if role == "staff" else [])
        for p in pages:
            sub = "common" if p in COMMON_PAGES else "staff"
            shutil.copyfile(os.path.join(SRC, "pages", sub, p), os.path.join(d, p))
        for inc in INCLUDES:
            src = open(os.path.join(SRC, "include", inc), "rb").read()
            if inc == "config.asp":
                src = src.replace(b'Const ROLE = "{{ROLE}}"', f'Const ROLE = "{role}"'.encode())
            open(os.path.join(d, "include", inc), "wb").write(src)
        shutil.copyfile(os.path.join(SRC, "css", "style.css"), os.path.join(d, "css", "style.css"))
        wc = open(os.path.join(SRC, "webconfig", "app.web.config"), "rb").read().replace(b"{{FOLDER}}", role.encode())
        open(os.path.join(d, "web.config"), "wb").write(wc)
    data = os.path.join(DIST, "data")
    os.makedirs(data)
    shutil.copyfile(accdb, os.path.join(data, DB_NAME))
    shutil.copyfile(os.path.join(SRC, "webconfig", "data.web.config"), os.path.join(data, "web.config"))
    # 取り違え防止の確認: part に職員用の画面が 1 つも無いこと
    for p in STAFF_PAGES:
        if os.path.exists(os.path.join(DIST, "part", p)):
            raise SystemExit(f"NG: part に職員用の画面 {p} があります")
    print(f"組み立て: {DIST}")
    return DIST


def main():
    what = sys.argv[1] if len(sys.argv) > 1 else "all"
    if what in ("accdb", "all"):
        accdb, dump = build_accdb()
        r = subprocess.run([sys.executable, os.path.join(HERE, "check_accdb.py"), accdb, dump])
        if r.returncode != 0:
            raise SystemExit("NG: .accdb の検査に通りませんでした")
    if what in ("assemble", "all"):
        assemble(os.path.join(BUILD, "db", DB_NAME))


if __name__ == "__main__":
    main()
