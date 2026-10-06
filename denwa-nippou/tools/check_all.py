"""すべての検査を順に流し、通ったら納品物の zip を作る。1 つでも失敗したら、そこで止める。

  python3 tools/check_all.py
"""
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
STEPS = [
    ("検査用の部品", ["fetch_deps.py"]),
    (".accdb を作って検査・組み立て", ["build.py"]),
    ("ASP・CSS の静的検査", ["lint_all.py"]),
    ("検査の道具の自己検査", ["tests/test_checkers.py"]),
    ("SQL を 1 本ずつ実行して期待値と突き合わせ", ["tests/test_sql.py"]),
    ("画面を通した通しの検査", ["tests/test_flow.py"]),
    ("困ったときの道具の検査", ["tests/test_troubleshoot.py"]),
    ("帳票（罫線・はみ出し・A4 1 枚）", ["render_sheet.py"]),
    ("手順書の検査", ["tests/test_docs.py"]),
    ("画面の見本（モックアップ）を作る", ["mockup.py"]),
    ("納品前の検査と zip", ["package.py"]),
]


def main():
    t0 = time.time()
    for i, (name, cmd) in enumerate(STEPS, 1):
        t = time.time()
        r = subprocess.run([sys.executable, os.path.join(HERE, cmd[0])] + cmd[1:], capture_output=True, text=True)
        out = "\n".join(l for l in (r.stdout + r.stderr).splitlines() if "JAVA_TOOL_OPTIONS" not in l)
        last = [l for l in out.splitlines() if l.startswith("OK") or l.startswith("NG")]
        if r.returncode != 0:
            print(f"[{i}/{len(STEPS)}] NG  {name}")
            print(out[-6000:])
            sys.exit(1)
        print(f"[{i}/{len(STEPS)}] OK  {name}（{time.time() - t:.0f} 秒） {last[-1] if last else ''}")
    print(f"すべて通りました（{time.time() - t0:.0f} 秒）")


if __name__ == "__main__":
    main()
