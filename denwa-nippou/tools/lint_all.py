"""組み立てた nippou の .asp と css を、すべて静的検査する。

  python3 tools/lint_all.py [build/dist/nippou]
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from vbsim.lint import lint_page  # noqa: E402
from vbsim import csscheck  # noqa: E402

ROOT = os.path.dirname(HERE)


def main():
    dist = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "build", "dist", "nippou")
    rel_base = os.path.dirname(os.path.dirname(dist))
    probs = []
    pages = 0
    for role in ("part", "staff"):
        d = os.path.join(dist, role)
        for name in sorted(os.listdir(d)):
            if name.endswith(".asp"):
                pages += 1
                for f in lint_page(os.path.join(d, name), rel_base, True):
                    probs.append(str(f))
        for name in sorted(os.listdir(os.path.join(d, "include"))):
            # include だけを単独で読んでも壊れていないこと（宣言だけであること）
            for f in lint_page(os.path.join(d, "include", name), rel_base, False):
                probs.append(str(f))
        for p in csscheck.check(open(os.path.join(d, "css", "style.css"), encoding="utf-8").read()):
            probs.append(f"{role}/css/style.css: {p}")
    probs = sorted(set(probs))
    for p in probs:
        print("NG:", p)
    if probs:
        print(f"NG: 静的検査で {len(probs)} 件見つかりました")
        sys.exit(1)
    print(f"OK: 静的検査（画面 {pages} 枚・include・css）")


if __name__ == "__main__":
    main()
