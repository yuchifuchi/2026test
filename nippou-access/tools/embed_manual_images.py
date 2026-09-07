# -*- coding: utf-8 -*-
"""操作マニュアル.html の画面写真を docs/manual/img/ の PNG で入れ替える。

このマニュアルは、共有フォルダに置いてダブルクリックで開く 1 ファイルなので、
画像も中に埋め込んである (data URI)。画面を撮り直したら、これを実行して
埋め込み直す。どの img がどのファイルかは、HTML 側の data-img 属性で決まる。

撮った PNG は 2 倍の大きさ (幅 1740) なので、そのまま入れるとファイルが
重くなりすぎる。埋め込むときに幅 940 に縮め、色数を落としてから入れている。
紙に印刷しても読める大きさで、1 ファイルが 1 MB 弱に収まる。

    python3 tools/shoot_manual_images.py    # 撮り直す
    python3 tools/embed_manual_images.py    # 埋め込み直す
"""
import base64
import io
import os
import re
import sys

from PIL import Image

EMBED_WIDTH = 940      # 埋め込むときの幅 (画面写真は 2 倍の大きさで撮っている)
EMBED_COLORS = 256     # 画面写真は色数が少ないので、これで見た目は変わらない

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(HERE, "docs", "manual", "操作マニュアル.html")
IMG = os.path.join(HERE, "docs", "manual", "img")


def shrink(path):
    """幅 940 に縮めて色数を落とした PNG のバイト列を返す。"""
    im = Image.open(path).convert("RGB")
    if im.width > EMBED_WIDTH:
        h = int(round(im.height * EMBED_WIDTH / float(im.width)))
        im = im.resize((EMBED_WIDTH, h), Image.LANCZOS)
    im = im.quantize(colors=EMBED_COLORS, method=Image.MEDIANCUT)
    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def main():
    s = io.open(HTML, encoding="utf-8").read()
    used, missing = [], []

    def repl(m):
        tag = m.group(0)
        name = re.search(r'data-img="([^"]+)"', tag)
        if not name:
            return tag
        path = os.path.join(IMG, name.group(1))
        if not os.path.exists(path):
            missing.append(name.group(1))
            return tag
        b64 = base64.b64encode(shrink(path)).decode("ascii")
        used.append((name.group(1), len(b64)))
        return re.sub(r'src="[^"]*"', 'src="data:image/png;base64,' + b64 + '"', tag)

    s = re.sub(r"<img [^>]*>", repl, s)
    if missing:
        print("画像が見つかりません:", ", ".join(missing))
        return 1
    io.open(HTML, "w", encoding="utf-8").write(s)
    for n, k in used:
        print("埋め込み: %-26s %6.0f KB" % (n, k / 1024.0))
    print("合計 %d 枚 / %.1f MB" % (len(used), len(s.encode("utf-8")) / 1048576.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
