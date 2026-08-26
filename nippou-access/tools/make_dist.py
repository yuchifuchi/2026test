# -*- coding: utf-8 -*-
"""src/*.bas を Windows の VBE がそのまま読める形 (Shift_JIS / CRLF) で dist/ に出す。

VBE の「ファイルのインポート」は UTF-8 の .bas を読むと日本語が化けるため、
配布用は CP932 に変換しておく。貼り付け運用しかしない場合は src/ のままでよい。
"""
import glob
import io
import os

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    out_dir = os.path.join(HERE, "dist")
    os.makedirs(out_dir, exist_ok=True)
    bad = 0
    for path in sorted(glob.glob(os.path.join(HERE, "src", "*.bas"))):
        name = os.path.basename(path)
        text = io.open(path, encoding="utf-8", newline="").read()
        text = text.replace("\r\n", "\n").replace("\n", "\r\n")
        try:
            data = text.encode("cp932")
        except UnicodeEncodeError as e:
            # 髙 (U+9AD9) などは CP932 に有る。無い文字だけを報告する。
            ch = text[e.start:e.end]
            print("  ! %s: CP932 に無い文字 %r (位置 %d)" % (name, ch, e.start))
            bad += 1
            continue
        with open(os.path.join(out_dir, name), "wb") as f:
            f.write(data)
        print("  dist/%s  (%d bytes)" % (name, len(data)))
    if bad:
        print("変換できなかったファイルがあります。該当文字を置き換えてください。")
    return bad


if __name__ == "__main__":
    raise SystemExit(main())
