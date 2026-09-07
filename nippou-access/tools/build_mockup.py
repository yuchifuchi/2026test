# -*- coding: utf-8 -*-
"""head.html + body.html + data.json + app.js を 1 枚の nippou-mockup.html にまとめる。

配る相手は Excel しか触らないかたなので、開くだけで動く 1 ファイルにしている。
"""
import io
import os

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
M = os.path.join(HERE, "mockup")


def read(n):
    return io.open(os.path.join(M, n), encoding="utf-8").read()


def main():
    html = (read("head.html") + "\n" + read("body.html") +
            "\n<script>\nvar D = " + read("data.json").strip() + ";\n</script>\n" +
            "<script>\n" + read("app.js") + "\n</script>\n")
    out = os.path.join(M, "nippou-mockup.html")
    io.open(out, "w", encoding="utf-8").write(html)
    print("出力:", out, "(%.1f KB)" % (len(html.encode("utf-8")) / 1024.0))


if __name__ == "__main__":
    main()
