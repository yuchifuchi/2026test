# -*- coding: utf-8 -*-
"""納品 zip を展開して、渡す前に確かめるべきことを機械的に見る。

これまでに実際にやらかしたこと（すべてここで捕まえられる）
  ・Windows パスを Python の文字列に直書きして \\a がベル文字になった
  ・__pycache__ を納品物に混ぜた
  ・Shift_JIS で渡すべきファイルを UTF-8 のままにした

ファイルごとに「どの文字コードで渡すのが正しいか」は場所で決まる。
  .txt / .vbs                     Shift_JIS  … メモ帳・WSH がそのまま読む
  任意_過去データ取込に使うVBA/*.bas  Shift_JIS  … VBE のインポートで化けない
  VBA原本_UTF8/*.bas               UTF-8      … 開発用の原本
  *.asp                           UTF-8 BOMなし … CODEPAGE=65001 と対
"""
import glob
import os
import sys
import zipfile

CTRL = {0: "NUL", 7: r"BEL (\a)", 8: r"BS (\b)", 11: r"VT (\v)",
        12: r"FF (\f)", 27: "ESC"}
BINARY = (".png", ".jpg", ".zip", ".accdb")


def want_encoding(rel):
    """このファイルはどの文字コードで渡すのが正しいか。"""
    if "VBA原本_UTF8" in rel:
        return "utf-8"
    if rel.endswith((".txt", ".vbs")):
        return "cp932"
    if rel.endswith(".bas"):
        return "cp932"          # 配布用。VBE がそのまま読める必要がある
    if rel.endswith(".asp"):
        return "utf-8-nobom"
    if rel.endswith((".html", ".css", ".md", ".csv", ".json", ".py", ".config")):
        return "utf-8"
    return None


def check(root):
    ng = []

    def bad(msg):
        ng.append(msg)

    files = [p for p in glob.glob(os.path.join(root, "**", "*"), recursive=True)
             if os.path.isfile(p)]

    for p in files:
        rel = os.path.relpath(p, root)
        data = open(p, "rb").read()

        if "__pycache__" in rel or rel.endswith((".pyc", ".pyo")):
            bad("納品物に不要: %s" % rel)
            continue

        if not rel.endswith(BINARY):
            for code, name in CTRL.items():
                if bytes([code]) in data:
                    bad("%s に制御文字 %s が %d 個" % (rel, name, data.count(bytes([code]))))

        enc = want_encoding(rel)
        if enc == "utf-8-nobom":
            if data.startswith(b"\xef\xbb\xbf"):
                bad("%s に BOM がある (ASP は BOM なしの UTF-8)" % rel)
            else:
                try:
                    data.decode("utf-8")
                except UnicodeDecodeError:
                    bad("%s が UTF-8 で読めない" % rel)
        elif enc:
            try:
                data.decode(enc)
            except UnicodeDecodeError:
                bad("%s が %s で読めない" % (rel, enc))

    # 入口となるファイルが揃っているか
    for must in ("00_はじめにお読みください.txt", "はじめに.html",
                 "はじめに読む_やさしい導入手順書.html",
                 "01_マニュアル/操作マニュアル.html",
                 "02_データベース_Access/データベースを作る.vbs",
                 "03_Webサイト_ASP/wwwroot/default.asp",
                 "04_モックアップ/モックアップ.html"):
        if not os.path.exists(os.path.join(root, must.replace("/", os.sep))):
            bad("見当たらない: %s" % must)

    return ng, len(files)


def check_zip(zip_path):
    ng = []
    z = zipfile.ZipFile(zip_path)
    for i in z.infolist():
        if any(ord(c) > 127 for c in i.filename) and not (i.flag_bits & 0x800):
            ng.append("zip 内の日本語名に UTF-8 フラグが無い: %s" % i.filename)
    return ng


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("使い方: python3 tools/check_package.py <展開したフォルダ> [zipのパス]")
    root = sys.argv[1]
    ng, n = check(root)
    if len(sys.argv) > 2:
        ng += check_zip(sys.argv[2])

    print("検査したファイル: %d" % n)
    if ng:
        print("問題 %d 件" % len(ng))
        for x in ng:
            print("  !", x)
        sys.exit(1)
    print("問題なし。納品できます。")
