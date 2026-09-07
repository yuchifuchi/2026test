# -*- coding: utf-8 -*-
"""マニュアルに載せる画面写真を、モックアップから撮り直す。

画面を直したら、これを実行してから tools/embed_manual_images.py を実行する。
撮った PNG は docs/manual/img/ に入り、マニュアル HTML にも埋め込まれる。

    python3 tools/shoot_manual_images.py

Playwright と Chromium が要る (この環境では /opt/pw-browsers/chromium)。
"""
import os
import sys

from playwright.sync_api import sync_playwright

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = "file://" + os.path.join(HERE, "mockup", "nippou-mockup.html")
OUT = os.path.join(HERE, "docs", "manual", "img")
CHROME = os.environ.get("CHROMIUM", "/opt/pw-browsers/chromium")

# (ファイル名, 画面, 縦長か, 役割, 前にやること)
SHOTS = [
    ("01_メニュー.png",           "menu",   True,  "staff", None),
    ("02_受付入力.png",           "entry",  False, "staff", None),
    ("03_その他業務.png",         "tasks",  False, "staff", None),
    ("04_日報.png",               "daily",  True,  "staff", None),
    ("05_帳票プレビュー.png",     "paper",  True,  "staff", None),
    ("06_集計表.png",             "sum",    False, "staff", None),
    ("07_入力もれ.png",           "check",  False, "staff", None),
    ("08_マスタ保守_担当者.png",  "master", False, "staff", '#ms-tabs button[data-mt="op"]'),
    ("09_マスタ保守_区分.png",    "master", False, "staff", '#ms-tabs button[data-mt="kb"]'),
    ("10_メニュー_パート職員.png", "menu",  True,  "part",  None),
]


def main():
    os.makedirs(OUT, exist_ok=True)
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=CHROME)
        pg = b.new_page(viewport={"width": 870, "height": 675}, device_scale_factor=2)
        pg.goto(SRC)
        pg.wait_for_timeout(700)
        for name, screen, tall, role, pre in SHOTS:
            pg.click('#rolesw button[data-role="%s"]' % role)
            pg.click('nav button[data-s="%s"]' % screen)
            if pre:
                pg.click(pre)
            pg.wait_for_timeout(350)
            pg.screenshot(path=os.path.join(OUT, name), full_page=tall)
            print("撮影:", name)
        b.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
