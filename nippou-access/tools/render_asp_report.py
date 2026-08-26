# -*- coding: utf-8 -*-
"""web/report.asp のマークアップを静的 HTML に起こして、印刷結果を検証できるようにする。

IIS が無い環境でも、帳票の罫線と A4 1 枚に収まるかを確かめるため。
report.asp が使っている構文 (<%= %> と For…Next) だけを解釈する簡易レンダラ。
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 検証用のサンプル値。実データ (8/25) に合わせてある。
VALUES = {
    "H(Wareki(dt))": "令和　8年　8月　25日（火）",
    'H(hd("出勤者数"))': "6",
    'H(hd("回線数"))': "5",
    'H(hd("合計"))': "119",
    'H(hd("申込"))': "41", 'H(hd("申込_職員"))': "1",
    'H(hd("抽選"))': "14", 'H(hd("抽選_職員"))': "0",
    'H(hd("払込用紙"))': "0", 'H(hd("払込用紙_職員"))': "0",
    'H(hd("商品発送"))': "3", 'H(hd("商品発送_職員"))': "0",
    'H(hd("その他"))': "61", 'H(hd("その他_職員"))': "1",
    'H(hd("内交換"))': "0", 'H(hd("内返金"))': "0",
    'H(hd("特記事項"))': "昭和100年記念貨幣の抽選結果について、発表日の問合せが集中した。",
    'H(hd("職員代替案件"))': "支払方法の変更依頼 1 件を職員が対応。",
    'H(hd("要望"))': "抽選結果の発表日をハガキにも明記していただけると、問合せが減ると思われます。",
    "Ymd(dt)": "2026-08-25",
}
NAMES = ["髙嶋 名保美", "石原 裕子", "石井 淳子", "鈴木", "寺本 法子", "顧客G", ""]
TASKS_L = ["①　受注入力", "②　受注チェック", "③　戻り郵便処理／払込書チェック",
           "④　ＤＭ処理／戻り郵便", "⑤　架電", "⑥　その他（　　　　　　　　）",
           "⑦　エクセル入力", "⑧　エクセルチェック"]
TASKS_R = ["⑨　アンケート入力", "⑩　ﾊｶﾞｷﾃﾞｰﾀ入力", "⑪　ハガキデータチェック",
           "⑫　新規ｺｰﾄﾞ取り", "⑬　顧客整理", "", "", ""]


def expand(body):
    """<% For i = 0 To 6 %> … <% Next %> を展開し、<%= %> をサンプル値に置き換える。"""
    def loop(m):
        n = int(m.group(1)) + 1
        inner = m.group(2)
        out = []
        for i in range(n):
            chunk = inner
            chunk = chunk.replace("<%= H(names(i)) %>", NAMES[i] if i < len(NAMES) else "")
            chunk = chunk.replace("<%= H(names(i + 7)) %>", "")
            chunk = chunk.replace('<%= IIfS(i <= 4, H(rightN(i)), "") %>',
                                  TASKS_R[i] if i < len(TASKS_R) else "")
            chunk = chunk.replace('<%= IIfS(i <= 4, H(rightV(i)), "") %>', "")
            chunk = chunk.replace('<%= IIfS(i <= 4 And Len("" & rightN(i)) > 0, "件", "") %>',
                                  "件" if i < 5 else "")
            chunk = chunk.replace("<%= H(leftN(i)) %>", TASKS_L[i] if i < len(TASKS_L) else "")
            chunk = chunk.replace("<%= H(leftV(i)) %>", "")
            chunk = chunk.replace('<%= IIfS(Len("" & leftN(i)) > 0, "件", "") %>',
                                  "件" if i < len(TASKS_L) and TASKS_L[i] else "")
            out.append(chunk)
        return "".join(out)

    # <% … For i = 0 To N %> のように、コードブロックの末尾が For で終わる形にも対応する
    body = re.sub(r"<%(?:[^%]|%(?!>))*?For i = 0 To (\d+)\s*%>(.*?)<%\s*Next\s*%>",
                  loop, body, flags=re.S)

    def val(m):
        expr = m.group(1).strip()
        return VALUES.get(expr, "")
    body = re.sub(r"<%=\s*(.*?)\s*%>", val, body, flags=re.S)
    body = re.sub(r"<%.*?%>", "", body, flags=re.S)          # 残りのコードブロック
    body = re.sub(r"<!--#include[^>]*-->", "", body)
    return body


def main():
    src = io.open(os.path.join(HERE, "web", "report.asp"), encoding="utf-8").read()
    css = io.open(os.path.join(HERE, "web", "css", "style.css"), encoding="utf-8").read()

    body = expand(src)
    # PageHead / PageFoot は ASP 側の関数なので、同等の外枠を自前で付ける
    html = ("<!doctype html><html lang='ja'><head><meta charset='utf-8'>"
            "<title>日報 印刷</title><style>" + css + "</style></head><body>"
            "<header class='appbar'><a class='brand' href='#'>電話応対日報 集計システム</a></header>"
            "<main>" + body + "</main>"
            "<footer>電話応対日報 集計システム</footer></body></html>")
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "web", "_report_preview.html")
    io.open(out, "w", encoding="utf-8").write(html)
    print("出力:", out)


if __name__ == "__main__":
    main()
