"""手順書の検査。

・手順書で「　」に入れた言葉（画面に出ている文字）が、本当に画面のどこかに出ているか。
  画面を直したのに手順書が古いまま、という食い違いを見つける。
  Windows・Excel・Edge の画面の言葉は、下の WINDOWS_UI に挙げたものだけを認める。
・手順書と画面に、利用者に通じないカタカナの技術用語が入っていないか。
・手順書の各段に「できたことの確かめ方」があるか。
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from harness import ROOT, Fail  # noqa: E402

DOCS = os.path.join(ROOT, "docs")
SRC = os.path.join(ROOT, "src")

# 手順書と画面で使わない言葉（手順 0 章の決まり）
BANNED = ["デプロイ", "バインド", "インスタンス", "リポジトリ", "コミット", "ビルド", "ディレクトリ", "セッション",
          "キャッシュ", "モジュール", "ランタイム", "トランザクション", "エンコード", "デコード", "スクリプト"]

# このシステムの外（Windows・Excel・Edge・Access）の画面に出る言葉
WINDOWS_UI = {
    "すべて展開...", "展開", "コピー", "貼り付け", "プロパティ", "セキュリティ", "編集", "追加", "場所", "OK",
    "選択するオブジェクト名を入力してください", "名前の確認", "変更", "許可", "システム", "デバイス名", "コンピューター名",
    "完了", "プログラムから開く", "メモ帳", "エクスプローラー", "サーバー マネージャー", "管理", "役割と機能の追加", "次へ",
    "サーバーの役割", "Web サーバー (IIS)", "Web サーバー", "アプリケーション開発", "ASP", "機能の追加", "インストール",
    "アプリケーション プール", "DefaultAppPool", "詳細設定", "32 ビット アプリケーションの有効化", "True",
    "IP アドレスおよびドメインの制限", "再表示", "保存しない", "ファイル", "情報", "データベースの最適化/修復",
    "名前を付けて保存", "文字コード", "UTF-8", "UTF-8 (BOM 付き)", "ANSI", "互換表示", "ヘッダーとフッター",
    "Microsoft Access データベース エンジン 2016 再頒布可能コンポーネント", "Windows ファイアウォール",
    "World Wide Web サービス（HTTP 受信トラフィック）", "インストール",
}
# IIS・ブラウザーが出す画面の文字（このシステムの外）
IIS_UI = {"HTTP エラー 500.19", "構成データが無効です", "404", "404.3", "ダウンロードしますか",
          "An error occurred on the server when processing the URL", "An error occurred on the server…"}
# 手順書の中の見出しの言葉・例に出した文字・フォルダの名前
DOC_WORDS = {"できたことの確かめ方", "うまくいかないとき", "縺", "電話応対日報_納品一式", "参考", "手順書", "nippou",
             "日報の控え", "日報集計_be.laccdb"}
# 手順書の中で「　」を、画面の文字ではなく言葉の引用に使っている所
QUOTES_OK = {"サーバー名", "変更", "未定義のエラー", "Access 2016 専用", "データベースの最適化/修復",
             "Microsoft Access データベース エンジン（64 ビット版）を入れてください", "ふだん使う画面を取り違えない",
             "UTF-8", "名前", "日報の列", "データベースの最適化（Access の『データベースの最適化/修復』）"}


def screen_text():
    out = []
    for dp, _, fns in os.walk(SRC):
        for fn in fns:
            if fn.endswith(".asp"):
                out.append(open(os.path.join(dp, fn), encoding="utf-8").read())
    # VBScript の文字列の "" は画面では " 1 つになる
    raw = "\n".join(out).replace('""', '"')
    # タグや VBScript の文字列のつなぎ目（" & 値 & "）を取り除いた形でも探せるようにする
    plain = re.sub(r"<[^>]+>", "", raw)
    plain = re.sub(r'"\s*&[^&\n]*?&\s*"', "", plain)
    return raw + "\n" + plain


def on_screen(q, screens):
    q = q.strip()
    if q.startswith("→"):
        q = q[1:]
    parts = [x for x in re.split(r"…|\s*\d+$", q) if x.strip()]
    return all(x.strip() in screens or re.sub(r"\s+", "", x) in re.sub(r"\s+", "", screens) for x in parts)


def strip_vb_comments(text):
    keep = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith("'"):
            continue
        keep.append(line)
    return "\n".join(keep)


def run():
    probs = []
    checks = 0
    screens = screen_text()
    doc_names = {os.path.splitext(fn)[0] for _, _, fns in os.walk(DOCS) for fn in fns}
    for dp, _, fns in os.walk(DOCS):
        for fn in fns:
            p = os.path.join(dp, fn)
            if not (fn.endswith(".md") or fn.endswith(".txt")):
                continue
            text = open(p, encoding="utf-8").read()
            rel = os.path.relpath(p, DOCS)
            if "手順書" in rel or fn.startswith("0_"):
                for w in BANNED:
                    if w in text:
                        probs.append(f"{rel}: 使わない言葉「{w}」があります")
                for q in re.findall(r"「([^「」\n]+)」", text):
                    checks += 1
                    if q in WINDOWS_UI or q in QUOTES_OK or q in IIS_UI or q in DOC_WORDS or q in doc_names:
                        continue
                    if re.search(r"\.(txt|md|xlsm|accdb|laccdb|pdf|asp|config)$", q) or q.startswith("include\\") or "○" in q:
                        continue
                    if q.endswith(("ほしい", "したい", "ください")):
                        continue   # 管理者への頼み方の例
                    if not on_screen(q, screens):
                        probs.append(f"{rel}: 「{q}」が画面のどこにもありません（画面の文字と手順書が食い違っていないか）")
                if "手順書" in rel:
                    for sec in re.split(r"\n## ", text)[1:]:
                        title = sec.splitlines()[0]
                        if re.match(r"\d+\.", title) and "できたことの確かめ方" not in sec and "参考" not in title and "始める前に" not in title \
                                and "症状から探す" not in title and "見方" not in title and "対処のしかた" not in title:
                            probs.append(f"{rel}: 「{title}」の段に「できたことの確かめ方」がありません")
    for dp, _, fns in os.walk(SRC):
        for fn in fns:
            if fn.endswith(".asp"):
                body = strip_vb_comments(open(os.path.join(dp, fn), encoding="utf-8").read())
                for w in BANNED:
                    if w in body:
                        probs.append(f"src/{fn}: 画面に出る文字に「{w}」があります")
    if probs:
        for p in probs:
            print("NG:", p)
        raise Fail(f"手順書の検査で {len(probs)} 件")
    return checks


if __name__ == "__main__":
    try:
        n = run()
    except Fail as e:
        print("NG:", e)
        sys.exit(1)
    print(f"OK: 手順書の検査（「　」の言葉 {n} か所が画面と一致・使わない言葉なし・各段に確かめ方あり）")
