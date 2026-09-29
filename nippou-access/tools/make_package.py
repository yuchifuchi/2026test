# -*- coding: utf-8 -*-
"""納品用の zip を作る。

GitHub を使わない相手に一式を渡すためのもの。
開発リポジトリの並びではなく、「誰が何を見ればよいか」で番号を振り直す。

  00_はじめにお読みください.txt   最初に開く案内 (Shift_JIS。メモ帳で開ける)
  はじめに.html                  オフラインで動く入口ページ
  01_マニュアル/                  利用者・担当者向け
  02_データベース_Access/         日報集計_be.accdb（作成済み。置くだけ）
  03_Webサイト_ASP/               IIS に置くファイル一式
  04_モックアップ/                評価・研修用
  05_設計資料/                    調査結果と設計
  06_開発ツール/                  作り直すときに使う
"""
import io
import os
import shutil
import sys
import zipfile
from datetime import date

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NAME = "電話応対日報_納品一式"


def copy(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)


SKIP_DIRS = {"__pycache__", ".git", ".idea", "node_modules", "_classes"}
SKIP_EXT = (".pyc", ".pyo", ".bak", ".tmp", ".class")
SKIP_FILES = ("_schema.json", "_schema_test.json", "_cases.json")


def copytree(src, dst, skip=()):
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs
                   if d not in skip and d not in SKIP_DIRS and not d.startswith(".")]
        for f in files:
            if f.startswith(".") or f.endswith(SKIP_EXT) or f in SKIP_FILES:
                continue
            s = os.path.join(root, f)
            copy(s, os.path.join(dst, os.path.relpath(s, src)))


def build(out_root):
    base = os.path.join(out_root, NAME)
    if os.path.exists(base):
        shutil.rmtree(base)
    os.makedirs(base)

    # --- 01 マニュアル ---
    m = os.path.join(base, "01_マニュアル")
    copytree(os.path.join(HERE, "docs", "manual"), m)
    os.remove(os.path.join(m, "README.md"))

    # --- 02 データベース (Access) ---
    #   .accdb は作った状態で入れる。置くだけで使えるようにするため。
    d = os.path.join(base, "02_データベース_Access")
    accdb = os.path.join(HERE, "dist", "日報集計_be.accdb")
    if not os.path.exists(accdb):
        raise SystemExit("dist/日報集計_be.accdb がありません。"
                         "先に python3 tools/gen_accdb.py を実行してください。")
    copy(accdb, os.path.join(d, "日報集計_be.accdb"))
    copy(os.path.join(HERE, "dist", "データベースを作る.vbs"),
         os.path.join(d, "任意_作り直すとき", "データベースを作る.vbs"))
    copy(os.path.join(HERE, "dist", "ログオン名の欄を足す.vbs"),
         os.path.join(d, "任意_すでに使い始めている場合", "ログオン名の欄を足す.vbs"))
    for f in sorted(os.listdir(os.path.join(HERE, "dist"))):
        if f.endswith(".bas"):
            copy(os.path.join(HERE, "dist", f),
                 os.path.join(d, "任意_過去データ取込に使うVBA", f))
    copytree(os.path.join(HERE, "data"), os.path.join(d, "マスタ初期データ"))

    order = """電話応対日報 集計システム ― データベース（Access ファイル）

------------------------------------------------------------
 このフォルダでやること（1 つだけ）
------------------------------------------------------------

 このフォルダの

     日報集計_be.accdb

 を、共有フォルダに移してください。それだけです。


 ★ 作る作業はもう済んでいます。
    表・マスタ・つながりまで入った状態でお渡ししています。
    Access をお持ちでなくても構いません。開く必要もありません。


 置いた場所（フルパス）をメモしておいてください。
 あとで Web サーバーの設定 (include\\db.asp) に書きます。

     例) \\\\fsx.sales.mint.go.jp\\share\\share_電話オペレータ\\日報\\


------------------------------------------------------------
 置き場所を決めるときの注意
------------------------------------------------------------

 ・毎日バックアップされるフォルダにしてください。
   このファイル 1 つに、すべてのデータが入ります。

 ・Web サーバーから読み書きできるフォルダにしてください。
   「読み取り」だけでは足りません（後述の .laccdb を作るため）。

 ・同じフォルダに 日報集計_be.laccdb というファイルが自動でできます。
   誰かが使っている間だけできる目印です。消さないでください。


------------------------------------------------------------
 中身について（読まなくても大丈夫です）
------------------------------------------------------------

   ・表          11 個（担当者・製品・区分などのマスタと、実績）
   ・初期データ  158 件（製品 34 種・区分 57 種・担当者 18 名など）
   ・表どうしのつながり 8 個（担当者や区分を間違って消せないようにする仕掛け）

 実績（受付の記録）は空です。使い始めた日から貯まっていきます。

 このファイルには VBA（プログラム）が入っていません。
 開いてもマクロの警告は出ません。


------------------------------------------------------------
 任意_作り直すとき
------------------------------------------------------------

 いまのファイルを捨てて、まっさらな状態から作り直したいときに使います。
 ふだんは開かなくて構いません。

 Access が入っているパソコンで データベースを作る.vbs をダブルクリックすると、
 同じ内容の 日報集計_be.accdb をその場に作ります（10 秒ほど）。

 ★ 作り直すと、入力済みのデータはすべて消えます。


------------------------------------------------------------
 任意_すでに使い始めている場合
------------------------------------------------------------

 すでに受付入力を始めたあとで、新しい版をお渡しした場合に使います。
 データベースを差し替えると入力済みのデータが消えるので、
 代わりに「足りない欄だけを足す」ためのものです。

 ログオン名の欄を足す.vbs
     マスタ保守で「ログオン名」を入れられるようにします。
     日報集計_be.accdb をこのファイルの上にドラッグして落とすか、
     ダブルクリックして、出てきた窓に場所を貼り付けてください。
     Access は要りません。何度実行しても大丈夫です。


------------------------------------------------------------
 過去の Excel データを取り込む場合（任意）
------------------------------------------------------------

 いままでの「集計表（…）.xlsm」を取り込みたいときだけの作業です。
 新しく使い始めるだけなら、この作業は要りません。

 (1) 日報集計_be.accdb を Access で開きます。

 (2) Alt キーを押しながら F11 キーを押します。

 (3) メニューの「ファイル」→「ファイルのインポート」で、
     「任意_過去データ取込に使うVBA」フォルダの中の

         modSetup.bas
         modImportExcel.bas

     の 2 つを読み込みます。

 (4) modImportExcel を開き、Import_Excel_Dialog の行にカーソルを置いて
     F5 キーを押します。

 (5) 取り込む「集計表（…）.xlsm」を選びます。
     週ごとのファイルがある分だけくり返します。

 ※ 取り込んだあと、担当者名が「顧客G」になっている行の振り分けが必要です。
    詳しくは 05_設計資料 の 03_導入手順.md をご覧ください。
"""
    with open(os.path.join(d, "このフォルダの使い方.txt"), "wb") as f:
        f.write(order.replace("\n", "\r\n").encode("cp932"))

    # --- 03 Web サイト (ASP) ---
    w = os.path.join(base, "03_Webサイト_ASP")
    copytree(os.path.join(HERE, "web"), os.path.join(w, "wwwroot"))
    # 検証用に作った静的プレビューは納品物ではないので落とす
    for f in ("_report_preview.html", "_report_preview_pdf.html", "README.md"):
        pv = os.path.join(w, "wwwroot", f)
        if os.path.exists(pv):
            os.remove(pv)

    setup = """電話応対日報 集計システム ― Web サーバー(IIS)への設置

くわしい手順は 01_マニュアル\\05_設置手順_情報システム担当向け.md にあります。
ここは要点だけです。

1. IIS で次の機能を有効にします。
     ・ASP           (Web サーバー → アプリケーション開発 → ASP)
     ・Windows 認証  (Web サーバー → セキュリティ → Windows 認証)

     そのあと IIS マネージャーで、対象サイト (またはアプリケーション) を選び

       「認証」   → Windows 認証 = 有効 ／ 匿名認証 = 無効
       「ASP」    → デバッグのプロパティ
                    → ブラウザーにエラーを送信する = True
                      (設置作業のあいだだけ。運用開始後は False に戻す)

     ★ この 2 つは web.config に書けません。
       書くとサイト全体が「HTTP 500.19 構成データが無効です」になり、
       どのページも開けなくなります。必ず画面から設定してください。

2. Microsoft Access Database Engine 2016 Redistributable を入れます。

     ★ アプリケーションプールのビット数を ACE と合わせてください。
       64bit の ACE を入れたなら「32 ビット アプリケーションの有効化」は False です。
       ここが合っていないと「Provider が見つかりません」というエラーになります。

3. wwwroot フォルダの中身を、そのまま C:\\inetpub\\wwwroot\\nippou\\ に置きます。
   web.config も一緒に置いてください。

4. 日報集計_be.accdb を置いた「フォルダ」に、
   アプリケーションプール ID (既定は IIS AppPool\\<プール名>) の
   「変更」権限を与えます。

     ★ 読み取りだけでは動きません。
       ACE がロックファイル (.laccdb) を同じフォルダに作るためです。

5. 次の 3 か所を書き換えます。

   ★ 設置ごとの設定は include\\config.asp の 1 ファイルにまとまっています。
     次回以降、直しをお渡ししたときも、このファイルだけ上書きしなければ
     設定は残ります。

   include\\config.asp
     Const DB_PATH = "D:\\nippou\\data\\日報集計_be.accdb"
       → 実際に置いた場所に直します。

   同じ include\\config.asp の中
     Const STAFF_PATH = ""
       → パート職員と職員を「アドレス」で分ける場合に書きます。

          (1) IIS マネージャーで、いまの nippou と同じフォルダを指す
              アプリケーションをもう 1 つ作ります。
                 エイリアス  nippou-staff
                 物理パス    nippou と同じフォルダ (コピーはしません)

          (2) 左のツリーで nippou-staff を選び、
              「IP アドレスおよびドメインの制限」を開きます。
                 機能設定の編集 → 未指定のクライアントのアクセス = 拒否
                 許可エントリの追加 → 職員の PC の IP を追加

          (3) ここに Const STAFF_PATH = "/nippou-staff" と書きます。

          これで、職員用アドレスから開いたときだけ職員として扱われます。
          パート職員用のアドレスで日報などを直接叩いても開けません。

       ★ IP 制限は web.config に書けません。IIS マネージャーから
         設定してください (書くと HTTP 500.19 になります)。

   同じ include\\config.asp の中
     Const STAFF_USERS = ""
       → 最初にマスタ保守を開くかた 1 人のログオン名を入れます。
          ドメイン名は書きません。例) Const STAFF_USERS = "t-okada"

       誰が職員かは、ふだんは「マスタ保守」の画面で決めます。
       担当者一覧の「ログオン名」に Windows のログオン名を入れ、
       区分を「職員」にすると、その方に日報・帳票印刷・集計表・入力もれ・
       マスタ保守が出るようになります。ファイルを直す必要はありません。

     ★ マスタのログオン名が 1 件も入っておらず、ここも空のままだと
       全員が職員扱いです。運用開始前に必ずどちらかを設定してください。

   同じ include\\config.asp の中 (帳票を PDF で印刷するため)
     Const PDF_EDGE_EXE = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
       → サーバーの中で msedge.exe がこの場所にあるか確認します。
          違う場所にあるときだけ書き換えてください。
          Edge が入っていないサーバーでは wkhtmltopdf を入れて
          Const PDF_ENGINE = "wkhtmltopdf" にします。

     ★ 作業用フォルダ (既定は Windows の一時フォルダ) にも、
       アプリケーションプール ID の「変更」権限が要ります。
       作った PDF は送った直後に消すので、たまることはありません。

6. まず http://<サーバー名>/nippou/setup_check.asp を開きます。

     「設置チェック」の画面が出て、つまずいている箇所が一覧で分かります。
     NG が無くなったら http://<サーバー名>/nippou/ を開いてください。

     ★ 真っ白な画面や「このページは現在機能していません」が出るときは、
       サーバー本体のブラウザから
         http://localhost/nippou/setup_check.asp
       を開いてください。サーバーの中から開いたときだけ、
       IIS が詳しい原因を表示します。

7. 利用者のデスクトップに、上のアドレスへのショートカットを配ります。
   ITに不慣れな方が多いので、アドレスを手で打たせない形にしてください。

--------------------------------------------------------------------
編集するときの注意

・.asp ファイルは UTF-8 (BOM なし) で保存されています。
  Shift_JIS で保存し直すと文字化けします。

・SQL は必ず DbQuery / DbExec のパラメータ経由で書いてください。
  画面から来た値を文字列でつなげて SQL に埋めないでください。

・setup_check.asp は設置が終わったら消してかまいません。
  (誰でも開ける画面なので、置いたままにするならその点だけご承知おきください)
"""
    with open(os.path.join(w, "設置手順.txt"), "wb") as f:
        f.write(setup.replace("\n", "\r\n").encode("cp932"))

    # --- 04 モックアップ ---
    k = os.path.join(base, "04_モックアップ")
    copy(os.path.join(HERE, "mockup", "nippou-mockup.html"),
         os.path.join(k, "モックアップ.html"))
    mock = """電話応対日報 集計システム ― モックアップ

「モックアップ.html」をダブルクリックすると、ブラウザで開きます。
サーバーもインターネットも要りません。

本番と同じ画面で、実際のデータを入れてあります。
導入前の確認や、操作の練習にお使いください。
このファイルの中だけで動いているので、何をしても本番には影響しません。
（ページを開き直すと元に戻ります）

試していただきたいこと

・受付入力で件数を足す
    → 画面上部の帯の「日報」と「集計表」が同時に同じ数だけ動きます。
      数字がずれない仕組みを目で確かめられます。

・日報の件数欄を直そうとする
    → 直せません。自動集計だけです。ここが今までとの一番の違いです。

・マスタ保守 →「区分」で区分を 1 つ足す
    → 受付入力の選択肢にすぐ出ます。Excel を配り直す必要がありません。

・使用件数が 1 以上の行で「削除」を押す
    → 拒否されます。「有効を外してください」と案内が出ます。

・担当者に退職日を入れる
    → 入力候補から消えます。過去のデータは変わりません。

・帳票プレビューで「PDF で印刷する」
    → 本番ではサーバーが PDF を作って別のタブに開きます。
      A4 縦 1 枚で、紙に日付や URL は入りません。
      (モックアップにはサーバーが無いので、ブラウザの印刷画面が出ます)

・右上の「いま見ている立場」を「パート職員」に切り替える
    → 上の帯が「受付入力」「その他業務」だけになります。
      本番ではログオン名で自動的に決まります。
"""
    with open(os.path.join(k, "使い方.txt"), "wb") as f:
        f.write(mock.replace("\n", "\r\n").encode("cp932"))

    # --- 05 設計資料 ---
    g = os.path.join(base, "05_設計資料")
    for f in sorted(os.listdir(os.path.join(HERE, "docs"))):
        if f.endswith(".md"):
            copy(os.path.join(HERE, "docs", f), os.path.join(g, f))

    # --- 06 開発ツール ---
    t = os.path.join(base, "06_開発ツール")
    copytree(os.path.join(HERE, "tools"), os.path.join(t, "tools"))
    copytree(os.path.join(HERE, "src"), os.path.join(t, "VBA原本_UTF8"))
    copy(os.path.join(HERE, "mockup", "app.js"), os.path.join(t, "モックアップ素材", "app.js"))
    copy(os.path.join(HERE, "mockup", "body.html"), os.path.join(t, "モックアップ素材", "body.html"))
    copy(os.path.join(HERE, "mockup", "head.html"), os.path.join(t, "モックアップ素材", "head.html"))
    copy(os.path.join(HERE, "mockup", "data.json"), os.path.join(t, "モックアップ素材", "data.json"))
    copy(os.path.join(HERE, "README.md"), os.path.join(t, "開発者向けREADME.md"))

    copy(os.path.join(HERE, "docs", "manual", "やさしい導入手順書.html"),
         os.path.join(base, "はじめに読む_やさしい導入手順書.html"))

    write_readme_txt(base)
    write_index_html(base)
    return base


def write_readme_txt(base):
    """最初に開く案内。メモ帳で文字化けしないよう Shift_JIS + CRLF で書く。"""
    txt = """============================================================
 電話応対日報 集計システム   納品一式
============================================================

 このフォルダを、共有フォルダなど分かりやすい場所に置いてください。

 導入されるかたは

     はじめに読む_やさしい導入手順書.html

 をダブルクリックしてください。
 使えるようにするまでの手順が、1 から 10 まで絵入りで並んでいます。
 むずかしい言葉には、そのつど説明をつけました。

 どのフォルダに何が入っているかを先に見たいときは
 「はじめに.html」を開いてください。
 （どちらもインターネットにつながっていなくても開けます）

------------------------------------------------------------
 フォルダの中身
------------------------------------------------------------

 01_マニュアル
     利用者・担当者向けの操作マニュアルです。
     「操作マニュアル.html」を開いて「印刷する」を押すと、
     A4 縦 25 ページで出ます。綴じて備え付けにしてください。

     ・パート職員のかた   → 2章 だけ読めば足ります
     ・職員のかた         → 3章
     ・担当者のかた       → 4章
     ・困ったとき         → 5章

 02_データベース_Access
     データベース（日報集計_be.accdb）です。もうできています。
     共有フォルダに移すだけです。詳しくは
     「このフォルダの使い方.txt」をご覧ください。

 03_Webサイト_ASP
     Web サーバー(IIS)に置くファイル一式です。
     「設置手順.txt」を参照してください。

 04_モックアップ
     本番と同じ画面が、ブラウザだけで動きます。
     導入前の確認や、操作の練習にお使いください。
     何をしても本番には影響しません。

 05_設計資料
     なぜ作り替えたのか、どういう仕組みかをまとめた資料です。
     引き継ぎのときにお読みください。

 06_開発ツール
     作り直すときに使う道具です。ふだんは開かなくて構いません。

------------------------------------------------------------
 導入の流れ（担当者のかた）
------------------------------------------------------------

  「はじめに読む_やさしい導入手順書.html」に、1 から 10 まで書いてあります。
  ざっくりは次のとおりです。

  1. 04_モックアップ を開いて、画面と操作を確認する
  2. 02_データベース_Access の 日報集計_be.accdb を共有フォルダに置く（1 分）
  3. 03_Webサイト_ASP のファイルを IIS に置き、3 か所だけ書き換える（30 分）
  4. ブラウザで開いて動作を確認する（5 分）
  5. 01_マニュアル を印刷して配り、デスクトップにショートカットを配る

------------------------------------------------------------
 設置前に必ず設定する箇所（include\\config.asp の中）
------------------------------------------------------------

  03_Webサイト_ASP\\wwwroot\\include\\config.asp
      DB_PATH … Access ファイルを置いた実際の場所

  同じ config.asp の中
      STAFF_USERS … 最初にマスタ保守を開くかた 1 人のログオン名
                    ふだんは「マスタ保守」の担当者一覧にある
                    「ログオン名」と「区分」で決めます (画面から変更できます)。
                    どちらも空のままだと全員が職員扱いです。

  同じ config.asp の中
      PDF_EDGE_EXE … サーバーの Edge (msedge.exe) の場所
                    既定の場所にあれば、そのままで構いません。

------------------------------------------------------------
 ご注意
------------------------------------------------------------

 ・利用者のパソコンに入れるものはありません。
   Access も Excel も要りません。ブラウザだけで使えます。

 ・帳票の「電話応対以外の業務」の欄は、
   左 8 件・右 5 件の 13 個までです。
   14 個目以降は画面には出ますが、紙には載りません。

 ・Access(ACE)の同時利用は数十人規模までです。
   それを超える場合は SQL Server への移行をご検討ください。
"""
    with open(os.path.join(base, "00_はじめにお読みください.txt"), "wb") as f:
        f.write(txt.replace("\n", "\r\n").encode("cp932"))


def write_index_html(base):
    """オフラインで開ける入口ページ。外部のフォントや画像は使わない。"""
    html = r"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>電話応対日報 集計システム ― 納品一式</title>
<style>
  :root{
    --ground:#f6f5f2; --panel:#fff; --ink:#17202c; --muted:#5b6675;
    --line:#dcd8d0; --accent:#1f4e79; --accent-soft:#e8eff6;
    --admin:#8a6a2c; --admin-soft:#f4eddc; --crit:#a02c2c;
  }
  *{box-sizing:border-box}
  body{
    margin:0; background:var(--ground); color:var(--ink);
    font-family:"Meiryo","Yu Gothic UI","MS PGothic",system-ui,sans-serif;
    font-size:16px; line-height:1.85;
  }
  header{background:var(--panel); border-bottom:1px solid var(--line); padding:40px 28px 30px}
  .in{max-width:960px; margin:0 auto}
  .kicker{font-size:12px; letter-spacing:.2em; color:var(--admin); font-weight:700}
  h1{font-size:clamp(26px,4vw,38px); margin:8px 0 12px; letter-spacing:.04em}
  header p{max-width:62ch; color:var(--muted); margin:0; font-size:15px}
  main{max-width:960px; margin:0 auto; padding:30px 28px 80px}
  h2{font-size:15px; letter-spacing:.1em; color:var(--muted); margin:34px 0 14px}
  .cards{display:grid; gap:14px; grid-template-columns:repeat(auto-fit,minmax(280px,1fr))}
  .card{
    display:block; background:var(--panel); border:1px solid var(--line);
    border-radius:8px; padding:20px 22px; text-decoration:none; color:var(--ink);
  }
  .card:hover{border-color:var(--accent); background:var(--accent-soft)}
  .card .n{font-size:12px; color:var(--muted); letter-spacing:.14em}
  .card b{display:block; font-size:18px; color:var(--accent); margin:3px 0 6px}
  .card span{font-size:14px; color:var(--muted)}
  .card.go{border-color:var(--accent); border-width:2px}
  ol{padding-left:1.3em} li{margin-bottom:7px}
  .note{
    border-left:4px solid var(--admin); background:var(--admin-soft);
    padding:14px 18px; border-radius:0 8px 8px 0; margin:22px 0; font-size:15px;
  }
  .note.crit{border-left-color:var(--crit); background:#f8e9e9}
  .note b{color:var(--admin)} .note.crit b{color:var(--crit)}
  code{background:#eceae4; padding:1px 6px; border-radius:4px; font-size:.9em}
  footer{
    max-width:960px; margin:0 auto; padding:20px 28px 60px;
    border-top:1px solid var(--line); color:var(--muted); font-size:13px;
  }
</style></head><body>

<header><div class="in">
  <div class="kicker">電話応対報告書 日報集計</div>
  <h1>納品一式</h1>
  <p>
    現行の Excel マクロ一式を置き換えるシステムです。
    データベースは Access、画面は社内 Web（ASP）です。
    利用者のパソコンに入れるものはありません。
  </p>
</div></header>

<main>
  <h2>導入するかたへ ― まずここから</h2>
  <div class="cards">
    <a class="card go" href="はじめに読む_やさしい導入手順書.html">
      <div class="n">START</div><b>やさしい導入手順書</b>
      <span>使えるようにするまでを、1 から 10 の手順で。むずかしい言葉には説明つき。
            印刷して持ち歩けます。</span></a>
  </div>

  <h2>使うかたへ</h2>
  <div class="cards">
    <a class="card go" href="04_モックアップ/モックアップ.html">
      <div class="n">04</div><b>モックアップを触ってみる</b>
      <span>本番と同じ画面がブラウザだけで動きます。何をしても本番には影響しません。</span></a>
    <a class="card go" href="01_マニュアル/操作マニュアル.html">
      <div class="n">01</div><b>操作マニュアルを読む・印刷する</b>
      <span>役割ごとの章立て。「印刷する」で A4 縦 25 ページ。</span></a>
  </div>

  <h2>導入の細かい手順</h2>
  <div class="cards">
    <a class="card" href="02_データベース_Access/このフォルダの使い方.txt">
      <div class="n">02</div><b>データベースを共有フォルダに置く</b>
      <span>日報集計_be.accdb はもうできています。移すだけです。Access は要りません。</span></a>
    <a class="card" href="03_Webサイト_ASP/設置手順.txt">
      <div class="n">03</div><b>Web サーバーに置く</b>
      <span>IIS の設定とファイルの配置。</span></a>
  </div>

  <h2>そのほか</h2>
  <div class="cards">
    <a class="card" href="05_設計資料/01_現行調査と不具合原因.md">
      <div class="n">05</div><b>設計資料</b>
      <span>なぜ作り替えたのか、どういう仕組みか。引き継ぎ用。</span></a>
    <a class="card" href="06_開発ツール/開発者向けREADME.md">
      <div class="n">06</div><b>開発ツール</b>
      <span>作り直すときに使う道具。ふだんは開かなくて構いません。</span></a>
  </div>

  <h2>導入の流れ</h2>
  <ol>
    <li><b>04</b> モックアップで画面と操作を確認する</li>
    <li><b>02</b> の <code>日報集計_be.accdb</code> を共有フォルダに置く（作る作業は済んでいます）</li>
    <li><b>03</b> の手順で IIS に置き、設定を 3 か所書き換える</li>
    <li><b>01</b> のマニュアルを印刷して配る</li>
    <li>利用者のデスクトップにショートカットを配る</li>
  </ol>

  <div class="note crit">
    <b>設置前に設定するのは <code>include\config.asp</code> の 1 ファイルだけです。</b><br>
    <code>03_Webサイト_ASP\wwwroot\include\config.asp</code> の <code>DB_PATH</code>
    … Access ファイルを置いた実際の場所<br>
    同じ <code>config.asp</code> の <code>STAFF_USERS</code>
    … 最初にマスタ保守を開くかた 1 人のログオン名。
    ふだんは「マスタ保守」の担当者一覧（ログオン名・区分）で決めます。
    <b>どちらも空のままだと全員が職員扱いです。</b><br>
    同じ <code>config.asp</code> の <code>PDF_EDGE_EXE</code>
    … サーバーの Edge の場所。既定の場所にあれば、そのままで構いません。
  </div>

  <div class="note">
    <b>知っておいていただきたい制約。</b><br>
    帳票の「電話応対以外の業務」の欄は左 8 件・右 5 件の <b>13 個まで</b>です。
    14 個目以降は画面には出ますが紙には載りません。<br>
    Access（ACE）の同時利用は数十人規模までです。
  </div>
</main>

<footer>
  電話応対日報 集計システム 納品一式 ／
  このページはインターネットにつながっていなくても開けます。
</footer>
</body></html>
"""
    io.open(os.path.join(base, "はじめに.html"), "w", encoding="utf-8").write(html)


BAD_BYTES = {0: "NUL", 7: r"BEL (\a)", 8: r"BS (\b)",
             11: r"VT (\v)", 12: r"FF (\f)", 27: "ESC"}


def check(base):
    """納品物に制御文字が混ざっていないか調べる。

    Python の文字列リテラルで \auth や \nippou と書くと、
    \a や \n がエスケープとして解釈されて制御文字になる。
    Windows のパス表記を埋め込むファイルで起きやすいので、必ず通す。
    """
    problems = []
    for root, _, files in os.walk(base):
        for f in files:
            p = os.path.join(root, f)
            if f.endswith((".png", ".zip", ".jpg", ".accdb")):
                continue  # バイナリは対象外
            data = open(p, "rb").read()
            for code, name in BAD_BYTES.items():
                if bytes([code]) in data:
                    problems.append("%s に %s が %d 個"
                                    % (os.path.relpath(p, base), name,
                                       data.count(bytes([code]))))
    return problems


def zip_dir(base, zip_path):
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for root, dirs, files in os.walk(base):
            dirs.sort()
            for f in sorted(files):
                p = os.path.join(root, f)
                arc = os.path.relpath(p, os.path.dirname(base))
                info = zipfile.ZipInfo(arc.replace(os.sep, "/"))
                info.compress_type = zipfile.ZIP_DEFLATED
                # 日本語のファイル名を Windows のエクスプローラーで正しく開けるよう
                # UTF-8 フラグ (0x800) を立てる
                info.flag_bits |= 0x800
                info.date_time = (2026, 8, 26, 12, 0, 0)
                info.external_attr = 0o644 << 16
                with open(p, "rb") as fh:
                    z.writestr(info, fh.read())


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "/tmp"
    base = build(out)

    problems = check(base)
    if problems:
        print("制御文字が混入しています。zip は作りません。")
        for x in problems:
            print("  !", x)
        sys.exit(1)

    zp = os.path.join(out, NAME + ".zip")
    zip_dir(base, zp)
    n = sum(len(f) for _, _, f in os.walk(base))
    print(f"作成: {zp}")
    print(f"  {n} ファイル / {os.path.getsize(zp)/1024/1024:.2f} MB")
