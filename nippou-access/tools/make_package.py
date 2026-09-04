# -*- coding: utf-8 -*-
"""納品用の zip を作る。

GitHub を使わない相手に一式を渡すためのもの。
開発リポジトリの並びではなく、「誰が何を見ればよいか」で番号を振り直す。

  00_はじめにお読みください.txt   最初に開く案内 (Shift_JIS。メモ帳で開ける)
  はじめに.html                  オフラインで動く入口ページ
  01_マニュアル/                  利用者・担当者向け
  02_データベース_Access/         Access に読み込ませる VBA と初期データ
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


def copytree(src, dst, skip=()):
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in skip and not d.startswith(".")]
        for f in files:
            if f.startswith("."):
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
    d = os.path.join(base, "02_データベース_Access")
    for f in sorted(os.listdir(os.path.join(HERE, "dist"))):
        copy(os.path.join(HERE, "dist", f), os.path.join(d, "VBAモジュール", f))
    copytree(os.path.join(HERE, "data"), os.path.join(d, "マスタ初期データ"))

    order = """電話応対日報 集計システム ― Access ファイルの作り方

Access が入っているパソコンで、1 回だけ行う作業です。
全部で 10 分ほどです。


------------------------------------------------------------
 1. 空の Access ファイルを作る
------------------------------------------------------------

 (1) Access を起動します。

 (2)「空のデータベース」を選びます。

 (3) ファイル名を   日報集計_be.accdb   にして「作成」を押します。


------------------------------------------------------------
 2. VBA の画面を開く
------------------------------------------------------------

 (4) キーボードの Alt キーを押しながら F11 キーを押します。

     「Microsoft Visual Basic for Applications」という
     別のウィンドウが開きます。以降はこのウィンドウで操作します。


------------------------------------------------------------
 3. モジュールを読み込む（7 回くり返します）
------------------------------------------------------------

 (5) メニューの「ファイル」→「ファイルのインポート」を選びます。
     （Ctrl キーを押しながら M でも開きます）

 (6)「VBAモジュール」フォルダの中の

         modSetup.bas

     を選んで「開く」を押します。

 (7) (5)(6) を残りの 6 個についてもくり返します。

         modSetupMaster.bas
         modSetupQuery.bas
         modSetupUI.bas
         modSetupReport.bas
         modApp.bas
         modImportExcel.bas

     ★ 順番は問いません。7 つすべて読み込めていれば大丈夫です。

 (8) 左側の一覧に「標準モジュール」という項目ができ、
     その下に 7 つの名前が並んでいることを確認します。

         modApp
         modImportExcel
         modSetup
         modSetupMaster
         modSetupQuery
         modSetupReport
         modSetupUI

     ★ 名前を付け直す必要はありません。
       ファイルの中に名前が書いてあるので、自動的に付きます。


------------------------------------------------------------
 4. 実行する
------------------------------------------------------------

 (9) 左の一覧の「modSetup」をダブルクリックして開きます。

 (10) 画面の中から

          Public Sub Setup_DBOnly()

      という行を探し、その行の上でクリックします。
      （カーソルがその行にあれば大丈夫です）

 (11) F5 キーを押します。

      ★ Setup_DBOnly です。Setup_All ではありません。
        Setup_All は Access だけで使う場合のもので、
        今回は使わない画面まで作ってしまいます。

 (12)「データベースを作成しました」と出れば完了です。
      1 分ほどかかります。


------------------------------------------------------------
 5. 置き場所を決める
------------------------------------------------------------

 (13) Access を閉じます。

 (14) できた 日報集計_be.accdb を共有フォルダに置きます。

      この場所を、あとで Web サーバー側の設定に書きます。
      （03_Webサイト_ASP の 設置手順.txt を参照）

 (15) 過去の Excel データを取り込む場合は、この時点で
      Access を開き、メニューの「Excel から取込」を実行します。
      （01_マニュアル の 05_設置手順 に詳しく書いてあります）


------------------------------------------------------------
 うまくいかないとき
------------------------------------------------------------

 ●「ユーザー定義型は定義されていません」と出る

    VBA の画面で「ツール」→「参照設定」を開き、

        Microsoft Office xx.x Access database engine Object Library

    にチェックを入れて「OK」を押し、もう一度 F5 を押してください。


 ●「ファイルのインポート」が見当たらない

    VBA の画面（Alt + F11 で開いたウィンドウ）のメニューです。
    Access 本体のメニューではありません。


 ● コピー＆貼り付けでやりたい

    できますが、おすすめしません。
    .bas ファイルの 1 行目

        Attribute VB_Name = "modSetup"

    はコピー＆貼り付けするとエラーになります（この行は手で入力できない
    決まりになっています）。貼り付ける場合は 2 行目から選んでコピーし、
    貼り付けたあとに、左の一覧でモジュール名を手で付け直してください。

    インポートならこの手間はありません。


 ● 文字化けした

    「VBAモジュール」フォルダの .bas は Shift_JIS で保存してあります。
    別の場所のファイルを使っていないか確認してください。
"""
    with open(os.path.join(d, "Accessファイルの作り方.txt"), "wb") as f:
        f.write(order.replace("\n", "\r\n").encode("cp932"))

    # --- 03 Web サイト (ASP) ---
    w = os.path.join(base, "03_Webサイト_ASP")
    copytree(os.path.join(HERE, "web"), os.path.join(w, "wwwroot"),
             skip=("_report_preview.html",))
    pv = os.path.join(w, "wwwroot", "_report_preview.html")
    if os.path.exists(pv):
        os.remove(pv)
    os.remove(os.path.join(w, "wwwroot", "README.md"))

    setup = """電話応対日報 集計システム ― Web サーバー(IIS)への設置

くわしい手順は 01_マニュアル\\05_設置手順_情報システム担当向け.md にあります。
ここは要点だけです。

1. IIS で次の機能を有効にします。
     ・ASP           (Web サーバー → アプリケーション開発 → ASP)
     ・Windows 認証  (Web サーバー → セキュリティ → Windows 認証)

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

5. 次の 2 か所を書き換えます。

   include\\db.asp の先頭
     Const DB_PATH = "D:\\nippou\\data\\日報集計_be.accdb"
       → 実際に置いた場所に直します。

   include\\auth.asp の先頭
     Const ADMIN_USERS = ""
       → マスタ保守を触れる人のログオン名を「,」区切りで入れます。
          例) Const ADMIN_USERS = "t-okada,y-fujita"

     ★ 空のままだと全員がマスタ保守を触れます。運用開始前に必ず設定してください。

6. http://<サーバー名>/nippou/ を開いて動作を確認します。

7. 利用者のデスクトップに、上のアドレスへのショートカットを配ります。
   ITに不慣れな方が多いので、アドレスを手で打たせない形にしてください。

--------------------------------------------------------------------
編集するときの注意

・.asp ファイルは UTF-8 (BOM なし) で保存されています。
  Shift_JIS で保存し直すと文字化けします。

・SQL は必ず DbQuery / DbExec のパラメータ経由で書いてください。
  画面から来た値を文字列でつなげて SQL に埋めないでください。
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

・帳票プレビューで「印刷する」
    → A4 縦 1 枚で出ます。画面の青い帯やボタンは印刷されません。
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

    write_readme_txt(base)
    write_index_html(base)
    return base


def write_readme_txt(base):
    """最初に開く案内。メモ帳で文字化けしないよう Shift_JIS + CRLF で書く。"""
    txt = """============================================================
 電話応対日報 集計システム   納品一式
============================================================

 このフォルダを、共有フォルダなど分かりやすい場所に置いてください。

 まず「はじめに.html」をダブルクリックしてください。
 どのフォルダに何が入っているかが、絵入りで開きます。
 （インターネットにつながっていなくても開けます）

------------------------------------------------------------
 フォルダの中身
------------------------------------------------------------

 01_マニュアル
     利用者・担当者向けの操作マニュアルです。
     「操作マニュアル.html」を開いて「印刷する」を押すと、
     A4 縦 20 ページで出ます。綴じて備え付けにしてください。

     ・パート職員のかた   → 2章 だけ読めば足ります
     ・職員のかた         → 3章
     ・担当者のかた       → 4章
     ・困ったとき         → 5章

 02_データベース_Access
     Access に読み込ませる VBA と、マスタの初期データです。
     「Accessファイルの作り方.txt」の手順どおりに進めてください。
     この作業は 1 回だけ、10 分ほどで終わります。

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

  1. 04_モックアップ を開いて、画面と操作を確認する
  2. 02_データベース_Access の手順で Access ファイルを作る
  3. 03_Webサイト_ASP の手順で IIS に置く
  4. 01_マニュアル を印刷して配る
  5. 利用者のデスクトップにショートカットを配る

------------------------------------------------------------
 設置前に必ず設定する箇所（2 か所）
------------------------------------------------------------

  03_Webサイト_ASP\\wwwroot\\include\\db.asp
      DB_PATH … Access ファイルを置いた実際の場所

  03_Webサイト_ASP\\wwwroot\\include\\auth.asp
      ADMIN_USERS … マスタ保守を触れる人のログオン名
                    空のままだと全員が触れます

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
  <h2>まず開くもの</h2>
  <div class="cards">
    <a class="card go" href="04_モックアップ/モックアップ.html">
      <div class="n">04</div><b>モックアップを触ってみる</b>
      <span>本番と同じ画面がブラウザだけで動きます。何をしても本番には影響しません。</span></a>
    <a class="card go" href="01_マニュアル/操作マニュアル.html">
      <div class="n">01</div><b>操作マニュアルを読む・印刷する</b>
      <span>役割ごとの章立て。「印刷する」で A4 縦 20 ページ。</span></a>
  </div>

  <h2>導入する（担当者・情報システム）</h2>
  <div class="cards">
    <a class="card" href="02_データベース_Access/Accessファイルの作り方.txt">
      <div class="n">02</div><b>データベースを作る</b>
      <span>Access に VBA を読み込ませて実行します。1 回だけ、10 分ほどの作業です。</span></a>
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
    <li><b>02</b> の手順で Access ファイルを作り、共有フォルダに置く</li>
    <li><b>03</b> の手順で IIS に置き、設定を 2 か所書き換える</li>
    <li><b>01</b> のマニュアルを印刷して配る</li>
    <li>利用者のデスクトップにショートカットを配る</li>
  </ol>

  <div class="note crit">
    <b>設置前に必ず設定する箇所が 2 つあります。</b><br>
    <code>03_Webサイト_ASP\wwwroot\include\db.asp</code> の <code>DB_PATH</code>
    … Access ファイルを置いた実際の場所<br>
    <code>03_Webサイト_ASP\wwwroot\include\auth.asp</code> の <code>ADMIN_USERS</code>
    … マスタ保守を触れる人のログオン名。<b>空のままだと全員が触れます。</b>
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
            if f.endswith((".png", ".zip", ".jpg")):
                continue
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
