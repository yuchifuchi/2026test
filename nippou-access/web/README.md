# Web 版（ASP + Access）

Access の `.accdb` をそのまま DB として使うイントラ Web サイトです。
デスクトップ版 Access と同じデータを見ます。

配置・権限・本番前にやることは [../docs/05_Web版.md](../docs/05_Web版.md) を参照してください。

## いちばん短い動かし方

1. IIS で「ASP」機能を有効にする
2. Microsoft Access Database Engine 2016 Redistributable を入れ、
   アプリケーションプールのビット数を合わせる
3. このフォルダの中身を `C:\inetpub\wwwroot\nippou\` に置く
4. `include/config.asp` の `DB_PATH` をバックエンドの実際のパスに書き換える
5. `.accdb` を置いた**フォルダ**にアプリケーションプール ID の変更権限を与える
   （ロックファイル `.laccdb` を作るため、読み取りだけでは動きません）
6. `http://<サーバー>/nippou/` を開く

## ファイル

画面は**役割で分かれています**。誰がどれを開けるかは `include/auth.asp` の
`M_担当者.ログオン名` と `職員区分`（マスタ保守の画面）で決まります。
マスタに無いときだけ `include/auth.asp` の `STAFF_USERS` を見ます。

```
【パート職員が使う画面】
default.asp      メニュー（受付入力とその他業務だけ）
entry.asp        受付入力
tasks.asp        その他業務（受注入力・架電など ①〜⑬）

【職員が使う画面】RequireStaff で入口を止めている
staff.asp        職員用メニュー
daily.asp        日報（出勤者・回線数・記述欄・確定）
report.asp       帳票プレビュー（現行「印刷用」シートの体裁）
printpdf.asp     帳票を PDF にして返す（report.asp の「PDF で印刷する」の行き先）
summary.asp      集計表
check.asp        入力もれチェック
master.asp       マスタ保守（担当者／製品／区分／業務項目）

【共通】
error.asp        エラー画面
setup_check.asp  設置チェック（ほかのファイルを読み込まないので単体で開ける）
include/config.asp  設置ごとの設定（DB の場所・職員の予備一覧・PDF）※ここだけ環境依存
include/auth.asp    ログオン名の取得と役割の判定（M_担当者 → STAFF_USERS の順に見る）
include/db.asp      接続・パラメータ化クエリ・共通関数
include/sql.asp     クエリの SQL（自動生成。Access の保存クエリは使わない）
include/layout.asp  ヘッダ・フッタ・日付ナビ（役割でメニューが変わる）
include/sheet.asp   帳票 1 枚ぶんの HTML（report.asp と printpdf.asp が共用）
include/pdf.asp     PDF 変換（Edge / wkhtmltopdf）
css/style.css       画面と印刷のスタイル
```

## 編集するときの注意

- **UTF-8（BOM なし）で保存すること。** Shift_JIS で保存し直すと文字化けします
- **SQL は必ず `DbQuery` / `DbExec` のパラメータ経由で書くこと。**
  値を文字列連結で SQL に埋めないでください
- **職員用の画面を足したら、先頭に `RequireStaff` を書くこと。**
  メニューから消すだけでは、URL を直接打てば開けてしまいます
- **帳票の中身は `include/sheet.asp` だけを直すこと。**
  画面（report.asp）と PDF（printpdf.asp）が同じ関数を呼んでいます
- **`include/sql.asp` は手で直さないこと。**
  `src/modSetupQuery.bas` から `tools/gen_sql_asp.py` が生成しています
- **Access 専用の関数（`Nz`・`DLookup` など）を SQL に書かないこと。**
  ACE 経由では使えず、「関数 'Nz' が定義されていません」で落ちます。
  `Nz(x,0)` の代わりに `IIf(IsNull(x),0,x)` を使ってください
- 認証は IIS の Windows 認証に任せています。運用前に
  「Windows 認証 = 有効／匿名認証 = 無効」にしてください

## 直したら実行する

```bash
python3 tools/check_asp.py          # <% %> の対応・呼び出し先・配列の添字を見る
python3 tools/gen_sql_asp.py        # SQL を作り直す (modSetupQuery.bas を直したとき)
python3 tools/test_sql.py           # その SQL を実際に .accdb に流してみる
python3 tools/render_asp_report.py  # 帳票を静的 HTML に起こして体裁を見る
```

IIS が無くても、この 4 つで「開いた瞬間に落ちる」類の壊れは見つかります。
