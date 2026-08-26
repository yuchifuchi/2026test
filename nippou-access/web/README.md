# Web 版（ASP + Access）

Access の `.accdb` をそのまま DB として使うイントラ Web サイトです。
デスクトップ版 Access と同じデータを見ます。

配置・権限・本番前にやることは [../docs/05_Web版.md](../docs/05_Web版.md) を参照してください。

## いちばん短い動かし方

1. IIS で「ASP」機能を有効にする
2. Microsoft Access Database Engine 2016 Redistributable を入れ、
   アプリケーションプールのビット数を合わせる
3. このフォルダの中身を `C:\inetpub\wwwroot\nippou\` に置く
4. `include/db.asp` の `DB_PATH` をバックエンドの実際のパスに書き換える
5. `.accdb` を置いた**フォルダ**にアプリケーションプール ID の変更権限を与える
   （ロックファイル `.laccdb` を作るため、読み取りだけでは動きません）
6. `http://<サーバー>/nippou/` を開く

## ファイル

```
default.asp      メニュー
entry.asp        受付入力
daily.asp        日報（出勤者・回線数・記述欄・確定）
report.asp       印刷用（現行「印刷用」シートの体裁）
summary.asp      集計表
check.asp        入力もれチェック
master.asp       マスタ保守（担当者／製品／区分／業務項目）
include/db.asp   接続・パラメータ化クエリ・共通関数
include/layout.asp  ヘッダ・フッタ・日付ナビ
css/style.css    画面と印刷のスタイル
```

## 編集するときの注意

- **UTF-8（BOM なし）で保存すること。** Shift_JIS で保存し直すと文字化けします
- **SQL は必ず `DbQuery` / `DbExec` のパラメータ経由で書くこと。**
  値を文字列連結で SQL に埋めないでください
- 認証はまだ入っていません。運用前に IIS の Windows 認証を有効にしてください
