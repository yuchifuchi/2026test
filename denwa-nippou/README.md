# 電話応対日報 集計システム（開発用の説明）

造幣局 販売・顧客サービス課の「電話応対報告書 日報集計」を、Excel マクロから
Classic ASP（VBScript）＋ Access（.accdb はデータの入れ物だけ）に置き換えるものです。

利用者向けの説明は `docs/`（納品物の `手順書\` `参考\`）にあります。ここは、作る人・直す人のための説明です。

## 納品物の作り方

```
python3 tools/check_all.py
```

上から順に次を流し、1 つでも失敗したらそこで止まります。すべて通ると
`build/package/電話応対日報_納品一式.zip` ができます。
納品した版は `dist/電話応対日報_納品一式.zip` に置いています。

| 順 | 内容 | 道具 |
|---|---|---|
| 1 | 検査用の Java 部品（Jackcess 4.0.5 / UCanAccess 5.0.1）をそろえる。照合値は固定 | tools/fetch_deps.py |
| 2 | `db/schema.sql` から Jackcess で .accdb を作り、形式バイト（0x14 が 0x02）と、表・索引・関連・初期データを、schema.sql を別の読み方で読んだものと突き合わせる。part / staff / data を組み立てる | tools/build.py, tools/java/BuildAccdb.java, tools/check_accdb.py |
| 3 | ASP の静的検査（<% %> の対応、include の実在・`..` 禁止、定義されていない名前・引数の数、配列の添字あふれ、ブロックの対応、On Error Resume Next の中の Response.End 系、SQL の文字列連結、ページ先頭の決まり）と CSS の検査（表の規則は .tw / .sheet で囲む、帳票の罫線を消さない） | tools/lint_all.py, tools/vbsim/lint.py, tools/vbsim/csscheck.py |
| 4 | 検査の道具の自己検査（わざと誤りを入れたものを必ず見つけるか） | tools/tests/test_checkers.py |
| 5 | `sql.asp` の SQL 61 本を 1 本ずつ実行し、期待値（合計 20／申込 12／うち職員 2／出勤者 3 など）と突き合わせる | tools/tests/test_sql.py |
| 6 | 画面を通した通しの検査（受付入力 → 日報確定 → 帳票・週の集計表・入力画面・入力もれチェックの数が一致、現行 Excel の 5 つの不具合が起きないこと） | tools/tests/test_flow.py |
| 7 | 困ったときの道具（setup_check・probe1～7・error.asp・致命エラー画面）を、わざと壊した状態で確かめる | tools/tests/test_troubleshoot.py |
| 8 | 帳票を printpdf.asp と同じ道筋で HTML・PDF にし、罫線が閉じている・はみ出し無し・A4 1 枚（上限いっぱいの内容でも）を確かめる | tools/render_sheet.py |
| 9 | 手順書の「　」の言葉が画面に実在するか、使わない言葉が無いか、各段に「できたことの確かめ方」があるか | tools/tests/test_docs.py |
| 10 | 全ファイルの文字コード・改行・制御文字・.accdb の形式・part/staff の分け方を検査し、通ったら UTF-8 フラグ付きの zip を作る | tools/package.py |

必要なもの: Python 3.11、Java 17 以上、`/opt/pw-browsers/chromium`（帳票の検査）、poppler（pdfinfo / pdftoppm）。

### 2 回目以降の納品（変わったファイルだけ）

```
python3 tools/package.py --since release/manifest_v1_20261006.json
```

前回の納品の一覧（`release/`）と比べ、変わったファイルだけの zip と「置き場所の表.txt」（CP932・CRLF）を作ります。
`data\*.accdb`（現地の実データ）と `include\config.asp`（現地の設定）は決して入れません。
納品したら、そのときの `build/package/manifest.json` を `release/` に日付付きで置いてください。

## フォルダ

```
db/schema.sql           .accdb の定義（DDL と初期データ）。これ 1 本から作る
src/pages/common/       part と staff の両方に置く画面（メニュー・受付入力・その他業務・エラー・設置チェック・probe）
src/pages/staff/        staff だけに置く画面（職員メニュー・日報・帳票・PDF・集計表・入力もれ・マスタ保守）
src/include/            画面の部品。part と staff に同じものを置く（config.asp の ROLE だけ置き換える）
src/css/style.css       画面（.tw）と帳票（.sheet）の見た目
src/webconfig/          web.config の元（part / staff はエラー画面の場所だけ違う）
docs/                   利用者向けの説明（0_最初にお読みください.txt はここでは UTF-8、zip では CP932・CRLF）
tools/vbsim/            VBScript / Classic ASP の模擬実行と静的検査
tools/tests/            検査
release/                納品したときのファイル一覧（差分の納品に使う）
```

include の読み込み順（下の部品は上の部品だけを使う。probe2～7 がこの順に 1 枚ずつ足していく）:
`config.asp → sql.asp → db.asp → auth.asp → layout.asp → (sheet.asp → pdf.asp)`

## 模擬実行（tools/vbsim）について

Windows・IIS・Access が無い所で品質を担保するため、VBScript と Classic ASP を Python で読み取って実行します。
納品する .asp をそのまま読み（検査側に書き写さない）、SQL は Java の中継役（tools/java/SqlBridge.java）経由で UCanAccess に流し、
`WScript.Shell.Exec` で Edge を呼ぶ所は、同じコマンド行で Chromium を動かします。

「厳格モード」で動かします。本物の VBScript が黙って行う暗黙の型変換（文字列と数値の比較、日付の文字列化、
`+` での文字列連結など）を検査の失敗として止め、ASP 側で明示的に変換させます。
ここで通るものは本物でも通る、を目指していますが、逆向き（本物でしか落ちない）は完全には保証できません。

## 確かめたこと・確かめていないこと

**確かめたこと（このリポジトリの検査で、毎回）**
- 上の表の 1～10 のすべて。

**確かめていないこと（Windows・IIS・Access・Edge の実物が無いため）**
- 本物の IIS / ASP / VBScript での動作。すべて上の模擬実行で確かめたもので、実機では一度も動かしていません。とくに次の点は実機で確かめるまで「たぶん」です。
  - web.config の `httpErrors`（500.100 → error.asp、`existingResponse="PassThrough"`）と、error.asp の中での `Response.Status = "200 OK"`、`Server.GetLastError()` の各値。error.asp の「URL」の行が、元の画面の URL になるか error.asp 自身の URL になるか。
  - web.config の `defaultDocument` の `remove` → `add`（既定の一覧に無い場合も 500.19 にならない、と見込んでいる）。
  - setup_check.asp が ServerXMLHTTP で自分自身を開く所（WinHTTP のプロキシ設定、名前の引き方）。
  - ページ先頭を `<%@ ... %><%` と改行を挟まずに書き、次の行に `Option Explicit` を置く形（改行の出力で `Option Explicit` の位置の誤りにならないようにした書き方）。
- 本物の ACE（Microsoft.ACE.OLEDB.12.0）での SQL。UCanAccess（HSQLDB）で実行し、ACE で使えない書き方は `tools/vbsim/sqlcheck.py` で止めていますが、ACE 自体では実行していません。`DateSerial(Year(..), Month(..), Day(..))`、`TOP 1 ... ORDER BY ... DESC`、`IS NOT NULL`、`= True`、ADO のパラメータの型（adVarWChar / adLongVarWChar / adBoolean / adDate）、`OpenSchema(4, Array(Empty, Empty, 表, 列))` はとくに未確認です。集計の列を ACE が何の型で返すかも未確認なので、画面側はすべて `ToLong` で変換しています。
- Jackcess で作った .accdb を、本物の Access（Microsoft 365 / バージョン 2608）で開くこと。形式バイト 0x02 と、Jackcess での読み直しまでは確かめています。
- サーバーの Edge での PDF 作成。IIS のアプリケーション プールの名前で、画面の無い所（セッション 0）から `--headless=new --no-sandbox --user-data-dir=...` で動くかは未確認です。帳票が A4 1 枚に収まることは Chromium 141 ＋ IPA ゴシックで確かめました（上限いっぱいの内容で 274.7mm／279mm）。Edge ＋ ＭＳ ゴシックでは確かめていません。
- Edge の IE モード（互換表示）での見た目。
- 手順書の Windows Server の画面の文言（サーバー マネージャー・フォルダのプロパティなど）は、実機の画面と見比べていません。
- 初期データの区分・製品・業務項目の名前と「日報の列」の割り当ては**仮のもの**です（現行 Excel を見ていないため）。手順書 04 の 1 章で、現行 Excel と合わせる手順を案内しています。

## 指示書から判断して決めたこと

- 出勤者の欄は、左の氏名列に 1～7 人目、右の氏名列に 8～14 人目を入れる（縦に読む）。備考の欄には、その行の 2 人の「姓：勤務時間 備考」を 1 人 1 行で出す。
- 「職員受電数」は、`M_担当者.[職員区分]` が "職員" の人が受けた件数。
- 帳票に出る文字は縮めずに収めるため、長さに上限を付けた（氏名 12 文字、業務の名前 16 文字、列の名前 6 文字、勤務時間＋備考 12 文字、記述欄 3 つで 15 行）。
- 指定外に足したもの: `T_業務実績` の参照整合性 2 本、`config.asp` の `PDF_WORK_DIR`（空なら data フォルダの中の pdfwork）、ACE 12.0 が無いときの 16.0 への切り替え、同じ人を 2 つの画面で同時に直したときの上書き防止、確定した日の入力の締め切り。
- 日付の切り替えは、前の日・次の日をリンク、日付を選んでの表示と保存はすべて `<form method="post">`（受け取った側が 302 で移す）。JavaScript は使っていません。
