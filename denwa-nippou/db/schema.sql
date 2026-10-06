-- ============================================================
-- 電話応対日報 集計システム  データベース定義（DDL と初期データ）
--
-- このファイル 1 本から 日報集計_be.accdb を作ります。
--   作り方: python3 tools/build.py accdb
--   （tools/java/BuildAccdb.java がこのファイルを読み、Jackcess で書き込みます）
--
-- 書式の決まり（読み取り側が簡単な作りなので、これ以外は書かない）:
--   ・文の終わりは ; ／ 注釈は -- から行末まで
--   ・名前は必ず [ ] で囲む
--   ・型は LONG / COUNTER / TEXT(n) / MEMO / BIT / DATETIME のみ
-- ============================================================

-- ---------- マスタ ----------

CREATE TABLE [M_集計列] (
  [集計列ID] LONG NOT NULL CONSTRAINT [PK_集計列] PRIMARY KEY,
  [集計列名] TEXT(50) NOT NULL,
  [表示順] LONG NOT NULL);
-- 日報の「申込／抽選／払込用紙／商品発送／その他」の 5 列

CREATE TABLE [M_ブロック] (
  [ブロックID] LONG NOT NULL CONSTRAINT [PK_ブロック] PRIMARY KEY,
  [ブロック名] TEXT(50) NOT NULL,
  [製品別] BIT NOT NULL,
  [表示順] LONG NOT NULL);
-- 入力画面の大きなまとまり。[製品別]=True なら製品を選ばせる

CREATE TABLE [M_製品] (
  [製品ID] LONG NOT NULL CONSTRAINT [PK_製品] PRIMARY KEY,
  [ブロックID] LONG NOT NULL,
  [製品名] TEXT(100) NOT NULL,
  [適用開始日] DATETIME,
  [適用終了日] DATETIME,
  [表示順] LONG NOT NULL,
  [有効] BIT NOT NULL);
-- 製品ID = 0 は「製品指定なし」として必ず 1 行入れておく。
-- [ブロックID] に参照整合性を張らないのは、製品ID = 0 の行がどのブロックにも属さない（0 を入れる）ため。

CREATE TABLE [M_区分] (
  [区分ID] LONG NOT NULL CONSTRAINT [PK_区分] PRIMARY KEY,
  [ブロックID] LONG NOT NULL,
  [区分名] TEXT(100) NOT NULL,
  [集計列ID] LONG NOT NULL,
  [内訳区分] TEXT(20),
  [表示順] LONG NOT NULL,
  [有効] BIT NOT NULL,
  [旧転記名] TEXT(100),
  CONSTRAINT [FK_区分_ブロック] FOREIGN KEY ([ブロックID]) REFERENCES [M_ブロック]([ブロックID]),
  CONSTRAINT [FK_区分_集計列] FOREIGN KEY ([集計列ID]) REFERENCES [M_集計列]([集計列ID]));
-- [集計列ID] が「この区分は日報のどの列に積むか」。
--   現行 Excel ではシートの 2 行目に埋め込まれていて、誰にも見えなかった情報。
-- [内訳区分] は "交換" / "返金"。日報の「内 交換 ○件／内 返金 ○件」に使う。
-- [旧転記名] は現行 Excel の表記。過去データ取り込みの突き合わせ用。

CREATE TABLE [M_担当者] (
  [担当者ID] LONG NOT NULL CONSTRAINT [PK_担当者] PRIMARY KEY,
  [担当者コード] TEXT(10) NOT NULL,
  [姓] TEXT(50) NOT NULL,
  [名] TEXT(50),
  [氏名] TEXT(100) NOT NULL,
  [カナ] TEXT(100),
  [ログオン名] TEXT(50),
  [職員区分] TEXT(10) NOT NULL,
  [在籍開始日] DATETIME,
  [在籍終了日] DATETIME,
  [表示順] LONG NOT NULL,
  [有効] BIT NOT NULL,
  [備考] MEMO);
CREATE INDEX [IX_担当者_コード] ON [M_担当者] ([担当者コード]);
CREATE INDEX [IX_担当者_ログオン名] ON [M_担当者] ([ログオン名]);
-- [職員区分] は "パート" / "職員"。
-- 担当者は「削除」しない。在籍終了日を入れると入力画面の候補から消えるが、
-- 過去の日報・集計表はその担当者のまま残る。
-- 担当者コードは席番号で、人が替わると再利用される。だから UNIQUE は張らない。

CREATE TABLE [M_業務項目] (
  [業務項目ID] LONG NOT NULL CONSTRAINT [PK_業務項目] PRIMARY KEY,
  [番号] TEXT(4) NOT NULL,
  [項目名] TEXT(100) NOT NULL,
  [帳票表示名] TEXT(100) NOT NULL,
  [表示順] LONG NOT NULL,
  [有効] BIT NOT NULL);
-- 日報の下段「電話対応以外の業務（データ入力作業等）」の ①～⑬

CREATE TABLE [M_回覧] (
  [回覧ID] LONG NOT NULL CONSTRAINT [PK_回覧] PRIMARY KEY,
  [表示名] TEXT(10),
  [表示順] LONG NOT NULL);
-- 日報の上の「回覧」の押印欄（指定外に足した表。課の帳票の様式に合わせるため）。
-- 表示順 1～4 が左の 4 つ、5～8 が右の 4 つ。人が替わったらマスタ保守で名前を直す。表示名が空の欄は枠だけ出す。

-- ---------- 実績 ----------

CREATE TABLE [T_日報] (
  [対象日] DATETIME NOT NULL CONSTRAINT [PK_日報] PRIMARY KEY,
  [回線数] LONG,
  [特記事項] MEMO,
  [職員代替案件] MEMO,
  [要望] MEMO,
  [状態] TEXT(10) NOT NULL,
  [確定日時] DATETIME,
  [更新日時] DATETIME);
-- [状態] は "作成中" / "確定"

CREATE TABLE [T_出勤] (
  [出勤ID] COUNTER NOT NULL CONSTRAINT [PK_出勤] PRIMARY KEY,
  [対象日] DATETIME NOT NULL,
  [担当者ID] LONG NOT NULL,
  [勤務時間] TEXT(50),
  [備考] TEXT(255),
  CONSTRAINT [UQ_出勤] UNIQUE ([対象日],[担当者ID]));

CREATE TABLE [T_受電] (
  [受電ID] COUNTER NOT NULL CONSTRAINT [PK_受電] PRIMARY KEY,
  [対象日] DATETIME NOT NULL,
  [担当者ID] LONG NOT NULL,
  [区分ID] LONG NOT NULL,
  [製品ID] LONG NOT NULL,
  [件数] LONG NOT NULL,
  [備考2] TEXT(255),
  [備考3] TEXT(255),
  [登録日時] DATETIME,
  [更新日時] DATETIME,
  [登録者] TEXT(100),
  CONSTRAINT [UQ_受電] UNIQUE ([対象日],[担当者ID],[区分ID],[製品ID]));
CREATE INDEX [IX_受電_日付] ON [T_受電] ([対象日]);
-- 1 行 = 1 (日付, 担当者, 区分, 製品) の件数。
-- 製品を伴わない区分は 製品ID = 0 を入れる。NULL を使わないことで
-- UNIQUE 制約が確実に効き、二重計上が構造的に起きない。

CREATE TABLE [T_業務実績] (
  [実績ID] COUNTER NOT NULL CONSTRAINT [PK_業務実績] PRIMARY KEY,
  [対象日] DATETIME NOT NULL,
  [担当者ID] LONG NOT NULL,
  [業務項目ID] LONG NOT NULL,
  [件数] LONG NOT NULL,
  CONSTRAINT [UQ_業務実績] UNIQUE ([対象日],[担当者ID],[業務項目ID]));

CREATE TABLE [T_受付メモ] (
  [メモID] COUNTER NOT NULL CONSTRAINT [PK_受付メモ] PRIMARY KEY,
  [対象日] DATETIME NOT NULL,
  [担当者ID] LONG NOT NULL,
  [内容] MEMO,
  [更新日時] DATETIME,
  CONSTRAINT [UQ_受付メモ] UNIQUE ([対象日],[担当者ID]));
-- 指定外に足した表。現行 Excel の記入用フォームで、特殊な問合せの「下記のとおり（別添不要）」の内容を
-- 書いていた欄（A64）にあたる。1 日・1 人に 1 行。職員は日報の画面で全員ぶんを読んで、特記事項を書く。

-- ---------- 参照整合性 ----------
-- 実績が「どこにも属さない行」にならないように張る。
-- （担当者・区分・製品を消そうとすると、使われている限り Access が止める）
ALTER TABLE [T_受電] ADD CONSTRAINT [FK_受電_担当者] FOREIGN KEY ([担当者ID]) REFERENCES [M_担当者]([担当者ID]);
ALTER TABLE [T_受電] ADD CONSTRAINT [FK_受電_区分] FOREIGN KEY ([区分ID]) REFERENCES [M_区分]([区分ID]);
ALTER TABLE [T_受電] ADD CONSTRAINT [FK_受電_製品] FOREIGN KEY ([製品ID]) REFERENCES [M_製品]([製品ID]);
ALTER TABLE [T_出勤] ADD CONSTRAINT [FK_出勤_担当者] FOREIGN KEY ([担当者ID]) REFERENCES [M_担当者]([担当者ID]);
-- 次の 2 本は指定外だが、同じ理由で張る（その他業務の件数が持ち主不明にならないように）
ALTER TABLE [T_業務実績] ADD CONSTRAINT [FK_業務実績_担当者] FOREIGN KEY ([担当者ID]) REFERENCES [M_担当者]([担当者ID]);
ALTER TABLE [T_業務実績] ADD CONSTRAINT [FK_業務実績_業務項目] FOREIGN KEY ([業務項目ID]) REFERENCES [M_業務項目]([業務項目ID]);
ALTER TABLE [T_受付メモ] ADD CONSTRAINT [FK_受付メモ_担当者] FOREIGN KEY ([担当者ID]) REFERENCES [M_担当者]([担当者ID]);

-- ============================================================
-- 初期データ
--   M_担当者 は空で出荷する（設置後、職員がマスタ保守画面から登録する）。
--   業務項目（①～⑬）と回覧の欄は、課の帳票の様式（印刷用シート）から写したもの。
--   ブロック・製品・区分は、現行 Excel の「記入用フォーム」の行と列から写したもの（2026 年 8 月時点の製品）。
--     区分の [集計列ID] は、記入用フォームの転記用シートの 2 行目に隠れていた番号（3 申込・4 抽選・5 払込用紙・
--     6 商品発送・7 その他・8 製品交換）と、日報集計印刷用フォームの式（製品交換はその他に足し、内 交換に数える）から決めた。
--     [旧転記名] は転記用シートの 3 行目の見出し（入力用シートの見出しと食い違っていたものも、そのまま残す）。
--   区分名の「／」は、入力画面で見出しを 2 段にする区切り（例「払込用紙／再発行（可）」→ 上の段「払込用紙」）。
-- ============================================================

INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (1,'申込',1);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (2,'抽選',2);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (3,'払込用紙',3);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (4,'商品発送',4);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (5,'その他',5);

INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (1,'製品ごとのお問合せ',TRUE,1);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (2,'【その他】製品のお問合せ（販売予定・過去の製品など）',TRUE,2);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (3,'【顧客情報】',FALSE,3);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (4,'【イベント関係】',FALSE,4);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (5,'【その他のその他①】',FALSE,5);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (6,'【その他のその他②】',FALSE,6);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (7,'【特殊な問合せ】',FALSE,7);

INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (0,0,'（製品指定なし）',NULL,NULL,0,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (1,1,'ミント',NULL,NULL,10,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (2,1,'通常プルーフ',NULL,NULL,20,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (3,1,'通年（記念日・ジャパン）',NULL,NULL,30,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (4,1,'国立公園記念貨',NULL,NULL,40,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (5,1,'ドラゴンプルーフ・貨幣セット',NULL,NULL,50,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (6,1,'鳥獣人物戯画貨幣セット',NULL,NULL,60,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (7,1,'ＩＣＤＣメダル',NULL,NULL,70,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (8,1,'桜の通り抜けプルーフ',NULL,NULL,80,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (9,1,'桜の通り抜け貨幣セット',NULL,NULL,90,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (10,1,'花のまわりみち貨幣セット',NULL,NULL,100,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (11,1,'桜の通り抜け記念メダル',NULL,NULL,110,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (12,1,'純金メダル-星座-',NULL,NULL,120,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (13,1,'国宝章牌「鳥獣人物戯画」',NULL,NULL,130,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (14,1,'干支メダル',NULL,NULL,140,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (15,1,'アジア大会記念貨',NULL,NULL,150,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (16,1,'コナンプルーフ・貨幣セット',NULL,NULL,160,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (17,1,'昭和100年記念貨幣',NULL,NULL,170,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (18,1,'七宝章牌「長浜曳山祭」',NULL,NULL,180,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (19,2,'オリンピック記念貨（過去）',NULL,NULL,10,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (20,2,'皇室関係記念貨（過去）',NULL,NULL,20,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (21,2,'ミント',NULL,NULL,30,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (22,2,'通常プルーフ',NULL,NULL,40,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (23,2,'通年（記念日・ジャパン）',NULL,NULL,50,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (24,2,'世界文化遺産セット',NULL,NULL,60,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (25,2,'ＩＣＤＣメダル',NULL,NULL,70,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (26,2,'干支メダル',NULL,NULL,80,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (27,2,'純金メダル-星座-',NULL,NULL,90,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (28,2,'万博記念貨',NULL,NULL,100,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (29,2,'国立公園記念貨',NULL,NULL,110,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (30,2,'万博メダル（過去）',NULL,NULL,120,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (31,2,'世界陸上プルーフ貨幣セット',NULL,NULL,130,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (32,2,'鳥獣人物戯画ケース',NULL,NULL,140,TRUE);

INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (1,1,'申込関係',1,NULL,10,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (2,1,'受注',1,NULL,20,TRUE,'受注');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (3,1,'抽選結果',2,NULL,30,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (4,1,'払込用紙／再発行（可）',3,NULL,40,TRUE,'払込用紙再発行（可）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (5,1,'払込用紙／再発行（不可）',3,NULL,50,TRUE,'払込用紙再発行（不可）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (6,1,'払込用紙／内容等照会',3,NULL,60,TRUE,'払込用紙内容照会');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (7,1,'払込用紙／発送状況',3,NULL,70,TRUE,'払込用紙発送状況');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (8,1,'入金関係',5,NULL,80,TRUE,'入金関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (9,1,'商品発送照会／問合せ',4,NULL,90,TRUE,'商品発送問合せ');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (10,1,'商品発送照会／未着',4,NULL,100,TRUE,'商品発送受領照会');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (11,1,'製品交換',5,'交換',110,TRUE,'製品交換');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (12,2,'販売予定',5,NULL,10,TRUE,'販売予定');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (13,2,'製品内容',5,NULL,20,TRUE,'製品内容');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (14,2,'在庫照会',5,NULL,30,TRUE,'在庫照会');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (15,2,'価格照会',5,NULL,40,TRUE,'価格照会');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (16,3,'ＤＭ停止等／死亡',5,NULL,10,TRUE,'ＤＭ停止(①死亡)');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (17,3,'ＤＭ停止等／病気等',5,NULL,20,TRUE,'ＤＭ停止(②病気等)');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (18,3,'ＤＭ停止等／種類多い',5,NULL,30,TRUE,'ＤＭ停止(③種類多い)');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (19,3,'ＤＭ停止等／特定商品のみ',5,NULL,40,TRUE,'ＤＭ停止(④特定商品のみ)');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (20,3,'ＤＭ停止等／理由なし',5,NULL,50,TRUE,'ＤＭ停止(⑤理由なし)');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (21,3,'ＤＭ再開・宛名人変更',5,NULL,60,TRUE,'ＤＭ再開・宛名人変更');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (22,3,'ＤＭ確認',5,NULL,70,TRUE,'ＤＭ確認');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (23,3,'住所変更',5,NULL,80,TRUE,'住所変更');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (24,3,'新規登録',5,NULL,90,TRUE,'新規登録');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (25,3,'資料送付',5,NULL,100,TRUE,'資料送付');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (26,3,'購入履歴',5,NULL,110,TRUE,'購入履歴');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (27,4,'お金と切手',5,NULL,10,TRUE,'イベント関係（お金と切手）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (28,4,'宇佐神宮',5,NULL,20,TRUE,'イベント関係（造幣局　ＩＮ）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (29,4,'桜の通り抜け',5,NULL,30,TRUE,'イベント関係（桜の通り抜け）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (30,4,'花のまわり道',5,NULL,40,TRUE,'イベント関係（まわり道）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (31,4,'ＴＩＣＣ',5,NULL,50,TRUE,'イベント関係（TICC）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (32,4,'国体',5,NULL,60,TRUE,'イベント関係（国体）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (33,4,'大阪コインショー',5,NULL,70,TRUE,'イベント関係（大阪コイン）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (34,4,'さいたまフェア',5,NULL,80,TRUE,'イベント関係（さいたまフェア）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (35,4,'佐伯区民まつり',5,NULL,90,TRUE,'イベント関係（名古屋貨幣まつり）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (36,4,'観桜会',5,NULL,100,TRUE,'イベント関係（観桜会の電話転送）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (37,4,'アンケート',5,NULL,110,TRUE,'イベント関係（予備）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (38,5,'オンライン登録関係',5,NULL,10,TRUE,'オンライン登録関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (39,5,'クレジット関係',5,NULL,20,TRUE,'クレジット関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (40,5,'金工品Ｇ関係',5,NULL,30,TRUE,'金工品Ｇへ電話転送');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (41,5,'広報（工場見学含む）',5,NULL,40,TRUE,'広報室（工場見学含む）');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (42,5,'ミントショップ案内',5,NULL,50,TRUE,'ミントショップ案内');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (43,5,'金工品懇談会',5,NULL,60,TRUE,'桜サポート');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (44,5,'流通貨幣関係',5,NULL,70,TRUE,'流通貨幣関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (45,5,'紙幣関係',5,NULL,80,TRUE,'紙幣関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (46,5,'ミントクラブ関係',5,NULL,90,TRUE,'ミントクラブ関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (47,5,'1円ぬいぐるみ',5,NULL,100,TRUE,'1円ぬいぐるみ');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (48,5,'記念貨引換関係',5,NULL,110,TRUE,'記念貨引換関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (49,6,'返品',5,NULL,10,TRUE,'返品');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (50,6,'返金',5,'返金',20,TRUE,'返金');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (51,6,'貨幣セット処分関係',5,NULL,30,TRUE,'貨幣ｾｯﾄ処分関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (52,6,'貨幣洗浄方法',5,NULL,40,TRUE,'貨幣洗浄方法');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (53,6,'名称未確定製品予定',5,NULL,50,TRUE,'名称未確定製品');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (54,6,'業者販売関係',5,NULL,60,TRUE,'業者販売関係');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (55,6,'抽選倍率等（過去）',5,NULL,70,TRUE,'抽選倍率等');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (56,7,'別添のとおり',5,NULL,10,TRUE,'特殊な問合せ');
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (57,7,'下記のとおり（別添不要）',5,NULL,20,TRUE,'特殊な問合せ');


INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (1,'①','受注入力','受注入力',1,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (2,'②','受注チェック','受注チェック',2,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (3,'③','払込書チェック','払込書チェック',3,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (4,'④','戻り郵便','戻り郵便',4,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (5,'⑤','架電','架電',5,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (6,'⑥','その他','その他（　　　　）',6,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (7,'⑦','エクセル入力','エクセル入力',7,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (8,'⑧','エクセルチェック','エクセルチェック',8,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (9,'⑨','アンケート入力','アンケート入力',9,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (10,'⑩','ハガキデータ入力','ハガキデータ入力',10,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (11,'⑪','ハガキデータチェック','ハガキデータチェック',11,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (12,'⑫','新規コード取り','新規コード取り',12,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (13,'⑬','顧客整理','顧客整理',13,TRUE);

INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (1,'補佐',1);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (2,'奥田',2);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (3,'藤田',3);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (4,'堀',4);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (5,'藤本課長',5);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (6,'岡田補佐',6);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (7,'出口補佐',7);
INSERT INTO [M_回覧] ([回覧ID],[表示名],[表示順]) VALUES (8,'佐藤補佐',8);
