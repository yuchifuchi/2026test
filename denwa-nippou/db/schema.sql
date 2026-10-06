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

-- ============================================================
-- 初期データ
--   M_担当者 は空で出荷する（設置後、職員がマスタ保守画面から登録する）。
--   区分・製品・業務項目の名前は仮のもの。現行 Excel と見比べて
--   マスタ保守画面で直す（手順書「04_マスタの直しかた」）。
-- ============================================================

INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (1,'申込',1);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (2,'抽選',2);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (3,'払込用紙',3);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (4,'商品発送',4);
INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (5,'その他',5);

INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (1,'商品ごとのお問合せ',TRUE,1);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (2,'手続き・案内のお問合せ',FALSE,2);
INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (3,'その他のお問合せ',FALSE,3);

INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (0,0,'（製品指定なし）',NULL,NULL,0,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (1,1,'貨幣セット',NULL,NULL,1,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (2,1,'プルーフ貨幣セット',NULL,NULL,2,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (3,1,'記念貨幣',NULL,NULL,3,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (4,1,'記念メダル',NULL,NULL,4,TRUE);
INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[適用終了日],[表示順],[有効]) VALUES (5,1,'その他の商品',NULL,NULL,5,TRUE);

INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (1,1,'申込方法・申込の受付',1,NULL,1,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (2,1,'申込内容の変更・取消',1,NULL,2,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (3,1,'抽選・当選結果',2,NULL,3,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (4,1,'払込用紙（届かない・再発行）',3,NULL,4,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (5,1,'払込方法・払込期限',3,NULL,5,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (6,1,'返金',3,'返金',6,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (7,1,'発送の時期・届かない',4,NULL,7,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (8,1,'交換（破損・不良）',4,'交換',8,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (9,1,'商品の内容',5,NULL,9,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (10,2,'販売予定・新しい商品',1,NULL,1,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (11,2,'案内・申込書の送付依頼',1,NULL,2,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (12,2,'インターネット申込の操作',1,NULL,3,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (13,2,'住所・氏名の変更',5,NULL,4,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (14,3,'見学・博物館',5,NULL,1,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (15,3,'桜の通り抜け',5,NULL,2,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (16,3,'貨幣について',5,NULL,3,TRUE,NULL);
INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (17,3,'その他',5,NULL,4,TRUE,NULL);

INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (1,'①','申込書の入力','申込書の入力',1,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (2,'②','申込内容の変更の入力','申込内容の変更入力',2,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (3,'③','払込の確認・入力','払込の確認・入力',3,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (4,'④','払込用紙の再発行','払込用紙の再発行',4,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (5,'⑤','発送データの作成','発送データの作成',5,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (6,'⑥','返品・交換の処理','返品・交換の処理',6,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (7,'⑦','返金の処理','返金の処理',7,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (8,'⑧','抽選結果の通知','抽選結果の通知',8,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (9,'⑨','郵便物の仕分け','郵便物の仕分け',9,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (10,'⑩','メール・FAX の対応','メール・FAX の対応',10,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (11,'⑪','住所変更の入力','住所変更の入力',11,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (12,'⑫','書類の整理','書類の整理',12,TRUE);
INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (13,'⑬','その他','その他',13,TRUE);
