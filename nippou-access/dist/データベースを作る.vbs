' ========================================================================
'  電話応対日報 集計システム
'  データベース (Access) を自動で作ります
'
'  このファイルをダブルクリックすると、同じフォルダに
'  日報集計_be.accdb を作ります。10 秒ほどで終わります。
'
'  Access がインストールされているパソコンで実行してください。
'  作られる accdb には VBA が入らないので、マクロの警告も出ません。
'
'  ※ このファイルは tools/gen_installer_vbs.py が src/*.bas から
'     自動生成しています。手で編集せず、生成し直してください。
' ========================================================================
Option Explicit

Dim fso, shell, here, dbPath, acc, db, i, total, done
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = fso.BuildPath(here, "日報集計_be.accdb")

' すでにある場合は作り直してよいか確認する
If fso.FileExists(dbPath) Then
  If MsgBox("すでに 日報集計_be.accdb があります。" & vbCrLf & vbCrLf & _
            "作り直すと、入力済みのデータはすべて消えます。" & vbCrLf & _
            "本当に作り直しますか？", _
            vbExclamation + vbYesNo + vbDefaultButton2, "確認") <> vbYes Then
    MsgBox "中止しました。", vbInformation, "電話応対日報"
    WScript.Quit
  End If
  On Error Resume Next
  fso.DeleteFile dbPath
  If Err.Number <> 0 Then
    MsgBox "古いファイルを消せませんでした。" & vbCrLf & vbCrLf & _
           "Access で開いたままになっていないか確認してください。", _
           vbCritical, "電話応対日報"
    WScript.Quit
  End If
  On Error GoTo 0
End If

On Error Resume Next
Set acc = CreateObject("Access.Application")
If Err.Number <> 0 Then
  MsgBox "Access が見つかりませんでした。" & vbCrLf & vbCrLf & _
         "このパソコンに Microsoft Access が入っているか確認してください。", _
         vbCritical, "電話応対日報"
  WScript.Quit
End If
On Error GoTo 0

acc.Visible = False
acc.NewCurrentDatabase dbPath
Set db = acc.CurrentDb

done = 0
total = 175

' --- 表を作る ---
Run "CREATE TABLE [M_集計列] ([集計列ID] LONG NOT NULL CONSTRAINT [PK_集計列] PRIMARY KEY,[集計列名] TEXT(50) NOT NULL,[表示順] LONG NOT NULL)"
Run "CREATE TABLE [M_ブロック] ([ブロックID] LONG NOT NULL CONSTRAINT [PK_ブロック] PRIMARY KEY,[ブロック名] TEXT(50) NOT NULL,[製品別] BIT NOT NULL,[表示順] LONG NOT NULL)"
Run "CREATE TABLE [M_製品] ([製品ID] LONG NOT NULL CONSTRAINT [PK_製品] PRIMARY KEY,[ブロックID] LONG NOT NULL,[製品名] TEXT(100) NOT NULL,[適用開始日] DATETIME,[適用終了日] DATETIME,[表示順] LONG NOT NULL,[有効] BIT NOT NULL)"
Run "CREATE TABLE [M_区分] ([区分ID] LONG NOT NULL CONSTRAINT [PK_区分] PRIMARY KEY,[ブロックID] LONG NOT NULL,[区分名] TEXT(100) NOT NULL,[集計列ID] LONG NOT NULL,[内訳区分] TEXT(20),[表示順] LONG NOT NULL,[有効] BIT NOT NULL,[旧転記名] TEXT(100),CONSTRAINT [FK_区分_ブロック] FOREIGN KEY ([ブロックID]) REFERENCES [M_ブロック]([ブロックID]),CONSTRAINT [FK_区分_集計列] FOREIGN KEY ([集計列ID]) REFERENCES [M_集計列]([集計列ID]))"
Run "CREATE TABLE [M_担当者] ([担当者ID] LONG NOT NULL CONSTRAINT [PK_担当者] PRIMARY KEY,[担当者コード] TEXT(10) NOT NULL,[姓] TEXT(50) NOT NULL,[名] TEXT(50),[氏名] TEXT(100) NOT NULL,[カナ] TEXT(100),[ログオン名] TEXT(50),[職員区分] TEXT(10) NOT NULL,[在籍開始日] DATETIME,[在籍終了日] DATETIME,[表示順] LONG NOT NULL,[有効] BIT NOT NULL,[備考] MEMO)"
Run "CREATE TABLE [M_業務項目] ([業務項目ID] LONG NOT NULL CONSTRAINT [PK_業務項目] PRIMARY KEY,[番号] TEXT(4) NOT NULL,[項目名] TEXT(100) NOT NULL,[帳票表示名] TEXT(100) NOT NULL,[表示順] LONG NOT NULL,[有効] BIT NOT NULL)"
Run "CREATE TABLE [T_日報] ([対象日] DATETIME NOT NULL CONSTRAINT [PK_日報] PRIMARY KEY,[回線数] LONG,[特記事項] MEMO,[職員代替案件] MEMO,[要望] MEMO,[状態] TEXT(10) NOT NULL,[確定日時] DATETIME,[更新日時] DATETIME)"
Run "CREATE TABLE [T_出勤] ([出勤ID] COUNTER NOT NULL CONSTRAINT [PK_出勤] PRIMARY KEY,[対象日] DATETIME NOT NULL,[担当者ID] LONG NOT NULL,[勤務時間] TEXT(50),[備考] TEXT(255),CONSTRAINT [UQ_出勤] UNIQUE ([対象日],[担当者ID]))"
Run "CREATE TABLE [T_受電] ([受電ID] COUNTER NOT NULL CONSTRAINT [PK_受電] PRIMARY KEY,[対象日] DATETIME NOT NULL,[担当者ID] LONG NOT NULL,[区分ID] LONG NOT NULL,[製品ID] LONG NOT NULL,[件数] LONG NOT NULL,[備考2] TEXT(255),[備考3] TEXT(255),[登録日時] DATETIME,[更新日時] DATETIME,[登録者] TEXT(100),CONSTRAINT [UQ_受電] UNIQUE ([対象日],[担当者ID],[区分ID],[製品ID]))"
Run "CREATE TABLE [T_業務実績] ([実績ID] COUNTER NOT NULL CONSTRAINT [PK_業務実績] PRIMARY KEY,[対象日] DATETIME NOT NULL,[担当者ID] LONG NOT NULL,[業務項目ID] LONG NOT NULL,[件数] LONG NOT NULL,CONSTRAINT [UQ_業務実績] UNIQUE ([対象日],[担当者ID],[業務項目ID]))"
Run "CREATE TABLE [T_取込ログ] ([ログID] COUNTER NOT NULL CONSTRAINT [PK_取込ログ] PRIMARY KEY,[取込日時] DATETIME,[取込元] TEXT(255),[行番号] LONG,[日付] TEXT(50),[担当者] TEXT(100),[製品名] TEXT(100),[区分名] TEXT(100),[件数] LONG,[理由] MEMO)"
Run "CREATE INDEX [IX_担当者_コード] ON [M_担当者] ([担当者コード])"
Run "CREATE INDEX [IX_担当者_ログオン名] ON [M_担当者] ([ログオン名])"
Run "CREATE INDEX [IX_受電_日付] ON [T_受電] ([対象日])"

' --- 選択肢のもとになるデータを入れる ---
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (3,'申込方法',1)"
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (4,'抽選結果',2)"
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (5,'納付書発送',3)"
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (6,'商品発送',4)"
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (7,'その他',5)"
Run "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (8,'商品交換',6)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (0,'（なし）',False,0)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (1,'区分',True,1)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (2,'その他',True,2)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (3,'顧客情報',False,3)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (4,'イベント関係',False,4)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (5,'その他のその他①',False,5)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (6,'その他のその他②',False,6)"
Run "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (7,'特殊な問合せ',False,7)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (0,0,'（製品指定なし）',0,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (1,1,'ミント',1,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (2,1,'通常プルーフ',2,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (3,1,'通年（記念日・ジャパン）',3,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (4,1,'国立公園記念貨',4,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (5,1,'ドラゴンプルーフ・貨幣セット',5,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (6,1,'鳥獣人物戯画貨幣セット',6,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (7,1,'ＩＣＤＣメダル',7,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (8,1,'桜の通り抜けプルーフ',8,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (9,1,'桜の通り抜け貨幣セット',9,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (10,1,'花のまわりみち貨幣セット',10,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (11,1,'桜の通り抜け記念メダル',11,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (12,1,'純金メダル-星座-',12,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (13,1,'国宝章牌「鳥獣人物戯画」',13,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (14,1,'干支メダル',14,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (15,1,'アジア大会記念貨',15,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (16,1,'コナンプルーフ・貨幣セット',16,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (17,1,'昭和100年記念貨幣',17,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (18,1,'七宝章牌「長浜曳山祭」',18,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (19,1,'（予備19）',19,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (20,1,'（予備20）',20,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (21,1,'（予備21）',21,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (22,1,'（予備22）',22,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (23,1,'（予備23）',23,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (24,2,'オリンピック記念貨(過去）',1,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (25,2,'皇室関係記念貨(過去)',2,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (26,2,'ミント',3,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (27,2,'通常プルーフ',4,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (28,2,'通年（記念日・ジャパン）',5,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (29,2,'世界文化遺産セット',6,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (30,2,'ＩＣＤＣメダル',7,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (31,2,'干支メダル',8,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (32,2,'純金メダル-星座-',9,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (33,2,'万博記念貨',10,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (34,2,'国立公園記念貨',11,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (35,2,'万博メダル(過去）',12,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (36,2,'世界陸上プルーフ貨幣セット',13,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (37,2,'鳥獣人物戯画ケース',14,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (38,2,'地方自治千円銀貨幣',15,True)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (39,2,'（予備16）',16,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (40,2,'（予備17）',17,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (41,2,'（予備18）',18,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (42,2,'（予備19）',19,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (43,2,'（予備20）',20,False)"
Run "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (44,2,'（予備21）',21,False)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (1,1,'申込関係',3,Null,1,True,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (2,1,'受注',3,Null,2,True,'受注')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (3,1,'抽選結果',4,Null,3,True,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (4,1,'払込用紙・再発行・可',5,Null,4,True,'払込用紙再発行（可）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (5,1,'払込用紙・再発行・不可',5,Null,5,True,'払込用紙再発行（不可）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (6,1,'払込用紙・内容等・照会',5,Null,6,True,'払込用紙内容照会')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (7,1,'払込用紙・発送状況・照会',5,Null,7,True,'払込用紙発送状況')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (8,1,'入金関係',7,Null,8,True,'入金関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (9,1,'商品発送照会・問合せ',6,Null,9,True,'商品発送問合せ')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (10,1,'商品発送照会・未着',6,Null,10,True,'商品発送受領照会')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (11,1,'製品交換',8,Null,11,True,'製品交換')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (12,2,'販売予定',7,Null,1,True,'販売予定')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (13,2,'製品内容',7,Null,2,True,'製品内容')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (14,2,'在庫照会',7,Null,3,True,'在庫照会')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (15,2,'価格照会',7,Null,4,True,'価格照会')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (16,2,'内容問合せ',7,Null,5,False,'内容問合せ')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (17,2,'組込みミス交換',7,Null,6,False,'組込みミス交換')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (18,2,'その他交換',7,Null,7,False,'その他交換')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (19,2,'（未使用8）',7,Null,8,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (20,2,'（未使用9）',7,Null,9,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (21,2,'（未使用10）',7,Null,10,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (22,2,'（未使用11）',7,Null,11,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (23,3,'ＤＭ停止等・死亡',7,Null,1,True,'ＤＭ停止(①死亡)')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (24,3,'ＤＭ停止等・病気等',7,Null,2,True,'ＤＭ停止(②病気等)')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (25,3,'ＤＭ停止等・種類多い',7,Null,3,True,'ＤＭ停止(③種類多い)')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (26,3,'ＤＭ停止等・特定商品のみ',7,Null,4,True,'ＤＭ停止(④特定商品のみ)')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (27,3,'ＤＭ停止等・理由なし',7,Null,5,True,'ＤＭ停止(⑤理由なし)')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (28,3,'ＤＭ再開・宛名人変更',7,Null,6,True,'ＤＭ再開・宛名人変更')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (29,3,'ＤＭ確認',7,Null,7,True,'ＤＭ確認')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (30,3,'住所変更',7,Null,8,True,'住所変更')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (31,3,'新規登録',7,Null,9,True,'新規登録')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (32,3,'資料送付',7,Null,10,True,'資料送付')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (33,3,'購入履歴',7,Null,11,True,'購入履歴')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (34,4,'お金と切手',7,Null,1,True,'イベント関係（お金と切手）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (35,4,'宇佐神宮',7,Null,2,True,'イベント関係（造幣局 ＩＮ）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (36,4,'桜の通り抜け',7,Null,3,True,'イベント関係（桜の通り抜け）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (37,4,'花のまわり道',7,Null,4,True,'イベント関係（まわり道）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (38,4,'ＴＩＣＣ',7,Null,5,True,'イベント関係（TICC）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (39,4,'国体',7,Null,6,True,'イベント関係（国体）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (40,4,'大阪コインショー',7,Null,7,True,'イベント関係（大阪コイン）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (41,4,'さいたまフェア',7,Null,8,True,'イベント関係（さいたまフェア）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (42,4,'佐伯区民まつり',7,Null,9,True,'イベント関係（名古屋貨幣まつり）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (43,4,'観桜会',7,Null,10,True,'イベント関係（観桜会の電話転送）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (44,4,'アンケート',7,Null,11,True,'イベント関係（予備）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (45,5,'オンライン登録関係',7,Null,1,True,'オンライン登録関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (46,5,'クレジット関係',7,Null,2,True,'クレジット関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (47,5,'金工品Ｇ関係',7,Null,3,True,'金工品Ｇへ電話転送')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (48,5,'広報（工場見学含む）',7,Null,4,True,'広報室（工場見学含む）')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (49,5,'ミントショップ案内',7,Null,5,True,'ミントショップ案内')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (50,5,'金工品懇談会',7,Null,6,True,'桜サポート')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (51,5,'流通貨幣関係',7,Null,7,True,'流通貨幣関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (52,5,'紙幣関係',7,Null,8,True,'紙幣関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (53,5,'ミントクラブ関係',7,Null,9,True,'ミントクラブ関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (54,5,'1円ぬいぐるみ',7,Null,10,True,'1円ぬいぐるみ')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (55,5,'記念貨引換関係',7,Null,11,True,'記念貨引換関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (56,6,'返 品',7,'返品',1,True,'返品')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (57,6,'返 金',7,'返金',2,True,'返金')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (58,6,'（未使用3）',7,Null,3,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (59,6,'（未使用4）',7,Null,4,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (60,6,'（未使用5）',7,Null,5,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (61,6,'貨幣ｾｯﾄ処分関係',7,Null,6,True,'貨幣ｾｯﾄ処分関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (62,6,'貨幣洗浄方法',7,Null,7,True,'貨幣洗浄方法')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (63,6,'名称未確定製品予定',7,Null,8,True,'名称未確定製品')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (64,6,'業者販売関係',7,Null,9,True,'業者販売関係')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (65,6,'抽選倍率等(過去）',7,Null,10,True,'抽選倍率等')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (66,6,'（未使用11）',7,Null,11,False,Null)"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (67,7,'別添のとおり',7,Null,1,True,'特殊な問合せ')"
Run "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (68,7,'下記のとおり（別添不要）',7,Null,2,True,'特殊な問合せ')"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (1,'①','①受注入力','①　受注入力',1,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (2,'②','②受注チェック','②　受注チェック',2,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (3,'③','③払込書チェック','③　戻り郵便処理／払込書チェック',3,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (4,'④','④戻り郵便','④　ＤＭ処理／戻り郵便',4,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (5,'⑤','⑤架電','⑤　架電',5,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (6,'⑥','⑥その他','⑥　その他（　　　　　　　　）',6,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (7,'⑦','⑦エクセル入力','⑦　エクセル入力',7,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (8,'⑧','⑧エクセルチェック','⑧　エクセルチェック',8,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (9,'⑨','⑨ｱﾝｹｰﾄ入力','⑨　アンケート入力',9,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (10,'⑩','⑩ﾊｶﾞｷﾃﾞｰﾀ入力','⑩　ﾊｶﾞｷﾃﾞｰﾀ入力',10,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (11,'⑪','⑪ﾊｶﾞｷﾃﾞｰﾀﾁｪｯｸ','⑪　ハガキデータチェック',11,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (12,'⑫','⑫新規ｺｰﾄﾞ取り','⑫　新規ｺｰﾄﾞ取り',12,True)"
Run "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (13,'⑬','⑬顧客整理','⑬　顧客整理',13,True)"

' --- 集計のしかたを登録する ---
MakeQuery "Q_選択_担当者", "SELECT OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[職員区分], OP.[表示順] FROM [M_担当者] AS OP WHERE OP.[有効]=True   AND (OP.[在籍開始日] Is Null Or OP.[在籍開始日] <= Date())   AND (OP.[在籍終了日] Is Null Or OP.[在籍終了日] >= Date()) ORDER BY OP.[表示順];"
MakeQuery "Q_選択_担当者_基準日", "PARAMETERS [基準日] DateTime; SELECT OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[職員区分], OP.[表示順] FROM [M_担当者] AS OP WHERE OP.[有効]=True   AND (OP.[在籍開始日] Is Null Or OP.[在籍開始日] <= [基準日])   AND (OP.[在籍終了日] Is Null Or OP.[在籍終了日] >= [基準日]) ORDER BY OP.[表示順];"
MakeQuery "Q_選択_区分", "SELECT KB.[区分ID], KB.[ブロックID], BK.[ブロック名], KB.[区分名],        BK.[ブロック名] & ' / ' & KB.[区分名] AS [表示名],        KB.[集計列ID], SC.[集計列名], BK.[製品別],        BK.[表示順]*1000 + KB.[表示順] AS [並び順] FROM ([M_区分] AS KB INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID])      INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID] WHERE KB.[有効]=True ORDER BY BK.[表示順]*1000 + KB.[表示順];"
MakeQuery "Q_選択_製品", "SELECT PR.[製品ID], PR.[ブロックID], BK.[ブロック名], PR.[製品名],        IIf(PR.[製品ID]=0, PR.[製品名], BK.[ブロック名] & ' / ' & PR.[製品名]) AS [表示名],        PR.[表示順] FROM [M_製品] AS PR INNER JOIN [M_ブロック] AS BK ON PR.[ブロックID]=BK.[ブロックID] WHERE PR.[有効]=True   AND (PR.[適用開始日] Is Null Or PR.[適用開始日] <= Date())   AND (PR.[適用終了日] Is Null Or PR.[適用終了日] >= Date()) ORDER BY PR.[ブロックID], PR.[表示順];"
MakeQuery "Q_受電明細", "SELECT J.[受電ID], J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名],        OP.[姓] AS [担当者], OP.[職員区分], BK.[ブロック名],        PR.[製品名], KB.[区分ID], KB.[区分名], KB.[集計列ID], SC.[集計列名],        IIf(KB.[集計列ID]=3,J.[件数],0) AS [申込方法],        IIf(KB.[集計列ID]=4,J.[件数],0) AS [抽選結果],        IIf(KB.[集計列ID]=5,J.[件数],0) AS [納付書発送],        IIf(KB.[集計列ID]=6,J.[件数],0) AS [商品発送],        IIf(KB.[集計列ID]=7,J.[件数],0) AS [その他],        " & _
        "IIf(KB.[集計列ID]=8,J.[件数],0) AS [商品交換],        J.[件数] AS [計], J.[備考2], J.[備考3] FROM ((((([T_受電] AS J   INNER JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID])   INNER JOIN [M_区分]   AS KB ON J.[区分ID]=KB.[区分ID])   INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID])   INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID])   INNER JOIN [M_製品]   AS PR ON J.[製品ID]=PR.[製品ID]) WHERE J.[件数]<>0;"
MakeQuery "Q_日別集計", "TRANSFORM Sum(J.[件数]) AS [件数計] SELECT J.[対象日], Sum(J.[件数]) AS [計] FROM ([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID])      INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID] GROUP BY J.[対象日] PIVOT SC.[集計列名] IN ('申込方法','抽選結果','納付書発送','商品発送','その他','商品交換');"
MakeQuery "Q_日報_受電", "SELECT J.[対象日],  Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込],  Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選],  Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [払込用紙],  Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送],  Sum(IIf(KB.[集計列ID] In (7,8),J.[件数],0)) AS [その他],  Sum(J.[件数]) AS [合計],  Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [内交換],  Sum(IIf(KB.[内訳区分]='返金',J.[件数],0)) AS [内返金],  Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=3,J.[件数],0)) " & _
        "AS [申込_職員],  Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=4,J.[件数],0)) AS [抽選_職員],  Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=5,J.[件数],0)) AS [払込用紙_職員],  Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=6,J.[件数],0)) AS [商品発送_職員],  Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID] In (7,8),J.[件数],0)) AS [その他_職員],  Sum(IIf(OP.[職員区分]='職員',J.[件数],0)) AS [合計_職員] FROM ([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID])      INNE" & _
        "R JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID] GROUP BY J.[対象日];"
MakeQuery "Q_日報_業務", "SELECT W.[対象日], TM.[業務項目ID], TM.[番号], TM.[項目名], TM.[帳票表示名],        TM.[表示順], Sum(W.[件数]) AS [件数] FROM [T_業務実績] AS W INNER JOIN [M_業務項目] AS TM      ON W.[業務項目ID]=TM.[業務項目ID] GROUP BY W.[対象日], TM.[業務項目ID], TM.[番号], TM.[項目名], TM.[帳票表示名], TM.[表示順];"
MakeQuery "Q_日報_出勤", "SELECT A.[対象日], OP.[担当者ID], OP.[氏名], OP.[職員区分],        A.[勤務時間], A.[備考], OP.[表示順] FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] ORDER BY A.[対象日], OP.[表示順];"
MakeQuery "Q_日報_ヘッダ", "SELECT H.[対象日], H.[回線数], H.[特記事項], H.[職員代替案件], H.[要望],        H.[状態], H.[確定日時],        (SELECT Count(*) FROM [T_出勤] AS A WHERE A.[対象日]=H.[対象日]) AS [出勤者数],        Nz(R.[申込],0) AS [申込], Nz(R.[抽選],0) AS [抽選],        Nz(R.[払込用紙],0) AS [払込用紙], Nz(R.[商品発送],0) AS [商品発送],        Nz(R.[その他],0) AS [その他], Nz(R.[合計],0) AS [合計],        Nz(R.[内交換],0) AS [内交換], Nz(R.[内返金],0) AS [内返金],        Nz(R.[申込_職員],0) AS [" & _
        "申込_職員], Nz(R.[抽選_職員],0) AS [抽選_職員],        Nz(R.[払込用紙_職員],0) AS [払込用紙_職員],        Nz(R.[商品発送_職員],0) AS [商品発送_職員],        Nz(R.[その他_職員],0) AS [その他_職員], Nz(R.[合計_職員],0) AS [合計_職員] FROM [T_日報] AS H LEFT JOIN [Q_日報_受電] AS R ON H.[対象日]=R.[対象日];"
MakeQuery "Q_担当者別日次", "SELECT J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名],        Sum(J.[件数]) AS [受電件数] FROM [T_受電] AS J INNER JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID] GROUP BY J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名];"
MakeQuery "Q_未入力チェック", "SELECT A.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[表示順] FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] WHERE NOT EXISTS (SELECT 1 FROM [T_受電] AS J                   WHERE J.[対象日]=A.[対象日] AND J.[担当者ID]=A.[担当者ID] AND J.[件数]<>0)   AND NOT EXISTS (SELECT 1 FROM [T_業務実績] AS W                   WHERE W.[対象日]=A.[対象日] AND W.[担当者ID]=A.[担当者ID] AND W.[件数]<>0) ORDER BY A.[対象日], OP.[表示順];"
MakeQuery "Q_週次明細", "PARAMETERS [開始日] DateTime, [終了日] DateTime; SELECT * FROM [Q_受電明細] WHERE [対象日] Between [開始日] And [終了日] ORDER BY [対象日], [氏名], [区分ID];"
MakeQuery "Q_週次集計", "PARAMETERS [開始日] DateTime, [終了日] DateTime; SELECT J.[対象日],  Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法],  Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選結果],  Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [納付書発送],  Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送],  Sum(IIf(KB.[集計列ID]=7,J.[件数],0)) AS [その他],  Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [商品交換],  Sum(J.[件数]) AS [計] FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[" & _
        "区分ID] WHERE J.[対象日] Between [開始日] And [終了日] GROUP BY J.[対象日] ORDER BY J.[対象日];"
MakeQuery "Q_週次集計指定", "SELECT J.[対象日],  Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法],  Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選結果],  Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [納付書発送],  Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送],  Sum(IIf(KB.[集計列ID]=7,J.[件数],0)) AS [その他],  Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [商品交換],  Sum(J.[件数]) AS [計] FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID] GROUP BY J.[対象日];"

' --- 表どうしのつながりを登録する ---
'     ここが効いていると、実績から使われているマスタは消せなくなる
AddRel "R_受電_担当者", "M_担当者", "T_受電", "担当者ID"
AddRel "R_受電_区分", "M_区分", "T_受電", "区分ID"
AddRel "R_受電_製品", "M_製品", "T_受電", "製品ID"
AddRel "R_出勤_担当者", "M_担当者", "T_出勤", "担当者ID"
AddRel "R_業務実績_担当者", "M_担当者", "T_業務実績", "担当者ID"
AddRel "R_業務実績_項目", "M_業務項目", "T_業務実績", "業務項目ID"

acc.CloseCurrentDatabase
acc.Quit
Set db = Nothing
Set acc = Nothing

MsgBox "データベースを作りました。" & vbCrLf & vbCrLf & _
       dbPath & vbCrLf & vbCrLf & _
       "このファイルを共有フォルダに置いてください。" & vbCrLf & _
       "置いた場所は、あとで Web サーバーの設定に書きます。", _
       vbInformation, "電話応対日報"

' ------------------------------------------------------------------------
Sub Run(sql)
  On Error Resume Next
  db.Execute sql, 128        ' 128 = dbFailOnError
  If Err.Number <> 0 Then Fail sql, Err.Description
  On Error GoTo 0
  done = done + 1
End Sub

Sub MakeQuery(nm, sql)
  On Error Resume Next
  db.CreateQueryDef nm, sql
  If Err.Number <> 0 Then Fail nm, Err.Description
  On Error GoTo 0
  done = done + 1
End Sub

Sub AddRel(nm, parentTable, childTable, fld)
  Dim rel
  On Error Resume Next
  Set rel = db.CreateRelation(nm, parentTable, childTable, 0)   ' 0 = 整合性を守る
  rel.Fields.Append rel.CreateField(fld)
  rel.Fields(fld).ForeignName = fld
  db.Relations.Append rel
  If Err.Number <> 0 Then Err.Clear    ' つながりは作れなくても致命傷ではない
  On Error GoTo 0
  done = done + 1
End Sub

Sub Fail(what, why)
  On Error Resume Next
  acc.CloseCurrentDatabase
  acc.Quit
  On Error GoTo 0
  MsgBox "作成に失敗しました。" & vbCrLf & vbCrLf & _
         "処理: " & Left(what, 120) & vbCrLf & _
         "理由: " & why & vbCrLf & vbCrLf & _
         "この画面を写真に撮って、担当者にお知らせください。", _
         vbCritical, "電話応対日報"
  WScript.Quit
End Sub
