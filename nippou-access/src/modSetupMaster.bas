Attribute VB_Name = "modSetupMaster"
Option Compare Database
Option Explicit

'==============================================================================
' マスタ初期データ
'
' このモジュールは tools/gen_master_vba.py が data/*.csv から自動生成する。
' 手で編集せず、CSV を直してから再生成すること。
'
' 出典: 現行の記入用フォーム (Sheet1 の見出しは日報集計管理用フォームへの外部リンク)
'       および Sheet2 の 2 行目 (集計表の列番号) / 3 行目 (旧転記名)。
'==============================================================================

Public Sub Setup_Master()
    Echo_ "マスタを登録しています..."

    ' マスタは毎回作り直すので、まず全消しする
    ExecSQL "DELETE FROM [M_区分]"
    ExecSQL "DELETE FROM [M_製品]"
    ExecSQL "DELETE FROM [M_ブロック]"
    ExecSQL "DELETE FROM [M_集計列]"
    ExecSQL "DELETE FROM [M_業務項目]"
    ExecSQL "DELETE FROM [M_担当者]"

    Master_集計列
    Master_ブロック
    Master_製品
    Master_区分_1
    Master_区分_2
    Master_担当者
    Master_業務項目

    Echo_ "  マスタ登録 完了"
End Sub

Private Sub Master_集計列()
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (3,'申込方法',1)"
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (4,'抽選結果',2)"
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (5,'納付書発送',3)"
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (6,'商品発送',4)"
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (7,'その他',5)"
    ExecSQL "INSERT INTO [M_集計列] ([集計列ID],[集計列名],[表示順]) VALUES (8,'商品交換',6)"
End Sub

Private Sub Master_ブロック()
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (0,'（なし）',False,0)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (1,'区分',True,1)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (2,'その他',True,2)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (3,'顧客情報',False,3)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (4,'イベント関係',False,4)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (5,'その他のその他①',False,5)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (6,'その他のその他②',False,6)"
    ExecSQL "INSERT INTO [M_ブロック] ([ブロックID],[ブロック名],[製品別],[表示順]) VALUES (7,'特殊な問合せ',False,7)"
End Sub

Private Sub Master_製品()
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (0,0,'（製品指定なし）',0,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (1,1,'ミント',1,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (2,1,'通常プルーフ',2,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (3,1,'通年（記念日・ジャパン）',3,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (4,1,'国立公園記念貨',4,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (5,1,'ドラゴンプルーフ・貨幣セット',5,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (6,1,'鳥獣人物戯画貨幣セット',6,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (7,1,'ＩＣＤＣメダル',7,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (8,1,'桜の通り抜けプルーフ',8,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (9,1,'桜の通り抜け貨幣セット',9,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (10,1,'花のまわりみち貨幣セット',10,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (11,1,'桜の通り抜け記念メダル',11,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (12,1,'純金メダル-星座-',12,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (13,1,'国宝章牌「鳥獣人物戯画」',13,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (14,1,'干支メダル',14,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (15,1,'アジア大会記念貨',15,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (16,1,'コナンプルーフ・貨幣セット',16,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (17,1,'昭和100年記念貨幣',17,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (18,1,'七宝章牌「長浜曳山祭」',18,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (19,1,'（予備19）',19,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (20,1,'（予備20）',20,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (21,1,'（予備21）',21,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (22,1,'（予備22）',22,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (23,1,'（予備23）',23,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (24,2,'オリンピック記念貨(過去）',1,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (25,2,'皇室関係記念貨(過去)',2,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (26,2,'ミント',3,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (27,2,'通常プルーフ',4,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (28,2,'通年（記念日・ジャパン）',5,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (29,2,'世界文化遺産セット',6,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (30,2,'ＩＣＤＣメダル',7,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (31,2,'干支メダル',8,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (32,2,'純金メダル-星座-',9,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (33,2,'万博記念貨',10,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (34,2,'国立公園記念貨',11,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (35,2,'万博メダル(過去）',12,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (36,2,'世界陸上プルーフ貨幣セット',13,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (37,2,'鳥獣人物戯画ケース',14,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (38,2,'地方自治千円銀貨幣',15,True)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (39,2,'（予備16）',16,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (40,2,'（予備17）',17,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (41,2,'（予備18）',18,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (42,2,'（予備19）',19,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (43,2,'（予備20）',20,False)"
    ExecSQL "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[表示順],[有効]) VALUES (44,2,'（予備21）',21,False)"
End Sub

Private Sub Master_区分_1()
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (1,1,'申込関係',3,Null,1,True,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (2,1,'受注',3,Null,2,True,'受注')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (3,1,'抽選結果',4,Null,3,True,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (4,1,'払込用紙・再発行・可',5,Null,4,True,'払込用紙再発行（可）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (5,1,'払込用紙・再発行・不可',5,Null,5,True,'払込用紙再発行（不可）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (6,1,'払込用紙・内容等・照会',5,Null,6,True,'払込用紙内容照会')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (7,1,'払込用紙・発送状況・照会',5,Null,7,True,'払込用紙発送状況')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (8,1,'入金関係',7,Null,8,True,'入金関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (9,1,'商品発送照会・問合せ',6,Null,9,True,'商品発送問合せ')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (10,1,'商品発送照会・未着',6,Null,10,True,'商品発送受領照会')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (11,1,'製品交換',8,Null,11,True,'製品交換')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (12,2,'販売予定',7,Null,1,True,'販売予定')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (13,2,'製品内容',7,Null,2,True,'製品内容')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (14,2,'在庫照会',7,Null,3,True,'在庫照会')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (15,2,'価格照会',7,Null,4,True,'価格照会')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (16,2,'内容問合せ',7,Null,5,False,'内容問合せ')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (17,2,'組込みミス交換',7,Null,6,False,'組込みミス交換')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (18,2,'その他交換',7,Null,7,False,'その他交換')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (19,2,'（未使用8）',7,Null,8,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (20,2,'（未使用9）',7,Null,9,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (21,2,'（未使用10）',7,Null,10,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (22,2,'（未使用11）',7,Null,11,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (23,3,'ＤＭ停止等・死亡',7,Null,1,True,'ＤＭ停止(①死亡)')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (24,3,'ＤＭ停止等・病気等',7,Null,2,True,'ＤＭ停止(②病気等)')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (25,3,'ＤＭ停止等・種類多い',7,Null,3,True,'ＤＭ停止(③種類多い)')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (26,3,'ＤＭ停止等・特定商品のみ',7,Null,4,True,'ＤＭ停止(④特定商品のみ)')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (27,3,'ＤＭ停止等・理由なし',7,Null,5,True,'ＤＭ停止(⑤理由なし)')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (28,3,'ＤＭ再開・宛名人変更',7,Null,6,True,'ＤＭ再開・宛名人変更')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (29,3,'ＤＭ確認',7,Null,7,True,'ＤＭ確認')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (30,3,'住所変更',7,Null,8,True,'住所変更')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (31,3,'新規登録',7,Null,9,True,'新規登録')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (32,3,'資料送付',7,Null,10,True,'資料送付')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (33,3,'購入履歴',7,Null,11,True,'購入履歴')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (34,4,'お金と切手',7,Null,1,True,'イベント関係（お金と切手）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (35,4,'宇佐神宮',7,Null,2,True,'イベント関係（造幣局 ＩＮ）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (36,4,'桜の通り抜け',7,Null,3,True,'イベント関係（桜の通り抜け）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (37,4,'花のまわり道',7,Null,4,True,'イベント関係（まわり道）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (38,4,'ＴＩＣＣ',7,Null,5,True,'イベント関係（TICC）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (39,4,'国体',7,Null,6,True,'イベント関係（国体）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (40,4,'大阪コインショー',7,Null,7,True,'イベント関係（大阪コイン）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (41,4,'さいたまフェア',7,Null,8,True,'イベント関係（さいたまフェア）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (42,4,'佐伯区民まつり',7,Null,9,True,'イベント関係（名古屋貨幣まつり）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (43,4,'観桜会',7,Null,10,True,'イベント関係（観桜会の電話転送）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (44,4,'アンケート',7,Null,11,True,'イベント関係（予備）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (45,5,'オンライン登録関係',7,Null,1,True,'オンライン登録関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (46,5,'クレジット関係',7,Null,2,True,'クレジット関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (47,5,'金工品Ｇ関係',7,Null,3,True,'金工品Ｇへ電話転送')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (48,5,'広報（工場見学含む）',7,Null,4,True,'広報室（工場見学含む）')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (49,5,'ミントショップ案内',7,Null,5,True,'ミントショップ案内')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (50,5,'金工品懇談会',7,Null,6,True,'桜サポート')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (51,5,'流通貨幣関係',7,Null,7,True,'流通貨幣関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (52,5,'紙幣関係',7,Null,8,True,'紙幣関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (53,5,'ミントクラブ関係',7,Null,9,True,'ミントクラブ関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (54,5,'1円ぬいぐるみ',7,Null,10,True,'1円ぬいぐるみ')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (55,5,'記念貨引換関係',7,Null,11,True,'記念貨引換関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (56,6,'返 品',7,'返品',1,True,'返品')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (57,6,'返 金',7,'返金',2,True,'返金')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (58,6,'（未使用3）',7,Null,3,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (59,6,'（未使用4）',7,Null,4,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (60,6,'（未使用5）',7,Null,5,False,Null)"
End Sub

Private Sub Master_区分_2()
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (61,6,'貨幣ｾｯﾄ処分関係',7,Null,6,True,'貨幣ｾｯﾄ処分関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (62,6,'貨幣洗浄方法',7,Null,7,True,'貨幣洗浄方法')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (63,6,'名称未確定製品予定',7,Null,8,True,'名称未確定製品')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (64,6,'業者販売関係',7,Null,9,True,'業者販売関係')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (65,6,'抽選倍率等(過去）',7,Null,10,True,'抽選倍率等')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (66,6,'（未使用11）',7,Null,11,False,Null)"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (67,7,'別添のとおり',7,Null,1,True,'特殊な問合せ')"
    ExecSQL "INSERT INTO [M_区分] ([区分ID],[ブロックID],[区分名],[集計列ID],[内訳区分],[表示順],[有効],[旧転記名]) VALUES (68,7,'下記のとおり（別添不要）',7,Null,2,True,'特殊な問合せ')"
End Sub

Private Sub Master_担当者()
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (1,'001','髙嶋','名保美','髙嶋 名保美',Null,'パート',Null,Null,1,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (2,'002','堀',Null,'堀 真澄',Null,'パート',Null,Null,2,True,'現行 Excel の Sheet1!B1 が『顧客G』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (3,'003','三原','多恵子','三原 多恵子',Null,'パート',Null,Null,3,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (4,'004','谷口',Null,'谷口 友佳子',Null,'パート',Null,Null,4,True,'現行 Excel の Sheet1!B1 が『顧客G』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (5,'006','国森',Null,'国森 紀子',Null,'パート',Null,Null,5,True,'現行 Excel の Sheet1!B1 が『顧客G』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (6,'007','山口','恵子','山口 恵子',Null,'パート',Null,Null,6,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (7,'008','逢坂','江里加','逢坂 江里加',Null,'パート',Null,Null,7,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (8,'009','西岡',Null,'西岡 由美子',Null,'パート',Null,Null,8,True,'現行 Excel の Sheet1!B1 が『顧客G』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (9,'010','石原','裕子','石原 裕子',Null,'パート',Null,Null,9,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (10,'011','石井','淳子','石井 淳子',Null,'パート',Null,Null,10,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (11,'012','髙橋','理恵子','髙橋 理恵子',Null,'パート',Null,Null,11,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (12,'013','予備',Null,'予備',Null,'パート',Null,Null,12,False,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (13,'014','福島',Null,'福島',Null,'パート',Null,Null,13,True,'現行 Excel の Sheet1!B1 が『髙橋』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (14,'015','鈴木',Null,'鈴木',Null,'パート',Null,Null,14,True,'現行 Excel の Sheet1!B1 が『顧客G』のままで、集計表に別人として計上されていた。Access では担当者マスタで管理する。')"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (15,'016','小澤','朋子','小澤 朋子',Null,'パート',Null,Null,15,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (16,'017','寺本','法子','寺本 法子',Null,'パート',Null,Null,16,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (17,'018','鈴木智','智子','鈴木 智子',Null,'パート',Null,Null,17,True,Null)"
    ExecSQL "INSERT INTO [M_担当者] ([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分],[在籍開始日],[在籍終了日],[表示順],[有効],[備考]) VALUES (18,'100','顧客Ｇ',Null,'顧客G',Null,'職員',Null,Null,18,True,Null)"
End Sub

Private Sub Master_業務項目()
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (1,'①','①受注入力','①　受注入力',1,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (2,'②','②受注チェック','②　受注チェック',2,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (3,'③','③払込書チェック','③　戻り郵便処理／払込書チェック',3,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (4,'④','④戻り郵便','④　ＤＭ処理／戻り郵便',4,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (5,'⑤','⑤架電','⑤　架電',5,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (6,'⑥','⑥その他','⑥　その他（　　　　　　　　）',6,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (7,'⑦','⑦エクセル入力','⑦　エクセル入力',7,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (8,'⑧','⑧エクセルチェック','⑧　エクセルチェック',8,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (9,'⑨','⑨ｱﾝｹｰﾄ入力','⑨　アンケート入力',9,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (10,'⑩','⑩ﾊｶﾞｷﾃﾞｰﾀ入力','⑩　ﾊｶﾞｷﾃﾞｰﾀ入力',10,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (11,'⑪','⑪ﾊｶﾞｷﾃﾞｰﾀﾁｪｯｸ','⑪　ハガキデータチェック',11,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (12,'⑫','⑫新規ｺｰﾄﾞ取り','⑫　新規ｺｰﾄﾞ取り',12,True)"
    ExecSQL "INSERT INTO [M_業務項目] ([業務項目ID],[番号],[項目名],[帳票表示名],[表示順],[有効]) VALUES (13,'⑬','⑬顧客整理','⑬　顧客整理',13,True)"
End Sub
