<%
' ==========================================================================
'  クエリの SQL
'
'  Web 版は Access の保存クエリ (Q_...) を使わない。
'  理由は 2 つ。
'    1. Access を 1 台も使わずに .accdb を配れるようにするため。
'    2. ACE 経由では Access 専用の関数 (Nz など) が使えないため。
'       Access の中では動くのに Web からだけ失敗する、という形を避ける。
'
'  使いかた: 副問い合わせとして埋め込み、並び順は呼ぶ側で付ける。
'      Set rs = DbQuery("SELECT * FROM (" & SQL_未入力チェック() & ") AS C " & _
'                       "WHERE C.[対象日]=? ORDER BY C.[表示順]", Array(dt))
'
'  ※ このファイルは tools/gen_sql_asp.py が src/modSetupQuery.bas から
'     自動生成しています。手で編集せず、生成し直してください。
' ==========================================================================

' 受付入力の区分の選択肢 (Access 版の Q_選択_区分 と同じ)
Function SQL_選択_区分()
    Dim s
    s = "SELECT KB.[区分ID], KB.[ブロックID], BK.[ブロック名], KB.[区分名], BK.[ブロック名] & ' / ' & KB.[区分名] AS "
    s = s & "[表示名], KB.[集計列ID], SC.[集計列名], BK.[製品別], BK.[表示順]*1000 + KB.[表示順] AS [並び順] FROM ([M_区分] "
    s = s & "AS KB INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID]) INNER JOIN [M_集計列] AS SC ON "
    s = s & "KB.[集計列ID]=SC.[集計列ID] WHERE KB.[有効]=True"
    SQL_選択_区分 = s
End Function

' 受付入力の製品の選択肢 (Access 版の Q_選択_製品 と同じ)
Function SQL_選択_製品()
    Dim s
    s = "SELECT PR.[製品ID], PR.[ブロックID], BK.[ブロック名], PR.[製品名], IIf(PR.[製品ID]=0, PR.[製品名], "
    s = s & "BK.[ブロック名] & ' / ' & PR.[製品名]) AS [表示名], PR.[表示順] FROM [M_製品] AS PR INNER JOIN [M_ブロック] "
    s = s & "AS BK ON PR.[ブロックID]=BK.[ブロックID] WHERE PR.[有効]=True AND (PR.[適用開始日] Is Null Or "
    s = s & "PR.[適用開始日] <= Date()) AND (PR.[適用終了日] Is Null Or PR.[適用終了日] >= Date())"
    SQL_選択_製品 = s
End Function

' 集計表の明細 1 行 = 受電 1 件 (Access 版の Q_受電明細 と同じ)
Function SQL_受電明細()
    Dim s
    s = "SELECT J.[受電ID], J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[姓] AS [担当者], OP.[職員区分], "
    s = s & "BK.[ブロック名], PR.[製品名], KB.[区分ID], KB.[区分名], KB.[集計列ID], SC.[集計列名], "
    s = s & "IIf(KB.[集計列ID]=3,J.[件数],0) AS [申込方法], IIf(KB.[集計列ID]=4,J.[件数],0) AS [抽選結果], "
    s = s & "IIf(KB.[集計列ID]=5,J.[件数],0) AS [納付書発送], IIf(KB.[集計列ID]=6,J.[件数],0) AS [商品発送], "
    s = s & "IIf(KB.[集計列ID]=7,J.[件数],0) AS [その他], IIf(KB.[集計列ID]=8,J.[件数],0) AS [商品交換], J.[件数] AS "
    s = s & "[計], J.[備考2], J.[備考3] FROM ((((([T_受電] AS J INNER JOIN [M_担当者] AS OP ON "
    s = s & "J.[担当者ID]=OP.[担当者ID]) INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID]) INNER JOIN [M_集計列] "
    s = s & "AS SC ON KB.[集計列ID]=SC.[集計列ID]) INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID]) "
    s = s & "INNER JOIN [M_製品] AS PR ON J.[製品ID]=PR.[製品ID]) WHERE J.[件数]<>0"
    SQL_受電明細 = s
End Function

' その日の出勤者 (Access 版の Q_日報_出勤 と同じ)
Function SQL_日報_出勤()
    Dim s
    s = "SELECT A.[対象日], OP.[担当者ID], OP.[氏名], OP.[職員区分], A.[勤務時間], A.[備考], OP.[表示順] FROM [T_出勤] "
    s = s & "AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID]"
    SQL_日報_出勤 = s
End Function

' 日報 1 日分。件数はすべて受電から計算 (Access 版の Q_日報_ヘッダ と同じ)
Function SQL_日報_ヘッダ()
    Dim s
    s = "SELECT H.[対象日], H.[回線数], H.[特記事項], H.[職員代替案件], H.[要望], H.[状態], H.[確定日時], (SELECT "
    s = s & "Count(*) FROM [T_出勤] AS A WHERE A.[対象日]=H.[対象日]) AS [出勤者数], IIf(IsNull(R.[申込]),0,R.[申込]) "
    s = s & "AS [申込], IIf(IsNull(R.[抽選]),0,R.[抽選]) AS [抽選], IIf(IsNull(R.[払込用紙]),0,R.[払込用紙]) AS "
    s = s & "[払込用紙], IIf(IsNull(R.[商品発送]),0,R.[商品発送]) AS [商品発送], IIf(IsNull(R.[その他]),0,R.[その他]) AS "
    s = s & "[その他], IIf(IsNull(R.[合計]),0,R.[合計]) AS [合計], IIf(IsNull(R.[内交換]),0,R.[内交換]) AS [内交換], "
    s = s & "IIf(IsNull(R.[内返金]),0,R.[内返金]) AS [内返金], IIf(IsNull(R.[申込_職員]),0,R.[申込_職員]) AS [申込_職員], "
    s = s & "IIf(IsNull(R.[抽選_職員]),0,R.[抽選_職員]) AS [抽選_職員], IIf(IsNull(R.[払込用紙_職員]),0,R.[払込用紙_職員]) AS "
    s = s & "[払込用紙_職員], IIf(IsNull(R.[商品発送_職員]),0,R.[商品発送_職員]) AS [商品発送_職員], "
    s = s & "IIf(IsNull(R.[その他_職員]),0,R.[その他_職員]) AS [その他_職員], IIf(IsNull(R.[合計_職員]),0,R.[合計_職員]) AS "
    s = s & "[合計_職員] FROM [T_日報] AS H LEFT JOIN (SELECT J.[対象日], Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS "
    s = s & "[申込], Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選], Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS "
    s = s & "[払込用紙], Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送], Sum(IIf(KB.[集計列ID] In "
    s = s & "(7,8),J.[件数],0)) AS [その他], Sum(J.[件数]) AS [合計], Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS "
    s = s & "[内交換], Sum(IIf(KB.[内訳区分]='返金',J.[件数],0)) AS [内返金], Sum(IIf(OP.[職員区分]='職員' And "
    s = s & "KB.[集計列ID]=3,J.[件数],0)) AS [申込_職員], Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=4,J.[件数],0)) "
    s = s & "AS [抽選_職員], Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=5,J.[件数],0)) AS [払込用紙_職員], "
    s = s & "Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=6,J.[件数],0)) AS [商品発送_職員], Sum(IIf(OP.[職員区分]='職員' "
    s = s & "And KB.[集計列ID] In (7,8),J.[件数],0)) AS [その他_職員], Sum(IIf(OP.[職員区分]='職員',J.[件数],0)) AS "
    s = s & "[合計_職員] FROM ([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID]) INNER JOIN "
    s = s & "[M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID] GROUP BY J.[対象日]) AS R ON H.[対象日]=R.[対象日]"
    SQL_日報_ヘッダ = s
End Function

' 出勤しているのに実績が 1 件も無い人 (Access 版の Q_未入力チェック と同じ)
Function SQL_未入力チェック()
    Dim s
    s = "SELECT A.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[表示順] FROM [T_出勤] AS A INNER JOIN "
    s = s & "[M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] WHERE NOT EXISTS (SELECT 1 FROM [T_受電] AS J WHERE "
    s = s & "J.[対象日]=A.[対象日] AND J.[担当者ID]=A.[担当者ID] AND J.[件数]<>0) AND NOT EXISTS (SELECT 1 FROM "
    s = s & "[T_業務実績] AS W WHERE W.[対象日]=A.[対象日] AND W.[担当者ID]=A.[担当者ID] AND W.[件数]<>0)"
    SQL_未入力チェック = s
End Function
%>
