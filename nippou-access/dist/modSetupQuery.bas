Attribute VB_Name = "modSetupQuery"
Option Compare Database
Option Explicit

'==============================================================================
' クエリ定義
'
' 設計の要:
'   日報 (R_日報) も 集計表 (R_集計表) も、同じ T_受電 1 本から集計する。
'   現行 Excel では「集計用フォームの SUM」と「集計表の SUMIF」という
'   別々の計算経路があり、しかも SUMIF の範囲が行ごとにバラバラだったため
'   数値が食い違っていた。ここでは経路が 1 本しかないので構造的に一致する。
'
' 帳票欄 と 集計列 の対応 (現行 印刷用シート 18 行目の数式から復元):
'   申込      = 集計列 3 (申込方法)                 … 旧 C29 + D29
'   抽選      = 集計列 4 (抽選結果)                 … 旧 E29
'   払込用紙  = 集計列 5 (納付書発送)               … 旧 F29 = SUM(F6:I28)
'   商品発送  = 集計列 6 (商品発送)                 … 旧 K29 + L29
'   その他    = 集計列 7 + 8 (その他 + 商品交換)    … 旧 J29 + M29 + N60
'   内 交換   = 集計列 8 (商品交換)                 … 旧 M29
'   内 返金   = 内訳区分 = '返金'                   … 旧 D60
'==============================================================================

Public Sub Setup_Queries()
    Echo_ "クエリを作成しています..."

    '--- 選択候補 (在籍期間・適用期間で絞る) ------------------------------------
    ' 退職者を候補から消しても、過去データの担当者名は保持される。
    ' コンボボックスの行ソースにはパラメータを渡せないので、
    ' 画面用 (今日時点で在籍) とプログラム用 (基準日指定) の 2 本を用意する。
    SaveQuery "Q_選択_担当者", _
        "SELECT OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[職員区分], OP.[表示順] " & _
        "FROM [M_担当者] AS OP " & _
        "WHERE OP.[有効]=True " & _
        "  AND (OP.[在籍開始日] Is Null Or OP.[在籍開始日] <= Date()) " & _
        "  AND (OP.[在籍終了日] Is Null Or OP.[在籍終了日] >= Date()) " & _
        "ORDER BY OP.[表示順];"

    SaveQuery "Q_選択_担当者_基準日", _
        "PARAMETERS [基準日] DateTime; " & _
        "SELECT OP.[担当者ID], OP.[担当者コード], OP.[氏名], OP.[職員区分], OP.[表示順] " & _
        "FROM [M_担当者] AS OP " & _
        "WHERE OP.[有効]=True " & _
        "  AND (OP.[在籍開始日] Is Null Or OP.[在籍開始日] <= [基準日]) " & _
        "  AND (OP.[在籍終了日] Is Null Or OP.[在籍終了日] >= [基準日]) " & _
        "ORDER BY OP.[表示順];"

    SaveQuery "Q_選択_区分", _
        "SELECT KB.[区分ID], KB.[ブロックID], BK.[ブロック名], KB.[区分名], " & _
        "       BK.[ブロック名] & ' / ' & KB.[区分名] AS [表示名], " & _
        "       KB.[集計列ID], SC.[集計列名], BK.[製品別], " & _
        "       BK.[表示順]*1000 + KB.[表示順] AS [並び順] " & _
        "FROM ([M_区分] AS KB INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID]) " & _
        "     INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID] " & _
        "WHERE KB.[有効]=True " & _
        "ORDER BY BK.[表示順]*1000 + KB.[表示順];"

    SaveQuery "Q_選択_製品", _
        "SELECT PR.[製品ID], PR.[ブロックID], BK.[ブロック名], PR.[製品名], " & _
        "       IIf(PR.[製品ID]=0, PR.[製品名], BK.[ブロック名] & ' / ' & PR.[製品名]) AS [表示名], " & _
        "       PR.[表示順] " & _
        "FROM [M_製品] AS PR INNER JOIN [M_ブロック] AS BK ON PR.[ブロックID]=BK.[ブロックID] " & _
        "WHERE PR.[有効]=True " & _
        "  AND (PR.[適用開始日] Is Null Or PR.[適用開始日] <= Date()) " & _
        "  AND (PR.[適用終了日] Is Null Or PR.[適用終了日] >= Date()) " & _
        "ORDER BY PR.[ブロックID], PR.[表示順];"

    '--- 明細 (現行「集計表」データシート相当) ----------------------------------
    SaveQuery "Q_受電明細", _
        "SELECT J.[受電ID], J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名], " & _
        "       OP.[姓] AS [担当者], OP.[職員区分], BK.[ブロック名], " & _
        "       PR.[製品名], KB.[区分ID], KB.[区分名], KB.[集計列ID], SC.[集計列名], " & _
        "       IIf(KB.[集計列ID]=3,J.[件数],0) AS [申込方法], " & _
        "       IIf(KB.[集計列ID]=4,J.[件数],0) AS [抽選結果], " & _
        "       IIf(KB.[集計列ID]=5,J.[件数],0) AS [納付書発送], " & _
        "       IIf(KB.[集計列ID]=6,J.[件数],0) AS [商品発送], " & _
        "       IIf(KB.[集計列ID]=7,J.[件数],0) AS [その他], " & _
        "       IIf(KB.[集計列ID]=8,J.[件数],0) AS [商品交換], " & _
        "       J.[件数] AS [計], J.[備考2], J.[備考3] " & _
        "FROM ((((([T_受電] AS J " & _
        "  INNER JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID]) " & _
        "  INNER JOIN [M_区分]   AS KB ON J.[区分ID]=KB.[区分ID]) " & _
        "  INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID]) " & _
        "  INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID]) " & _
        "  INNER JOIN [M_製品]   AS PR ON J.[製品ID]=PR.[製品ID]) " & _
        "WHERE J.[件数]<>0;"

    '--- 日別 集計列クロス集計 (現行「集計表」B2:I7 相当) -----------------------
    SaveQuery "Q_日別集計", _
        "TRANSFORM Sum(J.[件数]) AS [件数計] " & _
        "SELECT J.[対象日], Sum(J.[件数]) AS [計] " & _
        "FROM ([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID]) " & _
        "     INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID] " & _
        "GROUP BY J.[対象日] " & _
        "PIVOT SC.[集計列名] " & _
        "IN ('申込方法','抽選結果','納付書発送','商品発送','その他','商品交換');"

    '--- 日報 (帳票) 用 --------------------------------------------------------
    ' 帳票 1 行分の集計。1 論理行 1023 文字の制限に掛かるので変数に組み立てる。
    Dim s As String
    s = "SELECT J.[対象日], "
    s = s & " Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込], "
    s = s & " Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選], "
    s = s & " Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [払込用紙], "
    s = s & " Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送], "
    s = s & " Sum(IIf(KB.[集計列ID] In (7,8),J.[件数],0)) AS [その他], "
    s = s & " Sum(J.[件数]) AS [合計], "
    s = s & " Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [内交換], "
    s = s & " Sum(IIf(KB.[内訳区分]='返金',J.[件数],0)) AS [内返金], "
    s = s & " Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=3,J.[件数],0)) AS [申込_職員], "
    s = s & " Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=4,J.[件数],0)) AS [抽選_職員], "
    s = s & " Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=5,J.[件数],0)) AS [払込用紙_職員], "
    s = s & " Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID]=6,J.[件数],0)) AS [商品発送_職員], "
    s = s & " Sum(IIf(OP.[職員区分]='職員' And KB.[集計列ID] In (7,8),J.[件数],0)) AS [その他_職員], "
    s = s & " Sum(IIf(OP.[職員区分]='職員',J.[件数],0)) AS [合計_職員] "
    s = s & "FROM ([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID]) "
    s = s & "     INNER JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID] "
    s = s & "GROUP BY J.[対象日];"
    SaveQuery "Q_日報_受電", s

    SaveQuery "Q_日報_業務", _
        "SELECT W.[対象日], TM.[業務項目ID], TM.[番号], TM.[項目名], TM.[帳票表示名], " & _
        "       TM.[表示順], Sum(W.[件数]) AS [件数] " & _
        "FROM [T_業務実績] AS W INNER JOIN [M_業務項目] AS TM " & _
        "     ON W.[業務項目ID]=TM.[業務項目ID] " & _
        "GROUP BY W.[対象日], TM.[業務項目ID], TM.[番号], TM.[項目名], TM.[帳票表示名], TM.[表示順];"

    SaveQuery "Q_日報_出勤", _
        "SELECT A.[対象日], OP.[担当者ID], OP.[氏名], OP.[職員区分], " & _
        "       A.[勤務時間], A.[備考], OP.[表示順] " & _
        "FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] " & _
        "ORDER BY A.[対象日], OP.[表示順];"

    ' 帳票 1 ページ分をまとめた 1 行。R_日報 のレコードソース。
    SaveQuery "Q_日報_ヘッダ", _
        "SELECT H.[対象日], H.[回線数], H.[特記事項], H.[職員代替案件], H.[要望], " & _
        "       H.[状態], H.[確定日時], " & _
        "       (SELECT Count(*) FROM [T_出勤] AS A WHERE A.[対象日]=H.[対象日]) AS [出勤者数], " & _
        "       Nz(R.[申込],0) AS [申込], Nz(R.[抽選],0) AS [抽選], " & _
        "       Nz(R.[払込用紙],0) AS [払込用紙], Nz(R.[商品発送],0) AS [商品発送], " & _
        "       Nz(R.[その他],0) AS [その他], Nz(R.[合計],0) AS [合計], " & _
        "       Nz(R.[内交換],0) AS [内交換], Nz(R.[内返金],0) AS [内返金], " & _
        "       Nz(R.[申込_職員],0) AS [申込_職員], Nz(R.[抽選_職員],0) AS [抽選_職員], " & _
        "       Nz(R.[払込用紙_職員],0) AS [払込用紙_職員], " & _
        "       Nz(R.[商品発送_職員],0) AS [商品発送_職員], " & _
        "       Nz(R.[その他_職員],0) AS [その他_職員], Nz(R.[合計_職員],0) AS [合計_職員] " & _
        "FROM [T_日報] AS H LEFT JOIN [Q_日報_受電] AS R ON H.[対象日]=R.[対象日];"

    '--- 担当者別・チェック用 ---------------------------------------------------
    SaveQuery "Q_担当者別日次", _
        "SELECT J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名], " & _
        "       Sum(J.[件数]) AS [受電件数] " & _
        "FROM [T_受電] AS J INNER JOIN [M_担当者] AS OP ON J.[担当者ID]=OP.[担当者ID] " & _
        "GROUP BY J.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名];"

    ' 出勤登録があるのに実績が 1 件も無い担当者。
    ' 現行 Excel で「特定の人のデータが出てこない」ことに誰も気付けなかった問題への対策。
    SaveQuery "Q_未入力チェック", _
        "SELECT A.[対象日], OP.[担当者ID], OP.[担当者コード], OP.[氏名] " & _
        "FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] " & _
        "WHERE NOT EXISTS (SELECT 1 FROM [T_受電] AS J " & _
        "                  WHERE J.[対象日]=A.[対象日] AND J.[担当者ID]=A.[担当者ID] AND J.[件数]<>0) " & _
        "  AND NOT EXISTS (SELECT 1 FROM [T_業務実績] AS W " & _
        "                  WHERE W.[対象日]=A.[対象日] AND W.[担当者ID]=A.[担当者ID] AND W.[件数]<>0) " & _
        "ORDER BY A.[対象日], OP.[表示順];"

    '--- 週次 (現行「集計表(VBA版)」相当) --------------------------------------
    SaveQuery "Q_週次明細", _
        "PARAMETERS [開始日] DateTime, [終了日] DateTime; " & _
        "SELECT * FROM [Q_受電明細] " & _
        "WHERE [対象日] Between [開始日] And [終了日] " & _
        "ORDER BY [対象日], [氏名], [区分ID];"

    SaveQuery "Q_週次集計", _
        "PARAMETERS [開始日] DateTime, [終了日] DateTime; " & _
        "SELECT J.[対象日], " & _
        " Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法], " & _
        " Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選結果], " & _
        " Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [納付書発送], " & _
        " Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送], " & _
        " Sum(IIf(KB.[集計列ID]=7,J.[件数],0)) AS [その他], " & _
        " Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [商品交換], " & _
        " Sum(J.[件数]) AS [計] " & _
        "FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID] " & _
        "WHERE J.[対象日] Between [開始日] And [終了日] " & _
        "GROUP BY J.[対象日] " & _
        "ORDER BY J.[対象日];"

    ' 画面のサブフォームは RecordSource に WHERE を足して使うので、
    ' パラメータを持たない版も用意しておく (パラメータ付きは帳票・プログラム用)。
    SaveQuery "Q_週次集計指定", _
        "SELECT J.[対象日], " & _
        " Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法], " & _
        " Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選結果], " & _
        " Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [納付書発送], " & _
        " Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送], " & _
        " Sum(IIf(KB.[集計列ID]=7,J.[件数],0)) AS [その他], " & _
        " Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [商品交換], " & _
        " Sum(J.[件数]) AS [計] " & _
        "FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID] " & _
        "GROUP BY J.[対象日];"

    Echo_ "  クエリ作成 完了 (" & CurrentDb.QueryDefs.Count & " 件)"
End Sub
