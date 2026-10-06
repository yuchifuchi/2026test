<%
' ============================================================
'  SQL の置き場所
'
'  ・Access の「保存クエリ」は作らない。中で Access 専用の関数を使っていると、
'    ASP から（ACE OLEDB 経由で）呼んだときに、実行して初めて落ちるため。
'    SQL はすべてここに文字列として持ち、画面からは関数名で呼ぶ。
'  ・値は必ず ? で渡す（文字列でつなげない）。
'  ・NULL を 0 にするときは IIf(IsNull(Sum(x)), 0, Sum(x))。Nz() は ASP からは使えない。
'  ・派生表（FROM ( ... ) AS D）の外で ORDER BY に使う列は、必ず内側の SELECT にも入れる。
'  ・JOIN が 2 つ以上なら、(A JOIN B ON ..) JOIN C ON .. のようにかっこで囲む（ACE の決まり）。
'  ・ON には「表.[列] = 表.[列]」だけを書く。日付などで絞るのは派生表の中で。
'  ・日付は「その日の 0 時以上、翌日の 0 時未満」の 2 つの ? で渡す（時刻が入っていても漏れないように）。
'
'  tools/vbsim/sqlcheck.py がこの決まりを機械的に検査し、
'  tools/test_sql.py が検査用の .accdb に対して 1 本ずつ実行して期待値と突き合わせる。
' ============================================================

' ------------------------------------------------------------
'  数値の出どころ（ただ 1 本）
'  日報の紙・週の集計表・入力画面の合計・入力もれチェックは、すべてこの SELECT を元にする。
'  現行の Excel は「入力シート → 転記シート → 集計表」と経路が 2 本あって数字が食い違った。
'  経路を 1 本にすれば、食い違いは構造的に起きない。
'  ? : (開始日, 終了日の翌日)
' ------------------------------------------------------------
Function SqlCallsBase()
    Dim s
    s = "SELECT J.[対象日], J.[担当者ID], J.[区分ID], J.[製品ID], J.[件数], "
    ' [日] は時刻を落とした日付。日ごとに数えるときはこれでまとめる
    ' （Access で直接直して時刻が入った行があっても、同じ日が 2 つに割れないように）
    s = s & "DateSerial(Year(J.[対象日]), Month(J.[対象日]), Day(J.[対象日])) AS [日], "
    s = s & "K.[集計列ID], K.[内訳区分], K.[ブロックID], T.[職員区分], "
    s = s & "IIf(T.[職員区分] = '職員', J.[件数], 0) AS [職員件数] "
    s = s & "FROM ([T_受電] AS J INNER JOIN [M_区分] AS K ON J.[区分ID] = K.[区分ID]) "
    s = s & "INNER JOIN [M_担当者] AS T ON J.[担当者ID] = T.[担当者ID] "
    s = s & "WHERE J.[対象日] >= ? AND J.[対象日] < ?"
    SqlCallsBase = s
End Function

' 上の元の SELECT を、1 人ぶんに絞ったもの。? : (開始日, 終了日の翌日, 担当者ID)
Function SqlCallsBaseOnePerson()
    SqlCallsBaseOnePerson = SqlCallsBase() & " AND J.[担当者ID] = ?"
End Function

' 5 列（申込／抽選／払込用紙／商品発送／その他）の合計と、そのうち職員が受けた数。
' 件数が 0 の列も必ず 5 行出るように、集計列のマスタから LEFT JOIN する。
Function SqlColTotalsHead()
    Dim s
    s = "SELECT C.[集計列ID], C.[集計列名], C.[表示順], "
    s = s & "IIf(IsNull(Sum(D.[件数])), 0, Sum(D.[件数])) AS [件数計], "
    s = s & "IIf(IsNull(Sum(D.[職員件数])), 0, Sum(D.[職員件数])) AS [職員計] "
    s = s & "FROM [M_集計列] AS C LEFT JOIN ("
    SqlColTotalsHead = s
End Function

Function SqlColTotalsTail()
    SqlColTotalsTail = ") AS D ON C.[集計列ID] = D.[集計列ID] GROUP BY C.[集計列ID], C.[集計列名], C.[表示順] ORDER BY C.[表示順]"
End Function

' ? : (開始日, 終了日の翌日)
Function SqlColumnTotals()
    SqlColumnTotals = SqlColTotalsHead() & SqlCallsBase() & SqlColTotalsTail()
End Function

' ? : (開始日, 終了日の翌日, 担当者ID)
Function SqlColumnTotalsOnePerson()
    SqlColumnTotalsOnePerson = SqlColTotalsHead() & SqlCallsBaseOnePerson() & SqlColTotalsTail()
End Function

' 「内 交換 ○件／内 返金 ○件」。? : (開始日, 終了日の翌日)
Function SqlBreakdownTotals()
    SqlBreakdownTotals = "SELECT D.[内訳区分], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBase() & ") AS D WHERE D.[内訳区分] IS NOT NULL GROUP BY D.[内訳区分]"
End Function

' 1 人ぶんの内訳。? : (開始日, 終了日の翌日, 担当者ID)
Function SqlBreakdownTotalsOnePerson()
    SqlBreakdownTotalsOnePerson = "SELECT D.[内訳区分], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBaseOnePerson() & ") AS D WHERE D.[内訳区分] IS NOT NULL GROUP BY D.[内訳区分]"
End Function

' ---- 週の集計表で使う（どれも元は SqlCallsBase） ----
' 日 × 集計列。? : (開始日, 終了日の翌日)
Function SqlDayColumnTotals()
    SqlDayColumnTotals = "SELECT D.[日], D.[集計列ID], Sum(D.[件数]) AS [件数計], Sum(D.[職員件数]) AS [職員計] FROM (" & SqlCallsBase() & ") AS D GROUP BY D.[日], D.[集計列ID]"
End Function

' 日 × 内訳（交換・返金）。? : (開始日, 終了日の翌日)
Function SqlDayBreakdownTotals()
    SqlDayBreakdownTotals = "SELECT D.[日], D.[内訳区分], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBase() & ") AS D WHERE D.[内訳区分] IS NOT NULL GROUP BY D.[日], D.[内訳区分]"
End Function

' 日 × 区分。? : (開始日, 終了日の翌日)
Function SqlDayKubunTotals()
    SqlDayKubunTotals = "SELECT D.[日], D.[区分ID], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBase() & ") AS D GROUP BY D.[日], D.[区分ID]"
End Function

' 日 × 担当者。? : (開始日, 終了日の翌日)
Function SqlDayPersonTotals()
    SqlDayPersonTotals = "SELECT D.[日], D.[担当者ID], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBase() & ") AS D GROUP BY D.[日], D.[担当者ID]"
End Function

' 担当者ごとの合計（入力もれチェック）。? : (開始日, 終了日の翌日)
Function SqlPersonTotals()
    SqlPersonTotals = "SELECT D.[担当者ID], Sum(D.[件数]) AS [件数計] FROM (" & SqlCallsBase() & ") AS D GROUP BY D.[担当者ID]"
End Function

' ---- 突き合わせ（入力もれチェック） ----
' T_受電 を直接数えたものと、元の SELECT を通したもの。
' 参照整合性があるので必ず一致するはずだが、一致しなければ「どこかで行が落ちている」ので赤で出す。
Function SqlRawCallsCheck()
    SqlRawCallsCheck = "SELECT Count(J.[受電ID]) AS [行数], IIf(IsNull(Sum(J.[件数])), 0, Sum(J.[件数])) AS [件数計] FROM [T_受電] AS J WHERE J.[対象日] >= ? AND J.[対象日] < ?"
End Function

Function SqlBaseCallsCheck()
    SqlBaseCallsCheck = "SELECT Count(D.[区分ID]) AS [行数], IIf(IsNull(Sum(D.[件数])), 0, Sum(D.[件数])) AS [件数計] FROM (" & SqlCallsBase() & ") AS D"
End Function

' ------------------------------------------------------------
'  受付入力（1 人・1 日ぶんの明細）
' ------------------------------------------------------------
' ? : (開始日, 終了日の翌日, 担当者ID)
Function SqlCallRowsOnePerson()
    SqlCallRowsOnePerson = "SELECT J.[受電ID], J.[区分ID], J.[製品ID], J.[件数] FROM [T_受電] AS J WHERE J.[対象日] >= ? AND J.[対象日] < ? AND J.[担当者ID] = ?"
End Function

' 別の画面で同時に直されていないかを見るための「版」。? : (開始日, 終了日の翌日, 担当者ID)
Function SqlCallVersion()
    SqlCallVersion = "SELECT Count(J.[受電ID]) AS [行数], Sum(J.[件数]) AS [件数計], Max(J.[更新日時]) AS [最終] FROM [T_受電] AS J WHERE J.[対象日] >= ? AND J.[対象日] < ? AND J.[担当者ID] = ?"
End Function

Function SqlCallInsert()
    SqlCallInsert = "INSERT INTO [T_受電] ([対象日], [担当者ID], [区分ID], [製品ID], [件数], [登録日時], [更新日時], [登録者]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
End Function

Function SqlCallUpdate()
    SqlCallUpdate = "UPDATE [T_受電] SET [件数] = ?, [更新日時] = ?, [登録者] = ? WHERE [受電ID] = ?"
End Function

Function SqlCallDelete()
    SqlCallDelete = "DELETE FROM [T_受電] WHERE [受電ID] = ?"
End Function

' ------------------------------------------------------------
'  その他業務（①～⑬）
' ------------------------------------------------------------
' 1 人・1 日ぶん。? : (開始日, 終了日の翌日, 担当者ID)
Function SqlTaskRowsOnePerson()
    SqlTaskRowsOnePerson = "SELECT R.[実績ID], R.[業務項目ID], R.[件数] FROM [T_業務実績] AS R WHERE R.[対象日] >= ? AND R.[対象日] < ? AND R.[担当者ID] = ?"
End Function

' 日報の下段。全員の合計。? : (開始日, 終了日の翌日)
Function SqlTaskTotals()
    Dim s
    s = "SELECT G.[業務項目ID], G.[番号], G.[項目名], G.[帳票表示名], G.[表示順], G.[有効], "
    s = s & "IIf(IsNull(Sum(W.[件数])), 0, Sum(W.[件数])) AS [件数計] "
    s = s & "FROM [M_業務項目] AS G LEFT JOIN "
    s = s & "(SELECT R.[業務項目ID], R.[件数] FROM [T_業務実績] AS R WHERE R.[対象日] >= ? AND R.[対象日] < ?) AS W "
    s = s & "ON G.[業務項目ID] = W.[業務項目ID] "
    s = s & "GROUP BY G.[業務項目ID], G.[番号], G.[項目名], G.[帳票表示名], G.[表示順], G.[有効] "
    s = s & "ORDER BY G.[表示順], G.[業務項目ID]"
    SqlTaskTotals = s
End Function

' 担当者ごとの その他業務 の合計（入力もれチェック）。? : (開始日, 終了日の翌日)
Function SqlTaskPersonTotals()
    SqlTaskPersonTotals = "SELECT R.[担当者ID], Sum(R.[件数]) AS [件数計] FROM [T_業務実績] AS R WHERE R.[対象日] >= ? AND R.[対象日] < ? GROUP BY R.[担当者ID]"
End Function

Function SqlTaskInsert()
    SqlTaskInsert = "INSERT INTO [T_業務実績] ([対象日], [担当者ID], [業務項目ID], [件数]) VALUES (?, ?, ?, ?)"
End Function

Function SqlTaskUpdate()
    SqlTaskUpdate = "UPDATE [T_業務実績] SET [件数] = ? WHERE [実績ID] = ?"
End Function

Function SqlTaskDelete()
    SqlTaskDelete = "DELETE FROM [T_業務実績] WHERE [実績ID] = ?"
End Function

' ------------------------------------------------------------
'  出勤
' ------------------------------------------------------------
' ? : (開始日, 終了日の翌日)
Function SqlAttendance()
    Dim s
    s = "SELECT A.[対象日], A.[担当者ID], A.[勤務時間], A.[備考], T.[姓], T.[氏名], T.[職員区分], T.[表示順] "
    s = s & "FROM [T_出勤] AS A INNER JOIN [M_担当者] AS T ON A.[担当者ID] = T.[担当者ID] "
    s = s & "WHERE A.[対象日] >= ? AND A.[対象日] < ? "
    s = s & "ORDER BY A.[対象日], T.[表示順], T.[担当者ID]"
    SqlAttendance = s
End Function

Function SqlAttendanceDeleteDay()
    SqlAttendanceDeleteDay = "DELETE FROM [T_出勤] WHERE [対象日] >= ? AND [対象日] < ?"
End Function

Function SqlAttendanceInsert()
    SqlAttendanceInsert = "INSERT INTO [T_出勤] ([対象日], [担当者ID], [勤務時間], [備考]) VALUES (?, ?, ?, ?)"
End Function

' ------------------------------------------------------------
'  日報（記述欄・回線数・状態）
' ------------------------------------------------------------
' ? : (開始日, 終了日の翌日)
Function SqlDaily()
    SqlDaily = "SELECT N.[対象日], N.[回線数], N.[特記事項], N.[職員代替案件], N.[要望], N.[状態], N.[確定日時], N.[更新日時] FROM [T_日報] AS N WHERE N.[対象日] >= ? AND N.[対象日] < ? ORDER BY N.[対象日]"
End Function

' いちばん新しい「回線数」（新しい日の初期値に使う）。? : (この日)
Function SqlLastLines()
    SqlLastLines = "SELECT TOP 1 N.[回線数], N.[対象日] FROM [T_日報] AS N WHERE N.[対象日] < ? AND N.[回線数] IS NOT NULL ORDER BY N.[対象日] DESC"
End Function

Function SqlDailyInsert()
    SqlDailyInsert = "INSERT INTO [T_日報] ([対象日], [回線数], [特記事項], [職員代替案件], [要望], [状態], [確定日時], [更新日時]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
End Function

' ? : (回線数, 特記事項, 職員代替案件, 要望, 状態, 確定日時, 更新日時, 開始日, 終了日の翌日)
Function SqlDailyUpdate()
    SqlDailyUpdate = "UPDATE [T_日報] SET [回線数] = ?, [特記事項] = ?, [職員代替案件] = ?, [要望] = ?, [状態] = ?, [確定日時] = ?, [更新日時] = ? WHERE [対象日] >= ? AND [対象日] < ?"
End Function

' ? : (状態, 確定日時, 更新日時, 開始日, 終了日の翌日)
Function SqlDailySetState()
    SqlDailySetState = "UPDATE [T_日報] SET [状態] = ?, [確定日時] = ?, [更新日時] = ? WHERE [対象日] >= ? AND [対象日] < ?"
End Function

' ------------------------------------------------------------
'  マスタ
' ------------------------------------------------------------
Function SqlStaffAll()
    SqlStaffAll = "SELECT T.[担当者ID], T.[担当者コード], T.[姓], T.[名], T.[氏名], T.[カナ], T.[ログオン名], T.[職員区分], T.[在籍開始日], T.[在籍終了日], T.[表示順], T.[有効], T.[備考] FROM [M_担当者] AS T ORDER BY T.[表示順], T.[担当者ID]"
End Function

' 担当者の数（設置チェック・probe5 で使う）
Function SqlStaffCount()
    SqlStaffCount = "SELECT Count(T.[担当者ID]) AS [件数] FROM [M_担当者] AS T"
End Function

' ログオン名で「職員」かを調べる（ROLE が空のときだけ使う）。? : (ログオン名)
Function SqlStaffByLogon()
    SqlStaffByLogon = "SELECT Count(T.[担当者ID]) AS [件数] FROM [M_担当者] AS T WHERE T.[ログオン名] = ? AND T.[職員区分] = '職員' AND T.[有効] = True"
End Function

Function SqlStaffMaxId()
    SqlStaffMaxId = "SELECT Max(T.[担当者ID]) AS [最大] FROM [M_担当者] AS T"
End Function

Function SqlStaffInsert()
    SqlStaffInsert = "INSERT INTO [M_担当者] ([担当者ID], [担当者コード], [姓], [名], [氏名], [カナ], [ログオン名], [職員区分], [在籍開始日], [在籍終了日], [表示順], [有効], [備考]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
End Function

Function SqlStaffUpdate()
    SqlStaffUpdate = "UPDATE [M_担当者] SET [担当者コード] = ?, [姓] = ?, [名] = ?, [氏名] = ?, [カナ] = ?, [ログオン名] = ?, [職員区分] = ?, [在籍開始日] = ?, [在籍終了日] = ?, [表示順] = ?, [有効] = ?, [備考] = ? WHERE [担当者ID] = ?"
End Function

Function SqlColumnsAll()
    SqlColumnsAll = "SELECT C.[集計列ID], C.[集計列名], C.[表示順] FROM [M_集計列] AS C ORDER BY C.[表示順], C.[集計列ID]"
End Function

Function SqlColumnRename()
    SqlColumnRename = "UPDATE [M_集計列] SET [集計列名] = ? WHERE [集計列ID] = ?"
End Function

Function SqlBlocksAll()
    SqlBlocksAll = "SELECT B.[ブロックID], B.[ブロック名], B.[製品別], B.[表示順] FROM [M_ブロック] AS B ORDER BY B.[表示順], B.[ブロックID]"
End Function

Function SqlBlockMaxId()
    SqlBlockMaxId = "SELECT Max(B.[ブロックID]) AS [最大] FROM [M_ブロック] AS B"
End Function

Function SqlBlockInsert()
    SqlBlockInsert = "INSERT INTO [M_ブロック] ([ブロックID], [ブロック名], [製品別], [表示順]) VALUES (?, ?, ?, ?)"
End Function

Function SqlBlockUpdate()
    SqlBlockUpdate = "UPDATE [M_ブロック] SET [ブロック名] = ?, [製品別] = ?, [表示順] = ? WHERE [ブロックID] = ?"
End Function

' ブロックの「製品別」を外してよいか（製品を指定した実績が残っていないか）。? : (ブロックID)
Function SqlBlockProductRows()
    SqlBlockProductRows = "SELECT Count(J.[受電ID]) AS [件数] FROM [T_受電] AS J INNER JOIN [M_区分] AS K ON J.[区分ID] = K.[区分ID] WHERE K.[ブロックID] = ? AND J.[製品ID] <> 0"
End Function

Function SqlKubunAll()
    SqlKubunAll = "SELECT K.[区分ID], K.[ブロックID], K.[区分名], K.[集計列ID], K.[内訳区分], K.[表示順], K.[有効], K.[旧転記名] FROM [M_区分] AS K ORDER BY K.[ブロックID], K.[表示順], K.[区分ID]"
End Function

Function SqlKubunMaxId()
    SqlKubunMaxId = "SELECT Max(K.[区分ID]) AS [最大] FROM [M_区分] AS K"
End Function

Function SqlKubunInsert()
    SqlKubunInsert = "INSERT INTO [M_区分] ([区分ID], [ブロックID], [区分名], [集計列ID], [内訳区分], [表示順], [有効], [旧転記名]) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
End Function

Function SqlKubunUpdate()
    SqlKubunUpdate = "UPDATE [M_区分] SET [ブロックID] = ?, [区分名] = ?, [集計列ID] = ?, [内訳区分] = ?, [表示順] = ?, [有効] = ?, [旧転記名] = ? WHERE [区分ID] = ?"
End Function

' 区分をほかのブロックへ移してよいか（実績があると、入力画面のどこに出すかが変わる）。? : (区分ID)
Function SqlKubunUseCount()
    SqlKubunUseCount = "SELECT Count(J.[受電ID]) AS [件数] FROM [T_受電] AS J WHERE J.[区分ID] = ?"
End Function

Function SqlProductsAll()
    SqlProductsAll = "SELECT P.[製品ID], P.[ブロックID], P.[製品名], P.[適用開始日], P.[適用終了日], P.[表示順], P.[有効] FROM [M_製品] AS P WHERE P.[製品ID] <> 0 ORDER BY P.[ブロックID], P.[表示順], P.[製品ID]"
End Function

Function SqlProductMaxId()
    SqlProductMaxId = "SELECT Max(P.[製品ID]) AS [最大] FROM [M_製品] AS P"
End Function

Function SqlProductInsert()
    SqlProductInsert = "INSERT INTO [M_製品] ([製品ID], [ブロックID], [製品名], [適用開始日], [適用終了日], [表示順], [有効]) VALUES (?, ?, ?, ?, ?, ?, ?)"
End Function

Function SqlProductUpdate()
    SqlProductUpdate = "UPDATE [M_製品] SET [ブロックID] = ?, [製品名] = ?, [適用開始日] = ?, [適用終了日] = ?, [表示順] = ?, [有効] = ? WHERE [製品ID] = ?"
End Function

Function SqlProductUseCount()
    SqlProductUseCount = "SELECT Count(J.[受電ID]) AS [件数] FROM [T_受電] AS J WHERE J.[製品ID] = ?"
End Function

Function SqlTaskItemsAll()
    SqlTaskItemsAll = "SELECT G.[業務項目ID], G.[番号], G.[項目名], G.[帳票表示名], G.[表示順], G.[有効] FROM [M_業務項目] AS G ORDER BY G.[表示順], G.[業務項目ID]"
End Function

Function SqlTaskItemMaxId()
    SqlTaskItemMaxId = "SELECT Max(G.[業務項目ID]) AS [最大] FROM [M_業務項目] AS G"
End Function

Function SqlTaskItemInsert()
    SqlTaskItemInsert = "INSERT INTO [M_業務項目] ([業務項目ID], [番号], [項目名], [帳票表示名], [表示順], [有効]) VALUES (?, ?, ?, ?, ?, ?)"
End Function

Function SqlTaskItemUpdate()
    SqlTaskItemUpdate = "UPDATE [M_業務項目] SET [番号] = ?, [項目名] = ?, [帳票表示名] = ?, [表示順] = ?, [有効] = ? WHERE [業務項目ID] = ?"
End Function
%>
