Attribute VB_Name = "modSetup"
Option Compare Database
Option Explicit

'==============================================================================
' 電話応対日報 集計システム  ―  データベース構築モジュール
'
' 使い方:
'   1. 空の .accdb を作成して開く
'   2. Alt+F11 → 挿入 → 標準モジュール に src\ 配下の .bas をすべて貼り付ける
'   3. このモジュールの Setup_All を実行する (F5)
'
' Setup_All は何度実行しても同じ結果になる (既存の表・クエリ・画面を作り直す)。
' ただし T_ で始まる実績データは、DROP_DATA_TABLES = False の間は保持される。
'==============================================================================

' 実績テーブルまで作り直すかどうか。通常は False のままにすること。
Private Const DROP_DATA_TABLES As Boolean = False

Public Const APP_NAME As String = "電話応対日報 集計システム"
Public Const APP_VERSION As String = "1.0"

'------------------------------------------------------------------------------
' 入口
'------------------------------------------------------------------------------
Public Sub Setup_All()
    Dim started As Date
    started = Now

    On Error GoTo Fail

    Echo_ "=== " & APP_NAME & " セットアップ開始 ==="
    CreateTables
    Setup_Master
    Setup_Queries
    Setup_Forms
    Setup_Reports
    Setup_AppOptions

    Echo_ "=== セットアップ完了 (" & Format(Now - started, "nn分ss秒") & ") ==="
    MsgBox APP_NAME & " のセットアップが完了しました。" & vbCrLf & vbCrLf & _
           "「F_メニュー」を開いて運用を開始してください。", _
           vbInformation, APP_NAME
    Exit Sub

Fail:
    MsgBox "セットアップ中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

'------------------------------------------------------------------------------
' テーブル
'------------------------------------------------------------------------------
Public Sub CreateTables()
    Echo_ "テーブルを作成しています..."

    ' --- マスタ (毎回作り直す) ---
    DropTable "M_区分"
    DropTable "M_製品"
    DropTable "M_ブロック"
    DropTable "M_集計列"
    DropTable "M_業務項目"
    DropTable "M_担当者"

    ExecDDL "CREATE TABLE [M_集計列] (" & _
        "[集計列ID] LONG NOT NULL CONSTRAINT [PK_集計列] PRIMARY KEY," & _
        "[集計列名] TEXT(50) NOT NULL," & _
        "[表示順] LONG NOT NULL)"

    ExecDDL "CREATE TABLE [M_ブロック] (" & _
        "[ブロックID] LONG NOT NULL CONSTRAINT [PK_ブロック] PRIMARY KEY," & _
        "[ブロック名] TEXT(50) NOT NULL," & _
        "[製品別] BIT NOT NULL," & _
        "[表示順] LONG NOT NULL)"

    ExecDDL "CREATE TABLE [M_製品] (" & _
        "[製品ID] LONG NOT NULL CONSTRAINT [PK_製品] PRIMARY KEY," & _
        "[ブロックID] LONG NOT NULL," & _
        "[製品名] TEXT(100) NOT NULL," & _
        "[適用開始日] DATETIME," & _
        "[適用終了日] DATETIME," & _
        "[表示順] LONG NOT NULL," & _
        "[有効] BIT NOT NULL)"

    ' 区分ごとに、集計表のどの列に積むか (集計列ID) を必ず持たせる。
    ' 現行 Excel では Sheet2 の 2 行目に埋め込まれていた情報。
    ExecDDL "CREATE TABLE [M_区分] (" & _
        "[区分ID] LONG NOT NULL CONSTRAINT [PK_区分] PRIMARY KEY," & _
        "[ブロックID] LONG NOT NULL," & _
        "[区分名] TEXT(100) NOT NULL," & _
        "[集計列ID] LONG NOT NULL," & _
        "[内訳区分] TEXT(20)," & _
        "[表示順] LONG NOT NULL," & _
        "[有効] BIT NOT NULL," & _
        "[旧転記名] TEXT(100)," & _
        "CONSTRAINT [FK_区分_ブロック] FOREIGN KEY ([ブロックID]) REFERENCES [M_ブロック]([ブロックID])," & _
        "CONSTRAINT [FK_区分_集計列] FOREIGN KEY ([集計列ID]) REFERENCES [M_集計列]([集計列ID]))"

    ' 職員の入れ替わりが多いので、担当者は「削除」せず在籍期間で管理する。
    ' 過去データは退職後も正しい担当者に紐づいたまま残り、入力画面の候補には出なくなる。
    ' 担当者コードは席番号なので再利用されうる。よって UNIQUE は張らない。
    ExecDDL "CREATE TABLE [M_担当者] (" & _
        "[担当者ID] LONG NOT NULL CONSTRAINT [PK_担当者] PRIMARY KEY," & _
        "[担当者コード] TEXT(10) NOT NULL," & _
        "[姓] TEXT(50) NOT NULL," & _
        "[名] TEXT(50)," & _
        "[氏名] TEXT(100) NOT NULL," & _
        "[カナ] TEXT(100)," & _
        "[職員区分] TEXT(10) NOT NULL," & _
        "[在籍開始日] DATETIME," & _
        "[在籍終了日] DATETIME," & _
        "[表示順] LONG NOT NULL," & _
        "[有効] BIT NOT NULL," & _
        "[備考] MEMO)"
    ExecDDL "CREATE INDEX [IX_担当者_コード] ON [M_担当者] ([担当者コード])"

    ExecDDL "CREATE TABLE [M_業務項目] (" & _
        "[業務項目ID] LONG NOT NULL CONSTRAINT [PK_業務項目] PRIMARY KEY," & _
        "[番号] TEXT(4) NOT NULL," & _
        "[項目名] TEXT(100) NOT NULL," & _
        "[帳票表示名] TEXT(100) NOT NULL," & _
        "[表示順] LONG NOT NULL," & _
        "[有効] BIT NOT NULL)"

    ' --- 実績 (既定では作り直さない) ---
    If DROP_DATA_TABLES Then
        DropTable "T_受電"
        DropTable "T_業務実績"
        DropTable "T_出勤"
        DropTable "T_日報"
    End If

    If Not TableExists("T_日報") Then
        ExecDDL "CREATE TABLE [T_日報] (" & _
            "[対象日] DATETIME NOT NULL CONSTRAINT [PK_日報] PRIMARY KEY," & _
            "[回線数] LONG," & _
            "[特記事項] MEMO," & _
            "[職員代替案件] MEMO," & _
            "[要望] MEMO," & _
            "[状態] TEXT(10) NOT NULL," & _
            "[確定日時] DATETIME," & _
            "[更新日時] DATETIME)"
    End If

    If Not TableExists("T_出勤") Then
        ExecDDL "CREATE TABLE [T_出勤] (" & _
            "[出勤ID] COUNTER NOT NULL CONSTRAINT [PK_出勤] PRIMARY KEY," & _
            "[対象日] DATETIME NOT NULL," & _
            "[担当者ID] LONG NOT NULL," & _
            "[勤務時間] TEXT(50)," & _
            "[備考] TEXT(255)," & _
            "CONSTRAINT [UQ_出勤] UNIQUE ([対象日],[担当者ID]))"
    End If

    ' 1 行 = 1 (日付, 担当者, 区分, 製品) の件数。
    ' 製品を伴わない区分は 製品ID = 0 (「製品指定なし」) を使う。
    ' NULL を使わないので UNIQUE 制約が確実に効き、二重計上が起きない。
    If Not TableExists("T_受電") Then
        ExecDDL "CREATE TABLE [T_受電] (" & _
            "[受電ID] COUNTER NOT NULL CONSTRAINT [PK_受電] PRIMARY KEY," & _
            "[対象日] DATETIME NOT NULL," & _
            "[担当者ID] LONG NOT NULL," & _
            "[区分ID] LONG NOT NULL," & _
            "[製品ID] LONG NOT NULL," & _
            "[件数] LONG NOT NULL," & _
            "[備考2] TEXT(255)," & _
            "[備考3] TEXT(255)," & _
            "[登録日時] DATETIME," & _
            "[更新日時] DATETIME," & _
            "[登録者] TEXT(100)," & _
            "CONSTRAINT [UQ_受電] UNIQUE ([対象日],[担当者ID],[区分ID],[製品ID]))"
        ExecDDL "CREATE INDEX [IX_受電_日付] ON [T_受電] ([対象日])"
    End If

    If Not TableExists("T_業務実績") Then
        ExecDDL "CREATE TABLE [T_業務実績] (" & _
            "[実績ID] COUNTER NOT NULL CONSTRAINT [PK_業務実績] PRIMARY KEY," & _
            "[対象日] DATETIME NOT NULL," & _
            "[担当者ID] LONG NOT NULL," & _
            "[業務項目ID] LONG NOT NULL," & _
            "[件数] LONG NOT NULL," & _
            "CONSTRAINT [UQ_業務実績] UNIQUE ([対象日],[担当者ID],[業務項目ID]))"
    End If

    EnsureImportLog

    ' 実績→マスタの参照整合性。マスタを作り直した後に張り直す。
    AddRelation "R_受電_担当者", "M_担当者", "T_受電", "担当者ID"
    AddRelation "R_受電_区分", "M_区分", "T_受電", "区分ID"
    AddRelation "R_受電_製品", "M_製品", "T_受電", "製品ID"
    AddRelation "R_出勤_担当者", "M_担当者", "T_出勤", "担当者ID"
    AddRelation "R_業務実績_担当者", "M_担当者", "T_業務実績", "担当者ID"
    AddRelation "R_業務実績_項目", "M_業務項目", "T_業務実績", "業務項目ID"

    Echo_ "  テーブル作成 完了"
End Sub

'------------------------------------------------------------------------------
' 起動時オプション
'------------------------------------------------------------------------------
Public Sub Setup_AppOptions()
    On Error Resume Next
    SetDbProperty "AppTitle", dbText, APP_NAME
    SetDbProperty "StartUpForm", dbText, "F_メニュー"
    SetDbProperty "AllowSpecialKeys", dbBoolean, True
    SetDbProperty "StartUpShowDBWindow", dbBoolean, True
    Application.RefreshTitleBar
End Sub

'------------------------------------------------------------------------------
' 共通ヘルパ
'------------------------------------------------------------------------------
Public Sub ExecDDL(ByVal sql As String)
    CurrentDb.Execute sql, dbFailOnError
End Sub

Public Sub ExecSQL(ByVal sql As String)
    CurrentDb.Execute sql, dbFailOnError
End Sub

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim td As Object
    On Error Resume Next
    Set td = CurrentDb.TableDefs(tableName)
    TableExists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub DropTable(ByVal tableName As String)
    If Not TableExists(tableName) Then Exit Sub
    DropRelationsOf tableName
    On Error Resume Next
    CurrentDb.Execute "DROP TABLE [" & tableName & "]", dbFailOnError
    If Err.Number <> 0 Then
        Err.Clear
        DoCmd.DeleteObject acTable, tableName
    End If
    CurrentDb.TableDefs.Refresh
End Sub

' 指定テーブルが絡むリレーションを外す。これをしないと DROP TABLE が失敗する。
Private Sub DropRelationsOf(ByVal tableName As String)
    Dim db As DAO.Database, i As Long
    Set db = CurrentDb
    For i = db.Relations.Count - 1 To 0 Step -1
        With db.Relations(i)
            If .Table = tableName Or .ForeignTable = tableName Then
                db.Relations.Delete .Name
            End If
        End With
    Next i
    db.Relations.Refresh
End Sub

Public Sub AddRelation(ByVal relName As String, ByVal parentTable As String, _
                       ByVal childTable As String, ByVal keyField As String)
    Dim db As DAO.Database, rel As DAO.Relation
    Set db = CurrentDb

    On Error Resume Next
    db.Relations.Delete relName
    Err.Clear
    On Error GoTo Fail

    Set rel = db.CreateRelation(relName, parentTable, childTable, dbRelationDontEnforce)
    rel.Fields.Append rel.CreateField(keyField)
    rel.Fields(keyField).ForeignName = keyField
    ' まず参照整合性ありで試し、既存データが合わない場合は整合性なしで張る
    rel.Attributes = 0
    db.Relations.Append rel
    db.Relations.Refresh
    Exit Sub

Fail:
    Err.Clear
    On Error Resume Next
    Set rel = db.CreateRelation(relName, parentTable, childTable, dbRelationDontEnforce)
    rel.Fields.Append rel.CreateField(keyField)
    rel.Fields(keyField).ForeignName = keyField
    db.Relations.Append rel
    db.Relations.Refresh
    Echo_ "  ! リレーション " & relName & " は参照整合性なしで作成しました"
End Sub

Public Function QueryExists(ByVal queryName As String) As Boolean
    Dim qd As Object
    On Error Resume Next
    Set qd = CurrentDb.QueryDefs(queryName)
    QueryExists = (Err.Number = 0)
    Err.Clear
End Function

Public Sub SaveQuery(ByVal queryName As String, ByVal sql As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    If QueryExists(queryName) Then db.QueryDefs.Delete queryName
    db.CreateQueryDef queryName, sql
    db.QueryDefs.Refresh
End Sub

Public Sub DeleteObjectIfExists(ByVal objType As AcObjectType, ByVal objName As String)
    On Error Resume Next
    DoCmd.Close objType, objName, acSaveNo
    Err.Clear
    DoCmd.DeleteObject objType, objName
    Err.Clear
End Sub

Private Sub SetDbProperty(ByVal propName As String, ByVal propType As Long, ByVal propValue As Variant)
    Dim db As DAO.Database, prp As DAO.Property
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(propName).Value = propValue
    If Err.Number <> 0 Then
        Err.Clear
        Set prp = db.CreateProperty(propName, propType, propValue)
        db.Properties.Append prp
    End If
End Sub

Public Sub Echo_(ByVal msg As String)
    Debug.Print Format(Now, "hh:nn:ss") & "  " & msg
    SysCmd acSysCmdSetStatus, msg
    DoEvents
End Sub

' SQL に埋め込む文字列をエスケープする
Public Function Q(ByVal s As Variant) As String
    If IsNull(s) Then
        Q = "Null"
    Else
        Q = "'" & Replace(CStr(s), "'", "''") & "'"
    End If
End Function

' SQL に埋め込む日付リテラル (Access は #m/d/yyyy# 形式)
Public Function D(ByVal v As Variant) As String
    If IsNull(v) Then
        D = "Null"
    Else
        D = "#" & Format(CDate(v), "mm\/dd\/yyyy") & "#"
    End If
End Function
