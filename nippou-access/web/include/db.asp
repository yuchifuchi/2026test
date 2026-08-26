<%
' =============================================================================
'  共通: データベース接続とヘルパ
'
'  DB は Access (.accdb) をそのまま使う。デスクトップ版 Access アプリと
'  同じファイルを参照するので、どちらから入力しても同じデータになる。
'
'  前提:
'    - IIS に「Microsoft Access Database Engine 2016 Redistributable」を導入
'    - アプリケーションプールのビット数 (32/64) を ACE のビット数と合わせる
'    - .accdb を置くフォルダに、アプリケーションプール ID の変更権限を付与
'      (Jet/ACE はロックファイル .laccdb を同じフォルダに作るため)
' =============================================================================

' --- 設定 --------------------------------------------------------------------
Const DB_PATH = "D:\nippou\data\日報集計_be.accdb"   ' 環境に合わせて変更する
Const APP_NAME = "電話応対日報 集計システム"
Const APP_VERSION = "1.0 (Web モックアップ)"

Dim gConn

Sub OpenDb()
    If IsObject(gConn) Then
        If Not gConn Is Nothing Then Exit Sub
    End If
    Set gConn = Server.CreateObject("ADODB.Connection")
    gConn.CursorLocation = 3            ' adUseClient
    gConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & _
               ";Persist Security Info=False;"
End Sub

Sub CloseDb()
    On Error Resume Next
    If Not IsEmpty(gConn) Then
        If Not gConn Is Nothing Then gConn.Close
        Set gConn = Nothing
    End If
End Sub

' -----------------------------------------------------------------------------
'  すべての SQL はパラメータ経由で実行する。
'  画面から来た値を文字列連結で SQL に埋めることは絶対にしない。
' -----------------------------------------------------------------------------
Function DbQuery(sql, params)
    Dim cmd, i
    OpenDb
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = gConn
    cmd.CommandType = 1                 ' adCmdText
    cmd.CommandText = sql
    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append MakeParam(cmd, "p" & i, params(i))
        Next
    End If
    Set DbQuery = cmd.Execute()
End Function

Function DbExec(sql, params)
    Dim cmd, i, affected
    OpenDb
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = gConn
    cmd.CommandType = 1
    cmd.CommandText = sql
    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append MakeParam(cmd, "p" & i, params(i))
        Next
    End If
    cmd.Execute affected, , 128         ' adExecuteNoRecords
    DbExec = affected
End Function

Function MakeParam(cmd, name, value)
    ' adInteger=3 adDate=7 adVarWChar=202 adLongVarWChar=203
    Dim s
    If IsNull(value) Or IsEmpty(value) Then
        Set MakeParam = cmd.CreateParameter(name, 202, 1, 255, Null)
    ElseIf VarType(value) = 7 Then                     ' vbDate
        Set MakeParam = cmd.CreateParameter(name, 7, 1, , CDate(value))
    ElseIf VarType(value) = 2 Or VarType(value) = 3 Then  ' vbInteger / vbLong
        Set MakeParam = cmd.CreateParameter(name, 3, 1, , CLng(value))
    Else
        s = CStr(value)
        If Len(s) = 0 Then
            Set MakeParam = cmd.CreateParameter(name, 202, 1, 255, Null)
        ElseIf Len(s) > 255 Then
            ' メモ欄 (特記事項など) は長文になるので adLongVarWChar で渡す
            Set MakeParam = cmd.CreateParameter(name, 203, 1, Len(s), s)
        Else
            Set MakeParam = cmd.CreateParameter(name, 202, 1, 255, s)
        End If
    End If
End Function

' 単一値の取得
Function DbScalar(sql, params, fallback)
    Dim rs
    Set rs = DbQuery(sql, params)
    If rs.EOF Then
        DbScalar = fallback
    ElseIf IsNull(rs.Fields(0).Value) Then
        DbScalar = fallback
    Else
        DbScalar = rs.Fields(0).Value
    End If
    rs.Close
End Function

' -----------------------------------------------------------------------------
'  入力値の取り出し
' -----------------------------------------------------------------------------
Function ParamDate(name, fallback)
    Dim v : v = Trim(Request(name) & "")
    If Len(v) = 0 Then
        ParamDate = fallback
    ElseIf IsDate(v) Then
        ParamDate = CDate(v)
    Else
        ParamDate = fallback
    End If
End Function

Function ParamLong(name, fallback)
    Dim v : v = Trim(Request(name) & "")
    If Len(v) = 0 Or Not IsNumeric(v) Then
        ParamLong = fallback
    Else
        ParamLong = CLng(v)
    End If
End Function

Function ParamText(name)
    ParamText = Trim(Request(name) & "")
End Function

' -----------------------------------------------------------------------------
'  出力
' -----------------------------------------------------------------------------
Function H(v)
    Dim s : s = "" & v
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    H = s
End Function

Function Ymd(v)
    If IsNull(v) Or Not IsDate(v) Then
        Ymd = ""
    Else
        Ymd = Year(v) & "-" & Right("0" & Month(v), 2) & "-" & Right("0" & Day(v), 2)
    End If
End Function

Function YmdSlash(v)
    If IsNull(v) Or Not IsDate(v) Then
        YmdSlash = ""
    Else
        YmdSlash = Year(v) & "/" & Right("0" & Month(v), 2) & "/" & Right("0" & Day(v), 2)
    End If
End Function

Function Youbi(v)
    Dim w : w = Array("日","月","火","水","木","金","土")
    Youbi = w(Weekday(v) - 1)
End Function

' 現行 印刷用シート W2 と同じ和暦表記
Function Wareki(v)
    Dim g, y
    If Not IsDate(v) Then Wareki = "" : Exit Function
    If CDate(v) >= CDate("2019/05/01") Then
        g = "令和" : y = Year(v) - 2018
    ElseIf CDate(v) >= CDate("1989/01/08") Then
        g = "平成" : y = Year(v) - 1988
    Else
        g = "昭和" : y = Year(v) - 1925
    End If
    Wareki = g & "　" & y & "年　" & Month(v) & "月　" & Day(v) & "日（" & Youbi(v) & "）"
End Function

' 今日が属する週の月曜
Function MondayOf(d)
    Dim w : w = Weekday(d, 2)          ' 月曜=1
    MondayOf = DateAdd("d", 1 - w, d)
End Function

' -----------------------------------------------------------------------------
'  業務ロジック (デスクトップ版 modApp と同じ規則)
' -----------------------------------------------------------------------------
Sub Ensure日報(dt)
    If DbScalar("SELECT Count(*) FROM [T_日報] WHERE [対象日]=?", Array(dt), 0) > 0 Then Exit Sub
    DbExec "INSERT INTO [T_日報] ([対象日],[状態],[更新日時]) VALUES (?,'入力中',Now())", Array(dt)
End Sub

Sub Ensure出勤(dt, opId)
    If DbScalar("SELECT Count(*) FROM [T_出勤] WHERE [対象日]=? AND [担当者ID]=?", _
                Array(dt, opId), 0) > 0 Then Exit Sub
    DbExec "INSERT INTO [T_出勤] ([対象日],[担当者ID]) VALUES (?,?)", Array(dt, opId)
End Sub

' 在籍中の担当者だけを返す。退職者は候補から消えるが過去データは残る。
Function OperatorsOn(dt)
    Set OperatorsOn = DbQuery( _
        "SELECT [担当者ID],[担当者コード],[氏名],[職員区分] FROM [M_担当者] " & _
        "WHERE [有効]=True " & _
        "  AND ([在籍開始日] Is Null OR [在籍開始日] <= ?) " & _
        "  AND ([在籍終了日] Is Null OR [在籍終了日] >= ?) " & _
        "ORDER BY [表示順]", Array(dt, dt))
End Function
%>
