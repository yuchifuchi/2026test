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
'
'  クエリの SQL は sql.asp が持っている。Access の保存クエリ (Q_...) には
'  頼らない。ここから読み込んでおけば、どのページからも 1 回だけ読まれる。
' =============================================================================
%>
<!--#include file="sql.asp"-->
<%

' --- 設定 --------------------------------------------------------------------
Const DB_PATH = "D:\nippou\data\日報集計_be.accdb"   ' 環境に合わせて変更する
Const APP_NAME = "電話応対日報 集計システム"
Const APP_VERSION = "1.0 (Web モックアップ)"

Dim gConn

Sub OpenDb()
    If IsObject(gConn) Then
        If Not gConn Is Nothing Then Exit Sub
    End If

    ' つながらないときに真っ白な画面 (HTTP 500) にしない。
    ' 設置でつまずくのはほぼこの 1 行なので、原因の候補まで出す。
    On Error Resume Next
    Set gConn = Server.CreateObject("ADODB.Connection")
    If Err.Number <> 0 Then
        DbFatal "サーバーでデータベース機能を呼び出せませんでした。", Err.Description
    End If
    gConn.CursorLocation = 3            ' adUseClient
    gConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH & _
               ";Persist Security Info=False;"
    If Err.Number <> 0 Then
        DbFatal "データベースに接続できませんでした。", Err.Description
    End If
    On Error GoTo 0
End Sub

' 接続できないときの画面。ここで処理を打ち切る。
Sub DbFatal(what, detail)
    On Error Resume Next
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Write "<!doctype html><html lang=""ja""><head><meta charset=""utf-8"">" & _
        "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & _
        "<title>データベースにつながりません</title>" & _
        "<link rel=""stylesheet"" href=""css/style.css""></head><body><main>" & _
        "<h1>データベースにつながりません</h1>" & _
        "<p class=""notice err"">" & H(what) & "<br>" & _
        "入力した内容は保存されていません。少し時間をおいて開き直してください。" & _
        "直らないときは、下の内容を情報システム担当にお伝えください。</p>" & _
        "<div class=""panel""><h2 style=""margin-top:0"">担当者のかたへ</h2>" & _
        "<pre style=""white-space:pre-wrap; background:#f4f2ee; padding:12px 14px;" & _
        " border-radius:6px; font-size:13px"">" & H(detail) & vbCrLf & vbCrLf & _
        "接続先: " & H(DB_PATH) & "</pre>" & _
        "<p class=""lead"">よくある原因は次の 4 つです。</p><ul>" & _
        "<li><b>Microsoft Access Database Engine が入っていない</b>" & _
        "（「プロバイダーが見つかりません」と出ます）</li>" & _
        "<li><b>アプリケーション プールのビット数が合っていない</b>" & _
        "（64bit の ACE なら「32 ビット アプリケーションの有効化」は False）</li>" & _
        "<li><b>接続先の場所が違う</b>（include\db.asp の DB_PATH）</li>" & _
        "<li><b>フォルダに書き込めない</b>" & _
        "（.accdb を置いたフォルダに、アプリケーション プール ID の変更権限が要ります）</li>" & _
        "</ul><p class=""lead"">" & _
        "<a class=""btn"" href=""setup_check.asp"">設置チェックの画面を開く</a></p></div>" & _
        "</main></body></html>"
    Response.End
End Sub

' テキストファイルを UTF-8 として読む。
' FileSystemObject の OpenTextFile は UTF-8 を読めない (Shift_JIS か UTF-16 のみ)。
' このシステムのファイルはすべて UTF-8 なので、ADODB.Stream を使う。
Function ReadTextUtf8(path)
    Dim st
    ReadTextUtf8 = ""
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2                       ' テキスト
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    If Err.Number = 0 Then ReadTextUtf8 = st.ReadText
    st.Close
    Err.Clear
End Function

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
