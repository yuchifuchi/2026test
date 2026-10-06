<%
' ============================================================
'  データベースの出入り口
'
'  画面からデータベースを触るときは、必ずここの関数を通す:
'    DbQuery(SQL, Array(値, ...))          … 行の配列（1 行 = 列名で引ける Dictionary）
'    DbScalar(SQL, Array(値, ...), 既定値)  … 1 行目の 1 列目
'    DbExec(SQL, Array(値, ...))           … INSERT / UPDATE / DELETE。変わった行の数を返す
'    DbBegin / DbCommit                     … まとめて保存するとき
'  SQL は include/sql.asp の関数から渡し、値は必ず ? で渡す（文字列でつなげない）。
'
'  失敗したら DbFatal が「何が起きたか」を画面に出して止める（HTTP 200 で返す。
'  500 で返すと IIS が本文を捨てて既定のエラー画面に差し替えるため。手順書 7-3）。
' ============================================================

' ADO の定数（VBScript には最初から入っていないので、使うものだけここで決める）
Const adInteger = 3
Const adDouble = 5
Const adDate = 7
Const adBoolean = 11
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adParamInput = 1
Const adCmdText = 1

Dim g_conn      ' このページで使う接続（1 ページ 1 本）
Dim g_inTrans   ' まとめて保存している途中か

' DB_PATH を実際の場所に直す。利用者がいちばん書き間違えるところなので 3 通りを受け付ける。
Function DbResolvePath()
    Dim p, root, mapped, errNo, errDesc
    p = Trim(DB_PATH)
    If p = "" Then
        DbFatal "設定の誤り", "config.asp の DB_PATH が空です。データベースの場所を書いてください。"
    End If
    If Left(p, 2) = "\\" Then
        ' \\サーバー名\共有\... （共有フォルダ）はそのまま使う
        DbResolvePath = p
        Exit Function
    End If
    If Mid(p, 2, 1) = ":" Then
        ' D:\... （ドライブ文字）はそのまま使う
        DbResolvePath = p
        Exit Function
    End If
    ' それ以外はサイトの中の場所とみなして Server.MapPath で直す。
    ' ファイル名だけ（や data/xxx.accdb）なら、このシステムの data フォルダの中とみなす。
    p = Replace(p, "\", "/")
    If Left(p, 1) <> "/" Then
        root = DbAppRoot()
        If InStr(p, "/") = 0 Then
            p = root & "/data/" & p
        Else
            p = root & "/" & p
        End If
    End If
    On Error Resume Next
    mapped = Server.MapPath(p)
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        DbFatal "設定の誤り", "config.asp の DB_PATH（" & DB_PATH & "）を場所に直せませんでした。" & vbCrLf & _
            "「..」は使えません。/nippou/data/日報集計_be.accdb のように / から書くか、D:\ や \\ で始まる場所を書いてください。" & vbCrLf & errDesc
    End If
    DbResolvePath = mapped
End Function

' このシステムの入口（/nippou など）。ページの URL から 2 段上がった所。
Function DbAppRoot()
    Dim u, i
    u = Request.ServerVariables("SCRIPT_NAME")
    i = InStrRev(u, "/")
    If i > 0 Then u = Left(u, i - 1)
    i = InStrRev(u, "/")
    If i > 0 Then u = Left(u, i - 1) Else u = ""
    DbAppRoot = u
End Function

Sub DbOpen()
    Dim c, path, errNo, errDesc
    If IsObject(g_conn) Then Exit Sub
    path = DbResolvePath()
    Set c = Server.CreateObject("ADODB.Connection")
    On Error Resume Next
    c.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";"
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo = -2146824582 Then
        ' 12.0 が無いときは 16.0 も試す（Access データベース エンジン 2016 だけが入っている場合）
        On Error Resume Next
        c.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & path & ";"
        errNo = Err.Number
        errDesc = Err.Description
        On Error GoTo 0
    End If
    If errNo <> 0 Then
        DbFatal "データベースに接続できません", DbErrDetail(errNo, errDesc, "") & vbCrLf & "場所: " & path
    End If
    Set g_conn = c
    g_inTrans = False
End Sub

Sub DbClose()
    If IsObject(g_conn) Then
        If g_conn.State = 1 Then g_conn.Close
    End If
End Sub

' SQL の中の ? の数（'...' と [...] の中は数えない）
Function DbCountMarks(sql)
    Dim i, c, n, inQ, inB
    n = 0
    inQ = False
    inB = False
    For i = 1 To Len(sql)
        c = Mid(sql, i, 1)
        If inQ Then
            If c = "'" Then inQ = False
        ElseIf inB Then
            If c = "]" Then inB = False
        ElseIf c = "'" Then
            inQ = True
        ElseIf c = "[" Then
            inB = True
        ElseIf c = "?" Then
            n = n + 1
        End If
    Next
    DbCountMarks = n
End Function

' 値の種類（数・日付・文字など）に合わせて、ADO に渡す「値の入れ物」を作る
Function DbParam(cmd, v)
    Dim t, n
    t = VarType(v)
    Select Case t
        Case vbEmpty, vbNull
            Set DbParam = cmd.CreateParameter("", adVarWChar, adParamInput, 1, Null)
        Case vbInteger, vbLong, vbByte
            Set DbParam = cmd.CreateParameter("", adInteger, adParamInput, 4, CLng(v))
        Case vbSingle, vbDouble, vbCurrency
            Set DbParam = cmd.CreateParameter("", adDouble, adParamInput, 8, CDbl(v))
        Case vbDate
            Set DbParam = cmd.CreateParameter("", adDate, adParamInput, 8, v)
        Case vbBoolean
            Set DbParam = cmd.CreateParameter("", adBoolean, adParamInput, 2, v)
        Case vbString
            n = Len(v)
            If n = 0 Then
                ' 空欄は Null で保存する。空文字と Null が混ざると「未入力」の調べ方が 2 通りになるため。
                Set DbParam = cmd.CreateParameter("", adVarWChar, adParamInput, 1, Null)
            ElseIf n <= 255 Then
                Set DbParam = cmd.CreateParameter("", adVarWChar, adParamInput, n, v)
            Else
                Set DbParam = cmd.CreateParameter("", adLongVarWChar, adParamInput, n, v)
            End If
        Case Else
            DbFatal "プログラムの誤り", "データベースに渡せない種類の値です（VarType=" & t & "）"
    End Select
End Function

Function DbCommand(sql, args)
    Dim cmd, i, n
    If Not IsArray(args) Then
        DbFatal "プログラムの誤り", "値は Array(...) で渡してください。" & vbCrLf & sql
    End If
    n = DbCountMarks(sql)
    If UBound(args) + 1 <> n Then
        DbFatal "プログラムの誤り", "SQL の ? の数（" & n & "）と、渡した値の数（" & (UBound(args) + 1) & "）が合いません。" & vbCrLf & sql
    End If
    DbOpen
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = g_conn
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    For i = 0 To UBound(args)
        cmd.Parameters.Append DbParam(cmd, args(i))
    Next
    Set DbCommand = cmd
End Function

' 結果を「行の配列」で返す。各行は列名で引ける Dictionary。
' Recordset のまま渡さないのは、rs("列") を Dictionary などに入れると
' 値ではなく Field オブジェクトが入り、次の行へ進むと中身が変わってしまうため。
Function DbQuery(sql, args)
    Dim cmd, rs, rows, n, d, i, errNo, errDesc
    Set cmd = DbCommand(sql, args)
    On Error Resume Next
    Set rs = cmd.Execute
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        DbFatal "データの読み込み", DbErrDetail(errNo, errDesc, sql)
    End If
    ReDim rows(-1)
    n = -1
    Do While Not rs.EOF
        Set d = Server.CreateObject("Scripting.Dictionary")
        For i = 0 To rs.Fields.Count - 1
            d.Add rs.Fields(i).Name, rs.Fields(i).Value
        Next
        n = n + 1
        ReDim Preserve rows(n)
        Set rows(n) = d
        rs.MoveNext
    Loop
    rs.Close
    DbQuery = rows
End Function

' 1 行目の 1 列目。行が無いか Null なら dflt。
Function DbScalar(sql, args, dflt)
    Dim rows, its
    rows = DbQuery(sql, args)
    DbScalar = dflt
    If UBound(rows) >= 0 Then
        its = rows(0).Items
        If Not IsNull(its(0)) Then DbScalar = its(0)
    End If
End Function

Function DbExec(sql, args)
    Dim cmd, n, errNo, errDesc
    Set cmd = DbCommand(sql, args)
    n = 0
    On Error Resume Next
    cmd.Execute n
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        DbFatal "データの保存", DbErrDetail(errNo, errDesc, sql)
    End If
    DbExec = n
End Function

Sub DbBegin()
    DbOpen
    g_conn.BeginTrans
    g_inTrans = True
End Sub

Sub DbCommit()
    Dim errNo, errDesc
    On Error Resume Next
    g_conn.CommitTrans
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    g_inTrans = False
    If errNo <> 0 Then
        DbFatal "データの保存（確定）", DbErrDetail(errNo, errDesc, "")
    End If
End Sub

Sub DbRollback()
    If g_inTrans Then
        g_conn.RollbackTrans
        g_inTrans = False
    End If
End Sub

' データベースのエラーを、利用者が次に何をすればよいか分かる文にする
Function DbErrDetail(errNo, errDesc, sql)
    Dim s, hint
    hint = ""
    If errNo = -2146824582 Then
        hint = "このサーバーに「Microsoft Access データベース エンジン」が入っていないか、32 ビット版と 64 ビット版が合っていません。setup_check.asp を開いて確かめてください。"
    ElseIf DbHas(errDesc, "見つかりません") Or DbHas(errDesc, "could not find") Or DbHas(errDesc, "有効なパスではありません") Or DbHas(errDesc, "not a valid path") Then
        hint = "データベースのファイルが見つかりません。config.asp の DB_PATH と、ファイルの置き場所を確かめてください。"
    ElseIf DbHas(errDesc, "更新可能") Or DbHas(errDesc, "updateable") Or DbHas(errDesc, "読み取り専用") Or DbHas(errDesc, "read-only") Or DbHas(errDesc, "アクセス許可") Or DbHas(errDesc, "permission") Or DbHas(errDesc, "開けません") Or DbHas(errDesc, "cannot open") Then
        hint = "データベースのフォルダに書き込みの許可がありません。手順書「01_設置手順」の「書き込みの許可」をもう一度行ってください。"
    ElseIf DbHas(errDesc, "ロック") Or DbHas(errDesc, "lock") Or DbHas(errDesc, "使用中") Or DbHas(errDesc, "in use") Then
        hint = "ほかの人が保存している途中でした。少し待ってから、もう一度「保存」を押してください。"
    ElseIf DbHas(errDesc, "重複") Or DbHas(errDesc, "duplicate") Then
        hint = "同じ内容が二重に登録されそうになりました（別の画面で同時に保存した可能性があります）。画面を開き直してから、もう一度入力してください。"
    End If
    s = "番号: " & Hex(errNo) & vbCrLf & "内容: " & errDesc
    If hint <> "" Then s = hint & vbCrLf & vbCrLf & s
    If sql <> "" Then s = s & vbCrLf & "SQL: " & sql
    DbErrDetail = s
End Function

Function DbHas(s, part)
    DbHas = (InStr(1, s, part, vbTextCompare) > 0)
End Function

' 致命的なエラー。画面を「何が起きたか」の説明に差し替えて、ページを止める。
' On Error GoTo 0 を Response.End の前に必ず入れる。Response.End は「捕まえられる例外」として
' 作られているので、On Error Resume Next が生きているとページが止まらない（手順書 7-2）。
' この関数を呼ぶ側も、On Error Resume Next を生かしたまま呼んではいけない（tools/vbsim/lint.py で検査）。
Sub DbFatal(what, detail)
    On Error Resume Next
    If g_inTrans Then g_conn.RollbackTrans
    g_inTrans = False
    Response.Clear
    On Error GoTo 0
    Response.Status = "200 OK"
    Response.ContentType = "text/html"
    Response.CharSet = "utf-8"
    Response.Write DbFatalHtml(what, detail)
    Response.End
End Sub

Function DbFatalHtml(what, detail)
    Dim s, url
    url = Request.ServerVariables("URL")
    If Request.ServerVariables("QUERY_STRING") <> "" Then url = url & "?" & Request.ServerVariables("QUERY_STRING")
    s = "<!DOCTYPE html>" & vbCrLf & "<html lang=""ja""><head><meta charset=""utf-8"">"
    s = s & "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">"
    s = s & "<title>うまく動きませんでした</title><link rel=""stylesheet"" href=""css/style.css""></head><body>"
    s = s & "<!-- NIPPOU-FATAL -->"
    s = s & "<div class=""band band-ng""><div class=""band-in""><span class=""app"">" & DbH(APP_NAME) & "</span></div></div>"
    s = s & "<div class=""wrap""><h1>うまく動きませんでした</h1>"
    s = s & "<p class=""msg msg-ng"">" & DbH(what) & "</p>"
    s = s & "<table class=""tw err5"">"
    s = s & "<tr><th>ファイル名</th><td>" & DbH(Request.ServerVariables("SCRIPT_NAME")) & "</td></tr>"
    s = s & "<tr><th>行番号</th><td>（データベースの処理の中）</td></tr>"
    s = s & "<tr><th>内容</th><td class=""pre"">" & DbH(detail) & "</td></tr>"
    s = s & "<tr><th>種別</th><td>" & DbH(what) & "</td></tr>"
    s = s & "<tr><th>URL</th><td>" & DbH(url) & "</td></tr>"
    s = s & "</table>"
    s = s & "<p>上の 5 行を、そのまま職員（またはシステムの担当）にお知らせください。</p>"
    s = s & "<p><a class=""btn"" href=""default.asp"">メニューへ戻る</a></p></div></body></html>"
    DbFatalHtml = s
End Function

' HTML に埋め込む文字の置き換え（layout.asp の H と同じ。db.asp だけを読み込む probe でも使えるようにここにも置く）
Function DbH(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        DbH = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    DbH = s
End Function
%>
