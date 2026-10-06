<%
' ============================================================
'  PDF を作る（サーバーにある Microsoft Edge を画面なしで動かし、HTML を PDF にする）
'
'  ・作った HTML と PDF は作業フォルダ（PDF_WORK_DIR、空ならデータベースと同じフォルダの pdfwork）の
'    job_日時_番号 フォルダに置き、PDF_KEEP_HOURS 時間を過ぎたものは次に PDF を作るときに消す。
'  ・Edge には、利用者のプロファイルではなく、作業ごとの専用フォルダ（--user-data-dir）を使わせる。
'    IIS の中で動く Edge には「ふだんのユーザー」がいないため。
'  ・--no-sandbox を付けるのは、IIS の中（画面の無い場所）では Edge の安全装置が起動できない
'    ことがあるため。読み込むのはこのシステムが作った HTML だけで、外のページは開かない。
' ============================================================

' CSS などの UTF-8 のファイルを読む。
' FileSystemObject.OpenTextFile(..., -1) の -1 は UTF-16 の意味なので、UTF-8 のファイルを読むと文字化けする（7-5）。
Function ReadTextUtf8(path)
    Dim st
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextUtf8 = st.ReadText
    st.Close
End Function

' UTF-8 で書く（ADODB.Stream は先頭に BOM を付けるが、HTML なので問題ない）
Function PdfWriteUtf8(path, text)
    Dim st, errNo
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    On Error Resume Next
    st.SaveToFile path, 2
    errNo = Err.Number
    On Error GoTo 0
    st.Close
    PdfWriteUtf8 = (errNo = 0)
End Function

Function PdfWorkDir()
    Dim p, i, sh
    If PDF_WORK_DIR <> "" Then
        PdfWorkDir = PDF_WORK_DIR
        Exit Function
    End If
    p = DbResolvePath()
    If Left(p, 2) = "\\" Then
        ' データベースが共有フォルダにあるときは、サーバーの一時フォルダを使う
        ' （Edge は IIS のアプリケーション プールの名前で動き、共有フォルダに入れないことが多いため）
        Set sh = Server.CreateObject("WScript.Shell")
        PdfWorkDir = sh.ExpandEnvironmentStrings("%TEMP%") & "\nippou_pdf"
        Exit Function
    End If
    i = InStrRev(p, "\")
    PdfWorkDir = Left(p, i) & "pdfwork"
End Function

' フォルダが無ければ作る。作れなければ msg に理由を入れて False。
Function PdfEnsureFolder(dir, ByRef msg)
    Dim fso, errNo, errDesc
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    PdfEnsureFolder = True
    If fso.FolderExists(dir) Then Exit Function
    On Error Resume Next
    fso.CreateFolder dir
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        msg = "PDF の作業フォルダを作れませんでした: " & dir & "（" & errDesc & "）。このフォルダの親に書き込みの許可があるか確かめてください。"
        PdfEnsureFolder = False
    End If
End Function

' 古い作業フォルダを消す。消せないもの（まだ使われているもの）は次の回に回す。
Sub PdfCleanup(dir)
    Dim fso, f, paths(), n, i, limit
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then Exit Sub
    limit = DateAdd("h", -PDF_KEEP_HOURS, Now())
    n = -1
    ReDim paths(-1)
    For Each f In fso.GetFolder(dir).SubFolders
        If Left(f.Name, 4) = "job_" Then
            If f.DateCreated < limit Then
                n = n + 1
                ReDim Preserve paths(n)
                paths(n) = f.Path
            End If
        End If
    Next
    For i = 0 To n
        PdfTryDeleteFolder fso, paths(i)
    Next
End Sub

Sub PdfTryDeleteFolder(fso, path)
    On Error Resume Next
    fso.DeleteFolder path, True
    On Error GoTo 0
End Sub

' file:/// の URL にする（日本語や空白は UTF-8 の %xx にする）
Function PdfFileUrl(path)
    Dim parts, i, s
    parts = Split(path, "\")
    s = "file:///" & parts(0)
    For i = 1 To UBound(parts)
        s = s & "/" & Replace(Server.URLEncode(parts(i)), "+", "%20")
    Next
    PdfFileUrl = s
End Function

Function PdfElapsed(t0)
    Dim e
    e = Timer() - t0
    If e < 0 Then e = e + 86400
    PdfElapsed = e
End Function

Function PdfStamp()
    Dim n
    n = Now()
    PdfStamp = Year(n) & Pad2(Month(n)) & Pad2(Day(n)) & "_" & Pad2(Hour(n)) & Pad2(Minute(n)) & Pad2(Second(n))
End Function

' html を PDF にする。できたら True と pdfPath。できなければ False と msg。
Function MakePdf(html, ByRef pdfPath, ByRef msg)
    Dim fso, sh, ex, dir, job, htmlPath, cmd, t0, errNo, errDesc, code
    MakePdf = False
    msg = ""
    pdfPath = ""
    If LCase(PDF_ENGINE) <> "edge" Then
        msg = "config.asp の PDF_ENGINE が ""edge"" ではありません（今は edge だけ使えます）。"
        Exit Function
    End If
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(PDF_EDGE_EXE) Then
        msg = "Microsoft Edge が見つかりません: " & PDF_EDGE_EXE & "（config.asp の PDF_EDGE_EXE を確かめてください）"
        Exit Function
    End If
    dir = PdfWorkDir()
    If Not PdfEnsureFolder(dir, msg) Then Exit Function
    PdfCleanup dir
    Randomize
    job = dir & "\job_" & PdfStamp() & "_" & CStr(Int(Rnd() * 900000) + 100000)
    On Error Resume Next
    fso.CreateFolder job
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        msg = "PDF の作業フォルダを作れませんでした: " & job & "（" & errDesc & "）"
        Exit Function
    End If
    htmlPath = job & "\sheet.html"
    pdfPath = job & "\sheet.pdf"
    If Not PdfWriteUtf8(htmlPath, html) Then
        msg = "PDF の元になる HTML を書き込めませんでした: " & htmlPath
        Exit Function
    End If
    cmd = """" & PDF_EDGE_EXE & """ --headless=new --disable-gpu --no-pdf-header-footer --no-first-run" & _
          " --no-default-browser-check --disable-extensions --no-sandbox --log-level=3" & _
          " --user-data-dir=""" & job & "\profile""" & _
          " --print-to-pdf=""" & pdfPath & """ """ & PdfFileUrl(htmlPath) & """"
    Set sh = Server.CreateObject("WScript.Shell")
    On Error Resume Next
    Set ex = sh.Exec(cmd)
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        msg = "Microsoft Edge を起動できませんでした（" & errDesc & "）"
        Exit Function
    End If
    t0 = Timer()
    Do While ex.Status = 0
        If PdfElapsed(t0) > PDF_TIMEOUT Then
            PdfKill sh, ex
            msg = "Microsoft Edge が " & PDF_TIMEOUT & " 秒たっても終わりませんでした。もう一度試して、同じなら setup_check.asp を開いてください。"
            Exit Function
        End If
        ' VBScript には「待つ」命令が無いので、ping で 1 秒待つ
        sh.Run "%ComSpec% /c ping -n 2 127.0.0.1 > nul", 0, True
    Loop
    code = ex.ExitCode
    If Not fso.FileExists(pdfPath) Then
        msg = "PDF ができませんでした（Edge の終了コード " & code & "）。作業フォルダ: " & job
        Exit Function
    End If
    If fso.GetFile(pdfPath).Size = 0 Then
        msg = "PDF が空でした。作業フォルダ: " & job
        Exit Function
    End If
    PdfTryDeleteFolder fso, job & "\profile"
    MakePdf = True
End Function

' 時間切れの Edge を止める（Edge は子の処理を何本も起こすので、まとめて止める）
Sub PdfKill(sh, ex)
    On Error Resume Next
    sh.Run "%ComSpec% /c taskkill /F /T /PID " & ex.ProcessID & " > nul", 0, True
    ex.Terminate
    On Error GoTo 0
End Sub

' 帳票の HTML を、CSS ごと 1 枚の HTML にする（Edge はこのファイルだけを開く）
Function PdfDocument(bodyHtml)
    Dim css
    css = ReadTextUtf8(Server.MapPath("css/style.css"))
    PdfDocument = "<!DOCTYPE html>" & vbCrLf & "<html lang=""ja""><head><meta charset=""utf-8""><title>電話応対報告書日報集計表</title><style>" & _
                  css & "</style></head><body class=""pdf"">" & bodyHtml & "</body></html>"
End Function

' PDF を画面に送る。ページはここで止まる。
' On Error Resume Next が生きている所で呼ばないこと（止まらなくなる。手順書 7-2）。
Sub SendPdf(pdfPath, asciiName, jaName)
    Dim st, data
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile pdfPath
    data = st.Read
    st.Close
    Response.Clear
    Response.ContentType = "application/pdf"
    Response.AddHeader "Content-Disposition", "inline; filename=""" & asciiName & """; filename*=UTF-8''" & Server.URLEncode(jaName)
    Response.BinaryWrite data
    DbClose
    Response.End
End Sub
%>
