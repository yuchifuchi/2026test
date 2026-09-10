<%
' =============================================================================
'  帳票を PDF にする
'
'  なぜ PDF にするのか
'      ブラウザの印刷は、既定でヘッダーとフッター (日付・URL・ページ番号) を
'      紙に入れる。これは CSS では消せず、利用者が印刷設定を変えるしかない。
'      課内限りの回覧文書に URL が印刷されるのを、設定に頼らず仕組みで防ぐ。
'
'  作った PDF はブラウザに送ったら即座に消す。フォルダには溜まらない。
'  それでも消し残った古いファイルは、次に誰かが印刷したときに掃除する。
' =============================================================================

' 設定 (変換に使うプログラム・作業用フォルダ) は config.asp にあります。


' --- ここから下は通常さわらない ----------------------------------------------

Dim gPdfError
gPdfError = ""

Function PdfWorkDir()
    Dim fso
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Len(Trim(PDF_WORK_DIR)) > 0 Then
        PdfWorkDir = PDF_WORK_DIR
    Else
        PdfWorkDir = fso.GetSpecialFolder(2).Path    ' 2 = TemporaryFolder
    End If
End Function

' 同時に押されてもぶつからないよう、毎回ちがう名前にする
Function PdfTempName(ext)
    Dim s
    s = "nippou_" & Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "") & _
        "_" & Int(Rnd() * 100000) & "_" & Timer() * 100
    s = Replace(s, ".", "")
    PdfTempName = PdfWorkDir() & "\" & s & "." & ext
End Function

' 消し残した古いファイルを片付ける。
' 日報の PDF には氏名が入るので、置きっぱなしにしない。
Sub PdfSweep()
    Dim fso, folder, f, limit
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(PdfWorkDir())
    limit = DateAdd("h", -PDF_KEEP_HOURS, Now())
    For Each f In folder.Files
        If Left(f.Name, 7) = "nippou_" Then
            If f.DateLastModified < limit Then f.Delete True
        End If
    Next
    Err.Clear
End Sub

' HTML を UTF-8 でファイルに書く
Sub PdfWriteHtml(path, html)
    Dim st
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2                       ' テキスト
    st.Charset = "utf-8"
    st.Open
    st.WriteText html
    st.SaveToFile path, 2             ' 2 = 上書き
    st.Close
End Sub

' HTML から PDF を作る。成功したら PDF のパス、失敗したら "" を返す。
' 失敗の理由は gPdfError に入れる。
Function PdfFromHtml(html)
    Dim fso, shell, inPath, outPath, profile, cmd, rc

    PdfFromHtml = ""
    gPdfError = ""
    Randomize
    Server.ScriptTimeout = PDF_TIMEOUT      ' 変換が固まったときに永久に待たない

    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set shell = Server.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        gPdfError = "サーバーで外部プログラムを呼び出せませんでした。" & _
                    "（" & Err.Description & "）"
        Exit Function
    End If
    On Error GoTo 0

    PdfSweep

    inPath = PdfTempName("html")
    outPath = PdfTempName("pdf")

    On Error Resume Next
    PdfWriteHtml inPath, html
    If Err.Number <> 0 Then
        gPdfError = "作業用フォルダに書き込めませんでした。" & vbCrLf & _
                    PdfWorkDir() & vbCrLf & _
                    "アプリケーションプール ID に「変更」権限があるか確認してください。" & _
                    "（" & Err.Description & "）"
        Exit Function
    End If
    On Error GoTo 0

    If LCase(PDF_ENGINE) = "wkhtmltopdf" Then
        If Not fso.FileExists(PDF_WKHTML_EXE) Then
            gPdfError = "wkhtmltopdf が見つかりません。" & vbCrLf & PDF_WKHTML_EXE
            fso.DeleteFile inPath, True
            Exit Function
        End If
        cmd = """" & PDF_WKHTML_EXE & """" & _
              " --page-size A4 --orientation Portrait" & _
              " --margin-top 12mm --margin-bottom 12mm" & _
              " --margin-left 12mm --margin-right 12mm" & _
              " --encoding utf-8 --disable-smart-shrinking --quiet" & _
              " """ & inPath & """ """ & outPath & """"
    Else
        If Not fso.FileExists(PDF_EDGE_EXE) Then
            gPdfError = "Microsoft Edge が見つかりません。" & vbCrLf & PDF_EDGE_EXE & vbCrLf & _
                        "場所が違う場合は include\pdf.asp の PDF_EDGE_EXE を直すか、" & _
                        "PDF_ENGINE を ""wkhtmltopdf"" に変えてください。"
            fso.DeleteFile inPath, True
            Exit Function
        End If
        ' Edge は書き込めるプロファイル置き場を要求するので、作業用フォルダを渡す
        profile = PdfTempName("prof")
        cmd = """" & PDF_EDGE_EXE & """" & _
              " --headless=new --disable-gpu --no-sandbox" & _
              " --no-pdf-header-footer" & _
              " --user-data-dir=""" & profile & """" & _
              " --print-to-pdf=""" & outPath & """" & _
              " ""file:///" & Replace(inPath, "\", "/") & """"
    End If

    On Error Resume Next
    ' 0 = 画面を出さない / True = 終わるまで待つ
    ' cmd /c は挟まない。挟むと引用符の扱いが環境で変わり、動かないことがある。
    rc = shell.Run(cmd, 0, True)
    If Err.Number <> 0 Then
        gPdfError = "変換プログラムを起動できませんでした。（" & Err.Description & "）"
        fso.DeleteFile inPath, True
        Exit Function
    End If
    On Error GoTo 0

    fso.DeleteFile inPath, True
    If Len(profile) > 0 Then
        On Error Resume Next
        fso.DeleteFolder profile, True
        Err.Clear
        On Error GoTo 0
    End If

    If Not fso.FileExists(outPath) Then
        gPdfError = "PDF が作られませんでした。（終了コード " & rc & "）" & vbCrLf & _
                    "変換プログラムの場所と、作業用フォルダの権限を確認してください。"
        Exit Function
    End If
    If fso.GetFile(outPath).Size = 0 Then
        gPdfError = "できた PDF が空でした。（終了コード " & rc & "）"
        fso.DeleteFile outPath, True
        Exit Function
    End If

    PdfFromHtml = outPath
End Function

' PDF をブラウザに送り、サーバー側のファイルを消す
Sub PdfSendAndDelete(path, downloadName)
    Dim st, fso
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1                       ' バイナリ
    st.Open
    st.LoadFromFile path

    Response.Clear
    Response.ContentType = "application/pdf"
    ' inline = ダウンロードせずブラウザの PDF ビューアで開く
    Response.AddHeader "Content-Disposition", "inline; filename=""" & downloadName & """"
    Response.AddHeader "Content-Length", CStr(st.Size)
    Response.BinaryWrite st.Read
    st.Close

    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile path, True         ' 送ったら残さない
    Err.Clear
    Response.Flush
End Sub
%>
