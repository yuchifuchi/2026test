<%
Sub PageHead(title, activeNav)
    Dim navs, i, nm, url
    Response.Write "<!doctype html>" & vbCrLf
    Response.Write "<html lang=""ja""><head><meta charset=""utf-8"">" & vbCrLf
    Response.Write "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">" & vbCrLf
    Response.Write "<title>" & H(title) & " | " & H(APP_NAME) & "</title>" & vbCrLf
    Response.Write "<link rel=""stylesheet"" href=""css/style.css""></head><body>" & vbCrLf

    Response.Write "<header class=""appbar"">" & vbCrLf
    Response.Write "  <a class=""brand"" href=""default.asp"">" & H(APP_NAME) & "</a>" & vbCrLf
    Response.Write "  <nav>"
    navs = Array("entry.asp", "受付入力", "daily.asp", "日報", _
                 "summary.asp", "集計表", "check.asp", "入力もれ", _
                 "master.asp", "マスタ保守")
    For i = 0 To UBound(navs) Step 2
        url = navs(i) : nm = navs(i + 1)
        Response.Write "<a href=""" & url & """" & _
                       IIfS(url = activeNav, " class=""on""", "") & ">" & H(nm) & "</a>"
    Next
    Response.Write "</nav>" & vbCrLf
    Response.Write "</header>" & vbCrLf
    Response.Write "<main>" & vbCrLf
End Sub

Sub PageFoot()
    Response.Write "</main>" & vbCrLf
    Response.Write "<footer>" & H(APP_NAME) & " " & H(APP_VERSION) & _
                   " ／ データは Access (.accdb) に保存されます。" & _
                   "デスクトップ版と同じデータを見ています。</footer>" & vbCrLf
    Response.Write "</body></html>"
End Sub

Function IIfS(cond, a, b)
    If cond Then IIfS = a Else IIfS = b
End Function

Sub Notice(kind, msg)
    Response.Write "<p class=""notice " & kind & """>" & H(msg) & "</p>" & vbCrLf
End Sub

' 日付を前後に動かすリンク付きの日付ピッカー
Sub DateNav(page, dt, extra)
    Response.Write "<form class=""datenav"" method=""get"" action=""" & page & """>" & vbCrLf
    Response.Write "  <a class=""btn ghost"" href=""" & page & "?d=" & _
                   Ymd(DateAdd("d", -1, dt)) & extra & """>&laquo; 前日</a>" & vbCrLf
    Response.Write "  <input type=""date"" name=""d"" value=""" & Ymd(dt) & """>" & vbCrLf
    Response.Write "  <button class=""btn"" type=""submit"">表示</button>" & vbCrLf
    Response.Write "  <a class=""btn ghost"" href=""" & page & "?d=" & _
                   Ymd(DateAdd("d", 1, dt)) & extra & """>翌日 &raquo;</a>" & vbCrLf
    Response.Write "  <span class=""wareki"">" & H(Wareki(dt)) & "</span>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub
%>
