<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
' ============================================================
'  エラーの説明画面
'
'  web.config の httpErrors で、ASP の実行時エラー（500.100）のときに IIS がこのページを動かす。
'  ・Server.GetLastError() は、この形で呼ばれたページの中でだけ中身を返す。
'  ・include を 1 つも使わない。include 自体が壊れていても、この画面だけは出るように。
'  ・HTTP 200 で返す。500 のままだと IIS が本文を捨てて既定の画面に差し替えることがある（7-3）。
'  ・この画面自体が失敗すると何も出なくなるので、値を取り出す所は 1 行ずつ守る。
' ============================================================
Dim e, fFile, fLine, fDesc, fAspDesc, fCat, fNum, url, kind, content, copyText

fFile = ""
fLine = ""
fDesc = ""
fAspDesc = ""
fCat = ""
fNum = 0
On Error Resume Next
Set e = Server.GetLastError()
fFile = e.File
fLine = e.Line
fDesc = e.Description
fAspDesc = e.ASPDescription
fCat = e.Category
fNum = e.Number
Response.Status = "200 OK"
On Error GoTo 0

url = Request.ServerVariables("HTTP_X_ORIGINAL_URL")
If url = "" Then
    url = Request.ServerVariables("URL")
    If Request.ServerVariables("QUERY_STRING") <> "" Then url = url & "?" & Request.ServerVariables("QUERY_STRING")
End If

content = CStr(fDesc)
If CStr(fAspDesc) <> "" Then content = content & " / " & CStr(fAspDesc)
kind = CStr(fCat)
If fNum <> 0 Then kind = kind & "（" & Hex(fNum) & "）"

copyText = "ファイル名: " & CStr(fFile) & vbCrLf & "行番号: " & CStr(fLine) & vbCrLf & "内容: " & content & vbCrLf & _
           "種別: " & kind & vbCrLf & "URL: " & url

Function EH(v)
    Dim s
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EH = s
End Function
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>うまく動きませんでした</title>
<link rel="stylesheet" href="css/style.css">
<style>
  /* css/style.css が読めないときでも読めるように、最低限の見た目をここにも書く */
  body { font-family: "Meiryo UI", Meiryo, sans-serif; }
  .e5 { border-collapse: collapse; background: #fff; }
  .e5 th, .e5 td { border: 1px solid #999; padding: 6px 10px; text-align: left; vertical-align: top; }
  .e5 th { width: 7em; background: #f3e3e3; font-weight: normal; }
</style>
</head><body>
<!-- NIPPOU-ERROR -->
<div class="band band-ng"><div class="band-in"><span class="app">電話応対日報 集計システム</span></div></div>
<div class="wrap">
<h1>うまく動きませんでした</h1>
<p>画面を作る途中でエラーが起きました。下の <b>5 行をそのまま</b>、職員（またはシステムの担当）にお知らせください。</p>
<table class="e5">
<tr><th>ファイル名</th><td><%= EH(fFile) %></td></tr>
<tr><th>行番号</th><td><%= EH(fLine) %></td></tr>
<tr><th>内容</th><td><%= EH(content) %></td></tr>
<tr><th>種別</th><td><%= EH(kind) %></td></tr>
<tr><th>URL</th><td><%= EH(url) %></td></tr>
</table>
<p>メールなどに貼り付けるときは、下の枠の中をすべて選んでコピーしてください。</p>
<textarea rows="6" cols="90" readonly><%= EH(copyText) %></textarea>
<p><a href="default.asp">メニューへ戻る</a></p>
</div>
</body></html>
