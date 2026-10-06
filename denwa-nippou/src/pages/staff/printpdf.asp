<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/sheet.asp"-->
<!--#include file="include/pdf.asp"-->
<%
' 帳票を PDF にして送る（職員用）。画面で見る report.asp と同じ SheetHtml を使う。
' ?selftest=1 は設置チェックからの「試しに PDF を作る」。
Dim d, html, pdfPath, msg, stamp
RequireStaff
Server.ScriptTimeout = PDF_TIMEOUT + 60
d = PageDate()
If QueryVal("selftest") = "1" Then
    html = PdfDocument("<div class=""sheet""><div class=""sh-title"">PDF の試し作成</div><p>この PDF が開けば、PDF を作る仕組みは動いています。</p></div>")
Else
    html = PdfDocument(SheetHtml(d))
End If
stamp = Year(d) & Pad2(Month(d)) & Pad2(Day(d))
If MakePdf(html, pdfPath, msg) Then
    SendPdf pdfPath, "nippou_" & stamp & ".pdf", "電話応対日報_" & stamp & ".pdf"
End If
PageHead "PDF を作れませんでした"
ShowMsg "ng", msg
%>
<p>PDF を作る仕組み（サーバーの Microsoft Edge）がうまく動きませんでした。</p>
<p>急ぐときは、<a href="report.asp?d=<%= YMD(d) %>">帳票を画面で見る</a> を開いて、ブラウザーの「印刷」（Ctrl キーを押しながら P）から印刷してください。用紙は A4・縦、「ヘッダーとフッター」は外してください。</p>
<p>何度やっても同じときは、職員メニューの「設置チェック」を開き、赤い行を確かめてください。</p>
<% PageFoot %>
