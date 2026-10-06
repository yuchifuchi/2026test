<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/sheet.asp"-->
<%
' 帳票を画面で見る（職員用）。PDF（printpdf.asp）と同じ SheetHtml を使う。
Dim d, dd, msgNg, daily
RequireStaff
d = PageDate()
msgNg = ""
If IsPost() Then
    If ParseYMD(FormVal("d"), dd) Then Go "report.asp?d=" & YMD(dd)
    msgNg = "日付が読めません。2026-08-25 のように入れてください。"
End If
daily = DbQuery(SqlDaily(), Array(d, DateAdd("d", 1, d)))
PageHead "帳票を見る"
If msgNg <> "" Then ShowMsg "ng", msgNg
DateBar "report.asp", d, "", ""
%>
<p class="noprint">
  <a class="btn" href="printpdf.asp?d=<%= YMD(d) %>">PDF で出す（印刷はこちらから）</a>
  <a class="btn btn-sub" href="daily.asp?d=<%= YMD(d) %>">日報の作成・確定へ</a>
</p>
<%
If UBound(daily) < 0 Then
    ShowMsg "note", "この日の日報（出勤者・回線数・記述欄）は、まだ作っていません。件数は受付入力の数がそのまま出ています。"
ElseIf ToStr(daily(0)("状態")) <> "確定" Then
    ShowMsg "note", "この日の日報はまだ「作成中」です。確定すると、受付入力の数が直せなくなり、紙と画面の数が食い違わなくなります。"
End If
%>
<div class="paper">
<%= SheetHtml(d) %>
</div>
<% PageFoot %>
