<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/grid.asp"-->
<%
' 個人別の受付表（職員用）。その日に入力のあった人ごとに、受付の表を 1 人 1 枚で出す。
' 現行 Excel の「印刷」マクロ（入力のあった人の記入用フォームを 1 枚ずつ印刷していた）にあたる。
' PDF（printpdf.asp?kind=detail）も、同じ PersonSheetsHtml を使う。
Dim d, dd, msgNg, sheets
RequireStaff
d = PageDate()
msgNg = ""
If IsPost() Then
    If ParseYMD(FormVal("d"), dd) Then Go "detail.asp?d=" & YMD(dd)
    msgNg = "日付が読めません。2026-08-25 のように入れてください。"
End If
GridLoad
sheets = PersonSheetsHtml(d)
PageHead "個人別の受付表"
If msgNg <> "" Then ShowMsg "ng", msgNg
DateBar "detail.asp", d, "", ""
If sheets = "" Then
    ShowMsg "note", "この日は、受付入力・その他業務・特殊な問合せの内容のどれも入っている人がいません。"
Else
%>
<p class="noprint">
  <a class="btn" href="printpdf.asp?d=<%= YMD(d) %>&amp;kind=detail">PDF で出す（1 人 1 枚）</a>
  <a class="btn btn-sub" href="report.asp?d=<%= YMD(d) %>">日報の帳票を見る</a>
</p>
<p class="note noprint">入力のあった人だけを、1 人 1 枚で出します。数は受付入力の数そのものです（日報と同じ数え方）。</p>
<%
    Response.Write sheets
End If
PageFoot
%>
