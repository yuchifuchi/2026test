<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<%
' 切り分け 6：＋ 役割の判定（auth.asp）
Dim r
If IsStaff() Then r = "職員用" Else r = "パート職員用"
DbClose
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>切り分け 6</title></head><body>
<p>PROBE6-OK</p>
<p>切り分け 6：auth.asp を読めました。このフォルダは「<%= r %>」と判定されました。</p>
</body></html>
