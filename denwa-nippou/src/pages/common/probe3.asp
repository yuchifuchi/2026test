<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<%
' 切り分け 3：＋ SQL の置き場所（sql.asp）
Dim n
n = Len(SqlColumnTotals())
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>切り分け 3</title></head><body>
<p>PROBE3-OK</p>
<p>切り分け 3：sql.asp を読めました（SQL の長さ <%= n %> 文字）。</p>
</body></html>
