<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<%
' 切り分け 2：設定ファイル（config.asp）が読めることを確かめる
Function ProbeH(v)
    ProbeH = Replace(Replace(Replace(CStr(v), "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>切り分け 2</title></head><body>
<p>PROBE2-OK</p>
<p>切り分け 2：config.asp を読めました。ROLE = <%= ProbeH(ROLE) %> ／ DB_PATH = <%= ProbeH(DB_PATH) %></p>
</body></html>
