<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<%
' 切り分け 4：＋ データベースの出入り口（db.asp）。まだ接続はしない。
Dim p
p = DbResolvePath()
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>切り分け 4</title></head><body>
<p>PROBE4-OK</p>
<p>切り分け 4：db.asp を読めました。データベースの場所は <%= DbH(p) %> と解釈しました。</p>
</body></html>
