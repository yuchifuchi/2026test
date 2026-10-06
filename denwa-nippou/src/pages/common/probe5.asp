<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<%
' 切り分け 5：＋ データベースに実際に接続して、担当者の数を数える
Dim n
n = DbScalar(SqlStaffCount(), Array(), 0)
DbClose
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>切り分け 5</title></head><body>
<p>PROBE5-OK</p>
<p>切り分け 5：データベースに接続できました（担当者 <%= CLng(n) %> 人）。</p>
</body></html>
