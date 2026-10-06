<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' 切り分け 7：＋ 画面の枠（layout.asp）
PageHead "切り分け 7"
%>
<p>PROBE7-OK</p>
<p>切り分け 7：画面の枠（青い帯）が出ていれば、ここまでの部品はすべて読めています。</p>
<% PageFoot %>
