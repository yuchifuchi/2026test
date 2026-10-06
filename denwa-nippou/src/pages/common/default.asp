<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' メニュー。part と staff で同じファイル。staff では職員メニュー（staff.asp）へ移る。
If IsStaff() Then Go "staff.asp"
PageHead "メニュー"
%>
<div class="menu">
  <a class="btn" href="entry.asp">受付入力<small>電話の件数を入れる</small></a>
  <a class="btn btn-sub" href="tasks.asp">その他業務<small>電話以外の仕事（①～⑬）の件数を入れる</small></a>
</div>
<p class="note">このページを「お気に入り」に入れておくと、次から簡単に開けます。</p>
<% PageFoot %>
