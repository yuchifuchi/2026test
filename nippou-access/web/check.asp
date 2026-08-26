<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/layout.asp"-->
<%
Dim dt, rs, n
dt = ParamDate("d", Date())
PageHead "入力もれチェック", "check.asp"
%>
<h1>入力もれチェック</h1>
<p class="lead">
  出勤登録があるのに、受付・業務のどちらも 1 件も入っていない担当者を出します。
  現行 Excel では、担当者名の設定ミスで 6 名分のデータが別人に合算されたり
  消えたりしていましたが、誰も気付けませんでした。その再発を防ぐための画面です。
</p>
<% DateNav "check.asp", dt, "" %>

<div class="table-wrap">
<table>
  <thead><tr><th>コード</th><th>氏名</th><th></th></tr></thead>
  <tbody>
<%
n = 0
Set rs = DbQuery("SELECT * FROM [Q_未入力チェック] WHERE [対象日]=?", Array(dt))
Do While Not rs.EOF
    n = n + 1
%>
    <tr>
      <td><%= H(rs("担当者コード")) %></td>
      <td><%= H(rs("氏名")) %></td>
      <td><a class="btn ghost" href="entry.asp?d=<%= Ymd(dt) %>&op=<%= rs("担当者ID") %>">入力画面を開く</a></td>
    </tr>
<%
    rs.MoveNext
Loop
rs.Close
If n = 0 Then Response.Write "<tr><td colspan=""3"">入力もれはありません。</td></tr>"
%>
  </tbody>
</table>
</div>
<% PageFoot() : CloseDb %>
