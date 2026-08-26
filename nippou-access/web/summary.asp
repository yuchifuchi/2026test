<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/layout.asp"-->
<%
Dim d1, d2, rs, cols, i, tot(6), grand, v

' 既定は今週の月〜金 (現行「集計表」の B2:B6 と同じ範囲)
d1 = ParamDate("d1", MondayOf(Date()))
d2 = ParamDate("d2", DateAdd("d", 4, MondayOf(Date())))
If d2 < d1 Then d2 = d1

cols = Array("申込方法","抽選結果","納付書発送","商品発送","その他","商品交換","計")

PageHead "集計表", "summary.asp"
%>
<h1>集計表</h1>
<p class="lead">
  日報とまったく同じデータから集計しています。現行 Excel のように
  SUMIF の範囲がずれて数値が食い違うことは起きません。
</p>

<form class="datenav" method="get" action="summary.asp">
  <label>開始日 <input type="date" name="d1" value="<%= Ymd(d1) %>"></label>
  <label>終了日 <input type="date" name="d2" value="<%= Ymd(d2) %>"></label>
  <button class="btn" type="submit">表示</button>
  <a class="btn ghost" href="summary.asp?d1=<%= Ymd(MondayOf(Date())) %>&d2=<%= Ymd(DateAdd("d",4,MondayOf(Date()))) %>">今週</a>
  <a class="btn ghost" href="summary.asp?d1=<%= Ymd(DateAdd("d",-7,MondayOf(Date()))) %>&d2=<%= Ymd(DateAdd("d",-3,MondayOf(Date()))) %>">先週</a>
  <button class="btn ghost" type="button" onclick="window.print()">印刷</button>
</form>

<h2>日別</h2>
<div class="table-wrap">
<table>
  <thead><tr><th>日付</th>
<% For i = 0 To UBound(cols) %><th class="num"><%= H(cols(i)) %></th><% Next %>
  </tr></thead>
  <tbody>
<%
grand = 0
Set rs = DbQuery( _
  "SELECT J.[対象日], " & _
  " Sum(IIf(KB.[集計列ID]=3,J.[件数],0)) AS [申込方法], " & _
  " Sum(IIf(KB.[集計列ID]=4,J.[件数],0)) AS [抽選結果], " & _
  " Sum(IIf(KB.[集計列ID]=5,J.[件数],0)) AS [納付書発送], " & _
  " Sum(IIf(KB.[集計列ID]=6,J.[件数],0)) AS [商品発送], " & _
  " Sum(IIf(KB.[集計列ID]=7,J.[件数],0)) AS [その他], " & _
  " Sum(IIf(KB.[集計列ID]=8,J.[件数],0)) AS [商品交換], " & _
  " Sum(J.[件数]) AS [計] " & _
  "FROM [T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID] " & _
  "WHERE J.[対象日] BETWEEN ? AND ? " & _
  "GROUP BY J.[対象日] ORDER BY J.[対象日]", Array(d1, d2))
Do While Not rs.EOF
%>
    <tr>
      <td><a href="daily.asp?d=<%= Ymd(rs("対象日")) %>"><%= H(YmdSlash(rs("対象日"))) %>（<%= H(Youbi(rs("対象日"))) %>）</a></td>
<%
    For i = 0 To UBound(cols)
        v = CLng(rs(cols(i)))
        tot(i) = tot(i) + v
        Response.Write "<td class=""num"">" & v & "</td>"
    Next
%>
    </tr>
<%
    rs.MoveNext
Loop
rs.Close
%>
  </tbody>
  <tfoot><tr><td>計</td>
<% For i = 0 To UBound(cols) %><td class="num"><%= tot(i) %></td><% Next %>
  </tr></tfoot>
</table>
</div>

<h2>明細</h2>
<div class="table-wrap">
<table>
  <thead><tr>
    <th>日付</th><th>担当者</th><th>ブロック</th><th>製品</th><th>区分</th>
    <th>集計列</th><th class="num">件数</th><th>備考</th>
  </tr></thead>
  <tbody>
<%
Set rs = DbQuery( _
  "SELECT * FROM [Q_受電明細] WHERE [対象日] BETWEEN ? AND ? " & _
  "ORDER BY [対象日],[氏名],[区分ID]", Array(d1, d2))
Do While Not rs.EOF
%>
    <tr>
      <td><%= H(YmdSlash(rs("対象日"))) %></td>
      <td><%= H(rs("氏名")) %></td>
      <td><%= H(rs("ブロック名")) %></td>
      <td><%= H(rs("製品名")) %></td>
      <td><%= H(rs("区分名")) %></td>
      <td><%= H(rs("集計列名")) %></td>
      <td class="num"><%= H(rs("計")) %></td>
      <td><%= H(rs("備考2")) %></td>
    </tr>
<%
    rs.MoveNext
Loop
rs.Close
%>
  </tbody>
</table>
</div>
<% PageFoot() : CloseDb %>
