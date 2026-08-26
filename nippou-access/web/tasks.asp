<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' -----------------------------------------------------------------------------
'  電話応対以外の業務 (帳票の ①〜⑬ 欄)
'
'  受付入力と分けているのは、入力の性格が違うため。
'  受付は 1 件ずつ増えていくが、こちらは 1 日の終わりに件数をまとめて入れる。
'  そのため画面も「全項目を並べて一度に保存」の形にしてある。
' -----------------------------------------------------------------------------
Dim dt, opId, act, msg, msgKind, rs, v, saved, total

dt   = ParamDate("d", Date())
opId = ParamLong("op", 0)
act  = ParamText("act")
msg = "" : msgKind = "ok"

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" And act = "save" Then
    If opId = 0 Then
        msg = "担当者を選んでください。" : msgKind = "err"
    Else
        Ensure日報 dt
        Ensure出勤 dt, opId
        saved = 0

        Set rs = DbQuery("SELECT [業務項目ID] FROM [M_業務項目] WHERE [有効]=True " & _
                         "ORDER BY [表示順]", Empty)
        Do While Not rs.EOF
            v = ParamLong("t" & rs("業務項目ID"), 0)
            If v > 0 Then
                If DbScalar("SELECT Count(*) FROM [T_業務実績] " & _
                            "WHERE [対象日]=? AND [担当者ID]=? AND [業務項目ID]=?", _
                            Array(dt, opId, CLng(rs("業務項目ID"))), 0) > 0 Then
                    DbExec "UPDATE [T_業務実績] SET [件数]=? " & _
                           "WHERE [対象日]=? AND [担当者ID]=? AND [業務項目ID]=?", _
                           Array(v, dt, opId, CLng(rs("業務項目ID")))
                Else
                    DbExec "INSERT INTO [T_業務実績] " & _
                           "([対象日],[担当者ID],[業務項目ID],[件数]) VALUES (?,?,?,?)", _
                           Array(dt, opId, CLng(rs("業務項目ID")), v)
                End If
                saved = saved + v
            Else
                ' 0 を入れ直した項目は行ごと消す。0 の行を残すと
                ' 「入力済みなのか未入力なのか」が分からなくなるため。
                DbExec "DELETE FROM [T_業務実績] " & _
                       "WHERE [対象日]=? AND [担当者ID]=? AND [業務項目ID]=?", _
                       Array(dt, opId, CLng(rs("業務項目ID")))
            End If
            rs.MoveNext
        Loop
        rs.Close
        msg = "保存しました。合計 " & saved & " 件。"
    End If
End If

PageHead "その他業務", "tasks.asp"
%>
<h1>電話応対以外の業務</h1>
<p class="lead">
  受注入力や架電など、電話を受ける以外の作業の件数を入れます。
  日報の下半分（①〜⑬ の欄）になります。件数が無い項目は空のままで構いません。
</p>
<% If Len(msg) > 0 Then Notice msgKind, msg %>

<form class="datenav" method="get" action="tasks.asp">
  <a class="btn ghost" href="tasks.asp?d=<%= Ymd(DateAdd("d",-1,dt)) %>&op=<%= opId %>">&laquo; 前日</a>
  <input type="date" name="d" value="<%= Ymd(dt) %>">
  <select name="op">
    <option value="0">-- 担当者を選ぶ --</option>
<%
Set rs = OperatorsOn(dt)
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("担当者ID") & """" & _
        IIfS(CLng(rs("担当者ID")) = opId, " selected", "") & ">" & _
        H(rs("担当者コード") & "  " & rs("氏名")) & "</option>"
    rs.MoveNext
Loop
rs.Close
%>
  </select>
  <button class="btn" type="submit">表示</button>
  <a class="btn ghost" href="tasks.asp?d=<%= Ymd(DateAdd("d",1,dt)) %>&op=<%= opId %>">翌日 &raquo;</a>
  <span class="wareki"><%= H(Wareki(dt)) %></span>
</form>

<% If opId = 0 Then %>
  <div class="panel"><p>担当者を選ぶと入力欄が出ます。</p></div>
<% Else %>
<form method="post" action="tasks.asp">
  <input type="hidden" name="d" value="<%= Ymd(dt) %>">
  <input type="hidden" name="op" value="<%= opId %>">
  <input type="hidden" name="act" value="save">
  <div class="table-wrap">
  <table>
    <thead><tr><th style="width:60%">項目</th><th class="num" style="width:20%">件数</th><th></th></tr></thead>
    <tbody>
<%
total = 0
Set rs = DbQuery( _
  "SELECT TM.[業務項目ID], TM.[番号], TM.[項目名], " & _
  " (SELECT Nz(Sum(W.[件数]),0) FROM [T_業務実績] AS W " & _
  "   WHERE W.[対象日]=? AND W.[担当者ID]=? AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] " & _
  "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]", Array(dt, opId))
Do While Not rs.EOF
    v = CLng(rs("件数"))
    total = total + v
%>
      <tr>
        <td><%= H(rs("番号")) %>　<%= H(rs("項目名")) %></td>
        <td class="num">
          <input type="number" name="t<%= rs("業務項目ID") %>" min="0" step="1"
                 value="<%= IIfS(v = 0, "", v) %>" placeholder="0">
        </td>
        <td>件</td>
      </tr>
<%
    rs.MoveNext
Loop
rs.Close
%>
    </tbody>
    <tfoot><tr><td>合計</td><td class="num"><%= total %></td><td>件</td></tr></tfoot>
  </table>
  </div>
  <div class="row" style="margin-top:16px">
    <button class="btn" type="submit">保存する</button>
    <a class="btn ghost" href="entry.asp?d=<%= Ymd(dt) %>&op=<%= opId %>">受付入力に戻る</a>
  </div>
</form>
<% End If %>
<% PageFoot() : CloseDb %>
