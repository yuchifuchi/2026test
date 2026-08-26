<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/layout.asp"-->
<%
Dim dt, opId, act, msg, msgKind, rs, total, kbId, prId, cnt, rowId

dt   = ParamDate("d", Date())
opId = ParamLong("op", 0)
act  = ParamText("act")
msg = "" : msgKind = "ok"

' --- 更新系は POST のみ受け付ける ---
If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    Ensure日報 dt
    If opId > 0 Then Ensure出勤 dt, opId

    Select Case act
    Case "add"
        kbId = ParamLong("kubun", 0)
        prId = ParamLong("prod", 0)
        cnt  = ParamLong("cnt", 0)
        If opId = 0 Or kbId = 0 Then
            msg = "担当者と区分を選んでください。" : msgKind = "err"
        ElseIf cnt = 0 Then
            msg = "件数を入力してください。" : msgKind = "err"
        Else
            ' 同じ 日付・担当者・区分・製品 は 1 行にまとめる (二重計上を作らない)
            If DbScalar("SELECT Count(*) FROM [T_受電] WHERE [対象日]=? AND [担当者ID]=?" & _
                        " AND [区分ID]=? AND [製品ID]=?", Array(dt, opId, kbId, prId), 0) > 0 Then
                DbExec "UPDATE [T_受電] SET [件数]=[件数]+?, [更新日時]=Now() " & _
                       "WHERE [対象日]=? AND [担当者ID]=? AND [区分ID]=? AND [製品ID]=?", _
                       Array(cnt, dt, opId, kbId, prId)
                msg = "既存の行に " & cnt & " 件を足しました。"
            Else
                DbExec "INSERT INTO [T_受電] " & _
                       "([対象日],[担当者ID],[区分ID],[製品ID],[件数],[備考2]," & _
                       " [登録日時],[更新日時],[登録者]) VALUES (?,?,?,?,?,?,Now(),Now(),?)", _
                       Array(dt, opId, kbId, prId, cnt, ParamText("biko2"), "web")
                msg = cnt & " 件を登録しました。"
            End If
        End If

    Case "upd"
        rowId = ParamLong("id", 0)
        cnt = ParamLong("cnt", 0)
        If rowId > 0 Then
            If cnt <= 0 Then
                DbExec "DELETE FROM [T_受電] WHERE [受電ID]=?", Array(rowId)
                msg = "1 行削除しました。"
            Else
                DbExec "UPDATE [T_受電] SET [件数]=?,[備考2]=?,[更新日時]=Now() " & _
                       "WHERE [受電ID]=?", Array(cnt, ParamText("biko2"), rowId)
                msg = "件数を更新しました。"
            End If
        End If

    Case "del"
        rowId = ParamLong("id", 0)
        If rowId > 0 Then
            DbExec "DELETE FROM [T_受電] WHERE [受電ID]=?", Array(rowId)
            msg = "1 行削除しました。"
        End If
    End Select
End If

PageHead "受付入力", "entry.asp"
%>
<h1>受付入力</h1>
<p class="lead">
  担当者を選んでから件数を登録してください。ここで入れた数値がそのまま日報にも集計表にも出ます。
  転記の操作は要りません。
</p>
<% If Len(msg) > 0 Then Notice msgKind, msg %>

<form class="datenav" method="get" action="entry.asp">
  <a class="btn ghost" href="entry.asp?d=<%= Ymd(DateAdd("d",-1,dt)) %>&op=<%= opId %>">&laquo; 前日</a>
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
  <a class="btn ghost" href="entry.asp?d=<%= Ymd(DateAdd("d",1,dt)) %>&op=<%= opId %>">翌日 &raquo;</a>
  <span class="wareki"><%= H(Wareki(dt)) %></span>
</form>

<% If opId = 0 Then %>
  <div class="panel"><p>担当者を選ぶと入力欄が出ます。</p></div>
<% Else %>

<div class="panel">
  <h2 style="margin-top:0">追加する</h2>
  <form method="post" action="entry.asp">
    <input type="hidden" name="d" value="<%= Ymd(dt) %>">
    <input type="hidden" name="op" value="<%= opId %>">
    <input type="hidden" name="act" value="add">
    <div class="row">
      <div class="field" style="flex:2 1 320px">
        <label for="kubun">区分</label>
        <select id="kubun" name="kubun" required>
          <option value="">-- 選択 --</option>
<%
Set rs = DbQuery("SELECT [区分ID],[表示名] FROM [Q_選択_区分]", Empty)
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("区分ID") & """>" & H(rs("表示名")) & "</option>"
    rs.MoveNext
Loop
rs.Close
%>
        </select>
      </div>
      <div class="field" style="flex:2 1 320px">
        <label for="prod">製品（不要なら「製品指定なし」のまま）</label>
        <select id="prod" name="prod">
<%
Set rs = DbQuery("SELECT [製品ID],[表示名] FROM [Q_選択_製品]", Empty)
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("製品ID") & """>" & H(rs("表示名")) & "</option>"
    rs.MoveNext
Loop
rs.Close
%>
        </select>
      </div>
      <div class="field">
        <label for="cnt">件数</label>
        <input id="cnt" type="number" name="cnt" value="1" min="1" step="1">
      </div>
      <div class="field" style="flex:1 1 220px">
        <label for="biko2">備考</label>
        <input id="biko2" type="text" name="biko2" maxlength="255">
      </div>
      <div><button class="btn" type="submit">登録</button></div>
    </div>
  </form>
</div>

<h2>この日の入力内容</h2>
<div class="table-wrap">
<table>
  <thead><tr>
    <th>区分</th><th>製品</th><th class="num">件数</th><th>備考</th>
    <th>集計列</th><th></th>
  </tr></thead>
  <tbody>
<%
total = 0
Set rs = DbQuery( _
  "SELECT J.[受電ID], KB.[区分名], BK.[ブロック名], PR.[製品名], J.[件数], J.[備考2], SC.[集計列名] " & _
  "FROM ((([T_受電] AS J INNER JOIN [M_区分] AS KB ON J.[区分ID]=KB.[区分ID]) " & _
  "  INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID]) " & _
  "  INNER JOIN [M_集計列] AS SC ON KB.[集計列ID]=SC.[集計列ID]) " & _
  "  INNER JOIN [M_製品] AS PR ON J.[製品ID]=PR.[製品ID] " & _
  "WHERE J.[対象日]=? AND J.[担当者ID]=? ORDER BY KB.[区分ID]", Array(dt, opId))
Do While Not rs.EOF
    total = total + CLng(rs("件数"))
%>
    <tr>
      <td><%= H(rs("ブロック名") & " / " & rs("区分名")) %></td>
      <td><%= H(rs("製品名")) %></td>
      <td class="num">
        <form method="post" action="entry.asp" style="display:flex;gap:6px;justify-content:flex-end">
          <input type="hidden" name="d" value="<%= Ymd(dt) %>">
          <input type="hidden" name="op" value="<%= opId %>">
          <input type="hidden" name="act" value="upd">
          <input type="hidden" name="id" value="<%= rs("受電ID") %>">
          <input type="number" name="cnt" value="<%= H(rs("件数")) %>" min="0" step="1">
          <button class="btn ghost" type="submit">更新</button>
        </form>
      </td>
      <td><%= H(rs("備考2")) %></td>
      <td><%= H(rs("集計列名")) %></td>
      <td>
        <form method="post" action="entry.asp" onsubmit="return confirm('この行を削除します。よろしいですか？')">
          <input type="hidden" name="d" value="<%= Ymd(dt) %>">
          <input type="hidden" name="op" value="<%= opId %>">
          <input type="hidden" name="act" value="del">
          <input type="hidden" name="id" value="<%= rs("受電ID") %>">
          <button class="btn danger" type="submit">削除</button>
        </form>
      </td>
    </tr>
<%
    rs.MoveNext
Loop
rs.Close
If total = 0 Then Response.Write "<tr><td colspan=""6"">まだ入力がありません。</td></tr>"
%>
  </tbody>
  <tfoot><tr><td colspan="2">合計</td><td class="num"><%= total %></td><td colspan="3"></td></tr></tfoot>
</table>
</div>
<% End If %>
<% PageFoot() : CloseDb %>
