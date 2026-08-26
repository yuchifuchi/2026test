<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
Dim dt, act, msg, msgKind, rs, hd, n, missing, opId

dt  = ParamDate("d", Date())
act = ParamText("act")
msg = "" : msgKind = "ok"

Ensure日報 dt

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    Select Case act
    Case "save"
        DbExec "UPDATE [T_日報] SET [回線数]=?,[特記事項]=?,[職員代替案件]=?,[要望]=?," & _
               "[更新日時]=Now() WHERE [対象日]=?", _
               Array(ParamLong("kaisen", 0), ParamText("tokki"), _
                     ParamText("daitai"), ParamText("youbou"), dt)
        msg = "日報を保存しました。"

    Case "fix"
        missing = DbScalar("SELECT Count(*) FROM [Q_未入力チェック] WHERE [対象日]=?", Array(dt), 0)
        If missing > 0 And ParamText("force") <> "1" Then
            msg = "実績が 1 件も無い担当者が " & missing & " 名います。" & _
                  "確認してから、もう一度「確定する」を押してください。"
            msgKind = "warn"
        Else
            DbExec "UPDATE [T_日報] SET [状態]='確定',[確定日時]=Now(),[更新日時]=Now() " & _
                   "WHERE [対象日]=?", Array(dt)
            msg = "日報を確定しました。"
        End If

    Case "unfix"
        DbExec "UPDATE [T_日報] SET [状態]='入力中',[確定日時]=Null,[更新日時]=Now() " & _
               "WHERE [対象日]=?", Array(dt)
        msg = "確定を解除しました。"

    Case "fill"
        n = 0
        Set rs = OperatorsOn(dt)
        Do While Not rs.EOF
            If DbScalar("SELECT Count(*) FROM [T_出勤] WHERE [対象日]=? AND [担当者ID]=?", _
                        Array(dt, CLng(rs("担当者ID"))), 0) = 0 Then
                DbExec "INSERT INTO [T_出勤] ([対象日],[担当者ID]) VALUES (?,?)", _
                       Array(dt, CLng(rs("担当者ID")))
                n = n + 1
            End If
            rs.MoveNext
        Loop
        rs.Close
        msg = n & " 名を出勤者に追加しました。休みの方は行を削除してください。"

    Case "attend_del"
        opId = ParamLong("op", 0)
        If opId > 0 Then
            DbExec "DELETE FROM [T_出勤] WHERE [対象日]=? AND [担当者ID]=?", Array(dt, opId)
            msg = "出勤者から外しました。"
        End If

    Case "attend_add"
        opId = ParamLong("op", 0)
        If opId > 0 Then
            Ensure出勤 dt, opId
            msg = "出勤者に追加しました。"
        End If
    End Select
End If

Set hd = DbQuery("SELECT * FROM [Q_日報_ヘッダ] WHERE [対象日]=?", Array(dt))
missing = DbScalar("SELECT Count(*) FROM [Q_未入力チェック] WHERE [対象日]=?", Array(dt), 0)

PageHead "日報", "daily.asp"
%>
<h1>日報</h1>
<p class="lead">
  件数欄は受付入力から自動集計されます。手で直すことはできません
  ―― ここが現行 Excel との一番の違いで、日報と集計表がずれる原因を断っています。
</p>
<% If Len(msg) > 0 Then Notice msgKind, msg %>
<% DateNav "daily.asp", dt, "" %>

<div class="panel">
  <div class="row" style="justify-content:space-between">
    <div>
      状態：
<% If hd("状態") = "確定" Then %>
      <span class="badge fix">確定済み（<%= H(hd("確定日時")) %>）</span>
<% Else %>
      <span class="badge wip">入力中</span>
<% End If %>
    </div>
    <div class="row">
      <a class="btn ghost" href="report.asp?d=<%= Ymd(dt) %>">印刷用を開く</a>
<% If hd("状態") = "確定" Then %>
      <form method="post" action="daily.asp">
        <input type="hidden" name="d" value="<%= Ymd(dt) %>">
        <input type="hidden" name="act" value="unfix">
        <button class="btn ghost" type="submit">確定を解除</button>
      </form>
<% Else %>
      <form method="post" action="daily.asp">
        <input type="hidden" name="d" value="<%= Ymd(dt) %>">
        <input type="hidden" name="act" value="fix">
        <input type="hidden" name="force" value="<%= IIfS(msgKind="warn","1","0") %>">
        <button class="btn" type="submit">確定する</button>
      </form>
<% End If %>
    </div>
  </div>
</div>

<h2>問合せ件数（自動集計）</h2>
<div class="table-wrap">
<table>
  <thead><tr>
    <th>合計</th><th class="num">申込</th><th class="num">抽選</th><th class="num">払込用紙</th>
    <th class="num">商品発送</th><th class="num">その他</th>
    <th class="num">内 交換</th><th class="num">内 返金</th>
  </tr></thead>
  <tbody>
    <tr>
      <td class="num"><b style="font-size:19px"><%= H(hd("合計")) %></b> 件</td>
      <td class="num"><%= H(hd("申込")) %></td>
      <td class="num"><%= H(hd("抽選")) %></td>
      <td class="num"><%= H(hd("払込用紙")) %></td>
      <td class="num"><%= H(hd("商品発送")) %></td>
      <td class="num"><%= H(hd("その他")) %></td>
      <td class="num"><%= H(hd("内交換")) %></td>
      <td class="num"><%= H(hd("内返金")) %></td>
    </tr>
    <tr>
      <td>うち職員受電</td>
      <td class="num">(<%= H(hd("申込_職員")) %>)</td>
      <td class="num">(<%= H(hd("抽選_職員")) %>)</td>
      <td class="num">(<%= H(hd("払込用紙_職員")) %>)</td>
      <td class="num">(<%= H(hd("商品発送_職員")) %>)</td>
      <td class="num">(<%= H(hd("その他_職員")) %>)</td>
      <td class="num" colspan="2"></td>
    </tr>
  </tbody>
</table>
</div>

<h2>出勤者（<%= H(hd("出勤者数")) %> 名）</h2>
<% If missing > 0 Then %>
  <p class="notice warn">
    出勤登録があるのに実績が 1 件も無い担当者が <%= H(missing) %> 名います。
    <a href="check.asp?d=<%= Ymd(dt) %>">入力もれチェックを開く</a>
  </p>
<% End If %>
<div class="panel">
  <div class="row" style="margin-bottom:12px">
    <form method="post" action="daily.asp">
      <input type="hidden" name="d" value="<%= Ymd(dt) %>">
      <input type="hidden" name="act" value="fill">
      <button class="btn ghost" type="submit">在籍者を一括で追加</button>
    </form>
    <form method="post" action="daily.asp" class="row" style="gap:8px">
      <input type="hidden" name="d" value="<%= Ymd(dt) %>">
      <input type="hidden" name="act" value="attend_add">
      <select name="op">
<%
Set rs = OperatorsOn(dt)
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("担当者ID") & """>" & _
                   H(rs("担当者コード") & "  " & rs("氏名")) & "</option>"
    rs.MoveNext
Loop
rs.Close
%>
      </select>
      <button class="btn ghost" type="submit">1 名追加</button>
    </form>
  </div>
  <div class="table-wrap">
  <table>
    <thead><tr><th>コード</th><th>氏名</th><th>区分</th><th class="num">受電件数</th><th></th></tr></thead>
    <tbody>
<%
Set rs = DbQuery( _
  "SELECT OP.[担当者ID],OP.[担当者コード],OP.[氏名],OP.[職員区分], " & _
  " (SELECT Nz(Sum(J.[件数]),0) FROM [T_受電] AS J " & _
  "   WHERE J.[対象日]=A.[対象日] AND J.[担当者ID]=A.[担当者ID]) AS [件数] " & _
  "FROM [T_出勤] AS A INNER JOIN [M_担当者] AS OP ON A.[担当者ID]=OP.[担当者ID] " & _
  "WHERE A.[対象日]=? ORDER BY OP.[表示順]", Array(dt))
Do While Not rs.EOF
%>
      <tr>
        <td><%= H(rs("担当者コード")) %></td>
        <td><%= H(rs("氏名")) %></td>
        <td><%= H(rs("職員区分")) %></td>
        <td class="num"><%= H(rs("件数")) %></td>
        <td>
          <form method="post" action="daily.asp">
            <input type="hidden" name="d" value="<%= Ymd(dt) %>">
            <input type="hidden" name="act" value="attend_del">
            <input type="hidden" name="op" value="<%= rs("担当者ID") %>">
            <button class="btn ghost" type="submit">外す</button>
          </form>
        </td>
      </tr>
<%
    rs.MoveNext
Loop
rs.Close
%>
    </tbody>
  </table>
  </div>
</div>

<h2>電話応対以外の業務（帳票の ①〜⑬ 欄）</h2>
<div class="table-wrap" style="margin-bottom:8px">
<table>
  <thead><tr><th>項目</th><th class="num">件数</th><th>項目</th><th class="num">件数</th></tr></thead>
  <tbody>
<%
Dim tk, tkN, tkV, ti, tn
ReDim tkN(99) : ReDim tkV(99)
tn = 0
Set tk = DbQuery( _
  "SELECT TM.[番号], TM.[項目名], " & _
  " (SELECT Nz(Sum(W.[件数]),0) FROM [T_業務実績] AS W " & _
  "   WHERE W.[対象日]=? AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] " & _
  "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]", Array(dt))
Do While Not tk.EOF
    tkN(tn) = tk("番号") & "　" & tk("項目名")
    tkV(tn) = CLng(tk("件数"))
    tn = tn + 1
    tk.MoveNext
Loop
tk.Close

Dim half : half = Int((tn + 1) / 2)
For ti = 0 To half - 1
%>
    <tr>
      <td><%= H(tkN(ti)) %></td>
      <td class="num"><%= IIfS(tkV(ti) = 0, "―", tkV(ti)) %></td>
      <td><%= IIfS(ti + half < tn, H(tkN(ti + half)), "") %></td>
      <td class="num"><%= IIfS(ti + half < tn, IIfS(tkV(ti + half) = 0, "―", tkV(ti + half)), "") %></td>
    </tr>
<% Next %>
  </tbody>
</table>
</div>
<p class="lead">
  この欄は「その他業務」画面から担当者ごとに入力します。
  <a href="tasks.asp?d=<%= Ymd(dt) %>">その他業務を入力する</a>
</p>

<h2>記述欄</h2>
<form class="panel" method="post" action="daily.asp">
  <input type="hidden" name="d" value="<%= Ymd(dt) %>">
  <input type="hidden" name="act" value="save">
  <div class="field" style="max-width:200px">
    <label for="kaisen">回線数</label>
    <input id="kaisen" type="number" name="kaisen" value="<%= H(hd("回線数")) %>" min="0">
  </div>
  <div class="field" style="margin-top:14px">
    <label for="tokki">特記事項（報告書の電話件数だけでは伝わり難い事項など）</label>
    <textarea id="tokki" name="tokki"><%= H(hd("特記事項")) %></textarea>
  </div>
  <div class="field" style="margin-top:14px">
    <label for="daitai">職員に代わった案件（概要）</label>
    <textarea id="daitai" name="daitai"><%= H(hd("職員代替案件")) %></textarea>
  </div>
  <div class="field" style="margin-top:14px">
    <label for="youbou">要望（お客様からの問い合わせを減らすための改善提案等）</label>
    <textarea id="youbou" name="youbou"><%= H(hd("要望")) %></textarea>
  </div>
  <div style="margin-top:16px"><button class="btn" type="submit">保存する</button></div>
</form>
<% hd.Close : PageFoot() : CloseDb %>
