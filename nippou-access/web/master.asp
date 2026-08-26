<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/layout.asp"-->
<%
' -----------------------------------------------------------------------------
'  マスタ保守
'
'  担当者は「削除」しない。在籍終了日を入れると入力画面の候補から消えるが、
'  過去の日報・集計表はその担当者のまま残る。職員の入れ替わりが多くても
'  過去の数字が壊れないのがこの作りの狙い。
' -----------------------------------------------------------------------------
Dim tab_, act, msg, msgKind, rs, id, today

today = Date()
tab_ = ParamText("tab")
If Len(tab_) = 0 Then tab_ = "operator"
act = ParamText("act")
msg = "" : msgKind = "ok"

If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    Select Case act

    '--- 担当者 ---
    Case "op_add"
        If Len(ParamText("code")) = 0 Or Len(ParamText("sei")) = 0 Then
            msg = "担当者コードと姓は必須です。" : msgKind = "err"
        Else
            id = DbScalar("SELECT Max([担当者ID]) FROM [M_担当者]", Empty, 0) + 1
            DbExec "INSERT INTO [M_担当者] " & _
                   "([担当者ID],[担当者コード],[姓],[名],[氏名],[カナ],[職員区分]," & _
                   " [在籍開始日],[表示順],[有効]) VALUES (?,?,?,?,?,?,?,?,?,True)", _
                   Array(id, ParamText("code"), ParamText("sei"), ParamText("mei"), _
                         Trim(ParamText("sei") & " " & ParamText("mei")), ParamText("kana"), _
                         IIfS(ParamText("kbn") = "職員", "職員", "パート"), _
                         ParamDate("start", today), _
                         DbScalar("SELECT Max([表示順]) FROM [M_担当者]", Empty, 0) + 1)
            msg = ParamText("sei") & " さんを登録しました。"
        End If

    Case "op_upd"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_担当者] SET [担当者コード]=?,[姓]=?,[名]=?,[氏名]=?,[カナ]=?," & _
                   "[職員区分]=?,[在籍開始日]=?,[在籍終了日]=?,[表示順]=?,[有効]=? WHERE [担当者ID]=?", _
                   Array(ParamText("code"), ParamText("sei"), ParamText("mei"), _
                         Trim(ParamText("sei") & " " & ParamText("mei")), ParamText("kana"), _
                         IIfS(ParamText("kbn") = "職員", "職員", "パート"), _
                         ParamDate("start", Null), ParamDate("end", Null), _
                         ParamLong("ord", 99), IIfS(ParamText("act2") = "on", -1, 0), id)
            msg = "担当者を更新しました。"
        End If

    Case "op_retire"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_担当者] SET [在籍終了日]=? WHERE [担当者ID]=?", _
                   Array(ParamDate("end", today), id)
            msg = "退職日を登録しました。以降は入力画面の候補に出ません。" & _
                  "過去のデータはそのまま残ります。"
        End If

    Case "op_return"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_担当者] SET [在籍終了日]=Null,[有効]=True WHERE [担当者ID]=?", Array(id)
            msg = "在籍中に戻しました。"
        End If

    '--- 製品 ---
    Case "pr_add"
        id = DbScalar("SELECT Max([製品ID]) FROM [M_製品]", Empty, 0) + 1
        DbExec "INSERT INTO [M_製品] ([製品ID],[ブロックID],[製品名],[適用開始日],[表示順],[有効]) " & _
               "VALUES (?,?,?,?,?,True)", _
               Array(id, ParamLong("block", 1), ParamText("name"), ParamDate("start", today), _
                     ParamLong("ord", 99))
        msg = "製品を追加しました。"

    Case "pr_upd"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_製品] SET [製品名]=?,[適用開始日]=?,[適用終了日]=?," & _
                   "[表示順]=?,[有効]=? WHERE [製品ID]=?", _
                   Array(ParamText("name"), ParamDate("start", Null), ParamDate("end", Null), _
                         ParamLong("ord", 99), IIfS(ParamText("on") = "1", -1, 0), id)
            msg = "製品を更新しました。"
        End If

    '--- 区分 ---
    Case "kb_upd"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_区分] SET [区分名]=?,[集計列ID]=?,[内訳区分]=?," & _
                   "[表示順]=?,[有効]=? WHERE [区分ID]=?", _
                   Array(ParamText("name"), ParamLong("col", 7), ParamText("uchi"), _
                         ParamLong("ord", 99), IIfS(ParamText("on") = "1", -1, 0), id)
            msg = "区分を更新しました。"
        End If

    '--- 業務項目 ---
    Case "tk_upd"
        id = ParamLong("id", 0)
        If id > 0 Then
            DbExec "UPDATE [M_業務項目] SET [項目名]=?,[帳票表示名]=?,[表示順]=?,[有効]=? " & _
                   "WHERE [業務項目ID]=?", _
                   Array(ParamText("name"), ParamText("rname"), ParamLong("ord", 99), _
                         IIfS(ParamText("on") = "1", -1, 0), id)
            msg = "業務項目を更新しました。"
        End If
    End Select
End If

PageHead "マスタ保守", "master.asp"
%>
<h1>マスタ保守</h1>
<p class="lead">
  担当者・製品・区分の変更はここだけで完結します。
  Excel を配り直したり、ファイル名や B1 セルを直したりする必要はありません。
</p>
<% If Len(msg) > 0 Then Notice msgKind, msg %>

<div class="row" style="margin-bottom:18px">
  <a class="btn <%= IIfS(tab_="operator","","ghost") %>" href="master.asp?tab=operator">担当者</a>
  <a class="btn <%= IIfS(tab_="product","","ghost") %>" href="master.asp?tab=product">製品</a>
  <a class="btn <%= IIfS(tab_="kubun","","ghost") %>" href="master.asp?tab=kubun">区分</a>
  <a class="btn <%= IIfS(tab_="task","","ghost") %>" href="master.asp?tab=task">業務項目</a>
</div>

<% If tab_ = "operator" Then %>
<div class="panel">
  <h2 style="margin-top:0">担当者を追加する</h2>
  <form method="post" action="master.asp">
    <input type="hidden" name="tab" value="operator">
    <input type="hidden" name="act" value="op_add">
    <div class="row">
      <div class="field"><label>コード</label><input type="text" name="code" size="5" required></div>
      <div class="field"><label>姓</label><input type="text" name="sei" required></div>
      <div class="field"><label>名</label><input type="text" name="mei"></div>
      <div class="field"><label>カナ</label><input type="text" name="kana"></div>
      <div class="field"><label>区分</label>
        <select name="kbn"><option>パート</option><option>職員</option></select></div>
      <div class="field"><label>在籍開始日</label>
        <input type="date" name="start" value="<%= Ymd(today) %>"></div>
      <div><button class="btn" type="submit">追加</button></div>
    </div>
  </form>
</div>

<h2>担当者一覧</h2>
<p class="lead">
  退職する方は「退職日」を入れてください。<b>行は消さないでください。</b>
  消すと過去の日報・集計表からその方の名前が失われます。
</p>
<div class="table-wrap">
<table>
  <thead><tr>
    <th>コード</th><th>姓</th><th>名</th><th>カナ</th><th>区分</th>
    <th>在籍開始</th><th>退職日</th><th class="num">順</th><th>有効</th><th>状態</th><th></th>
  </tr></thead>
  <tbody>
<%
Set rs = DbQuery("SELECT * FROM [M_担当者] ORDER BY [表示順]", Empty)
Do While Not rs.EOF
  Dim zaiseki
  zaiseki = (IsNull(rs("在籍終了日")) Or CDate(rs("在籍終了日")) >= today) And rs("有効")
%>
<% Dim fid : fid = "opf" & rs("担当者ID") %>
    <tr>
      <td><input form="<%= fid %>" type="text" name="code" size="4" value="<%= H(rs("担当者コード")) %>"></td>
      <td><input form="<%= fid %>" type="text" name="sei" size="6" value="<%= H(rs("姓")) %>"></td>
      <td><input form="<%= fid %>" type="text" name="mei" size="6" value="<%= H(rs("名")) %>"></td>
      <td><input form="<%= fid %>" type="text" name="kana" size="8" value="<%= H(rs("カナ")) %>"></td>
      <td><select form="<%= fid %>" name="kbn">
        <option<%= IIfS(rs("職員区分")="パート"," selected","") %>>パート</option>
        <option<%= IIfS(rs("職員区分")="職員"," selected","") %>>職員</option>
      </select></td>
      <td><input form="<%= fid %>" type="date" name="start" value="<%= Ymd(rs("在籍開始日")) %>"></td>
      <td><input form="<%= fid %>" type="date" name="end" value="<%= Ymd(rs("在籍終了日")) %>"></td>
      <td class="num"><input form="<%= fid %>" type="number" name="ord" value="<%= H(rs("表示順")) %>" style="width:4em"></td>
      <td><input form="<%= fid %>" type="checkbox" name="act2" value="on"<%= IIfS(rs("有効")," checked","") %>></td>
      <td><%= IIfS(zaiseki, "<span class=""badge fix"">在籍</span>", "<span class=""badge wip"">退職</span>") %></td>
      <td>
        <form id="<%= fid %>" method="post" action="master.asp">
          <input type="hidden" name="tab" value="operator">
          <input type="hidden" name="act" value="op_upd">
          <input type="hidden" name="id" value="<%= rs("担当者ID") %>">
          <button class="btn ghost" type="submit">保存</button>
        </form>
      </td>
    </tr>
<%
  If Len("" & rs("備考")) > 0 Then
%>
    <tr><td colspan="11" style="background:#fdf3e0; font-size:13px">
      <b>移行メモ：</b><%= H(rs("備考")) %></td></tr>
<%
  End If
  rs.MoveNext
Loop
rs.Close
%>
  </tbody>
</table>
</div>

<% ElseIf tab_ = "product" Then %>
<div class="panel">
  <h2 style="margin-top:0">製品を追加する</h2>
  <form method="post" action="master.asp">
    <input type="hidden" name="tab" value="product">
    <input type="hidden" name="act" value="pr_add">
    <div class="row">
      <div class="field"><label>ブロック</label>
        <select name="block">
<%
Set rs = DbQuery("SELECT [ブロックID],[ブロック名] FROM [M_ブロック] WHERE [製品別]=True ORDER BY [表示順]", Empty)
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("ブロックID") & """>" & H(rs("ブロック名")) & "</option>"
    rs.MoveNext
Loop
rs.Close
%>
        </select></div>
      <div class="field" style="flex:1 1 320px"><label>製品名</label>
        <input type="text" name="name" required style="width:100%"></div>
      <div class="field"><label>販売開始日</label>
        <input type="date" name="start" value="<%= Ymd(today) %>"></div>
      <div class="field"><label>表示順</label>
        <input type="number" name="ord" value="99" style="width:5em"></div>
      <div><button class="btn" type="submit">追加</button></div>
    </div>
  </form>
</div>

<h2>製品一覧</h2>
<p class="lead">販売が終わった製品は「終了日」を入れると入力候補から消えます。過去データは残ります。</p>
<div class="table-wrap">
<table>
  <thead><tr><th>ID</th><th>ブロック</th><th>製品名</th><th>開始</th><th>終了</th>
    <th class="num">順</th><th>有効</th><th></th></tr></thead>
  <tbody>
<%
Set rs = DbQuery( _
  "SELECT PR.*, BK.[ブロック名] FROM [M_製品] AS PR " & _
  "INNER JOIN [M_ブロック] AS BK ON PR.[ブロックID]=BK.[ブロックID] " & _
  "ORDER BY PR.[ブロックID], PR.[表示順]", Empty)
Do While Not rs.EOF
%>
<% Dim pfid : pfid = "prf" & rs("製品ID") %>
    <tr>
      <td><%= H(rs("製品ID")) %></td>
      <td><%= H(rs("ブロック名")) %></td>
      <td><input form="<%= pfid %>" type="text" name="name" value="<%= H(rs("製品名")) %>" style="width:22em"></td>
      <td><input form="<%= pfid %>" type="date" name="start" value="<%= Ymd(rs("適用開始日")) %>"></td>
      <td><input form="<%= pfid %>" type="date" name="end" value="<%= Ymd(rs("適用終了日")) %>"></td>
      <td class="num"><input form="<%= pfid %>" type="number" name="ord" value="<%= H(rs("表示順")) %>" style="width:4em"></td>
      <td><input form="<%= pfid %>" type="checkbox" name="on" value="1"<%= IIfS(rs("有効")," checked","") %>></td>
      <td>
        <form id="<%= pfid %>" method="post" action="master.asp">
          <input type="hidden" name="tab" value="product">
          <input type="hidden" name="act" value="pr_upd">
          <input type="hidden" name="id" value="<%= rs("製品ID") %>">
          <button class="btn ghost" type="submit">保存</button>
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

<% ElseIf tab_ = "kubun" Then %>
<h2>区分一覧</h2>
<p class="lead">
  <b>集計列</b>が、その区分を集計表のどの列に積むかを決めます
  （3=申込方法 4=抽選結果 5=納付書発送 6=商品発送 7=その他 8=商品交換）。
  現行 Excel では Sheet2 の 2 行目に隠れていた設定です。
  <b>旧転記名</b>は移行時の照合用で、現行 Excel が集計表に書いていた文言です。
</p>
<div class="table-wrap">
<table>
  <thead><tr><th>ID</th><th>ブロック</th><th>区分名</th><th>集計列</th><th>内訳</th>
    <th class="num">順</th><th>有効</th><th>旧転記名</th><th></th></tr></thead>
  <tbody>
<%
Set rs = DbQuery( _
  "SELECT KB.*, BK.[ブロック名] FROM [M_区分] AS KB " & _
  "INNER JOIN [M_ブロック] AS BK ON KB.[ブロックID]=BK.[ブロックID] " & _
  "ORDER BY KB.[ブロックID], KB.[表示順]", Empty)
Do While Not rs.EOF
%>
<% Dim kfid : kfid = "kbf" & rs("区分ID") %>
    <tr>
      <td><%= H(rs("区分ID")) %></td>
      <td><%= H(rs("ブロック名")) %></td>
      <td><input form="<%= kfid %>" type="text" name="name" value="<%= H(rs("区分名")) %>" style="width:18em"></td>
      <td><select form="<%= kfid %>" name="col">
<%
      Dim c
      For c = 3 To 8
          Response.Write "<option value=""" & c & """" & _
              IIfS(CLng(rs("集計列ID")) = c, " selected", "") & ">" & c & ": " & _
              H(DbScalar("SELECT [集計列名] FROM [M_集計列] WHERE [集計列ID]=?", Array(c), "")) & _
              "</option>"
      Next
%>
      </select></td>
      <td><input form="<%= kfid %>" type="text" name="uchi" value="<%= H(rs("内訳区分")) %>" size="6"></td>
      <td class="num"><input form="<%= kfid %>" type="number" name="ord" value="<%= H(rs("表示順")) %>" style="width:4em"></td>
      <td><input form="<%= kfid %>" type="checkbox" name="on" value="1"<%= IIfS(rs("有効")," checked","") %>></td>
      <td style="font-size:13px;color:#5b6b7c"><%= H(rs("旧転記名")) %></td>
      <td>
        <form id="<%= kfid %>" method="post" action="master.asp">
          <input type="hidden" name="tab" value="kubun">
          <input type="hidden" name="act" value="kb_upd">
          <input type="hidden" name="id" value="<%= rs("区分ID") %>">
          <button class="btn ghost" type="submit">保存</button>
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

<% Else %>
<h2>業務項目（電話応対以外の業務）</h2>
<p class="lead">帳票の ①〜⑬ の欄です。「帳票表示名」がそのまま印刷されます。</p>
<div class="table-wrap">
<table>
  <thead><tr><th>ID</th><th>番号</th><th>項目名</th><th>帳票表示名</th>
    <th class="num">順</th><th>有効</th><th></th></tr></thead>
  <tbody>
<%
Set rs = DbQuery("SELECT * FROM [M_業務項目] ORDER BY [表示順]", Empty)
Do While Not rs.EOF
%>
<% Dim tfid : tfid = "tkf" & rs("業務項目ID") %>
    <tr>
      <td><%= H(rs("業務項目ID")) %></td>
      <td><%= H(rs("番号")) %></td>
      <td><input form="<%= tfid %>" type="text" name="name" value="<%= H(rs("項目名")) %>" style="width:16em"></td>
      <td><input form="<%= tfid %>" type="text" name="rname" value="<%= H(rs("帳票表示名")) %>" style="width:20em"></td>
      <td class="num"><input form="<%= tfid %>" type="number" name="ord" value="<%= H(rs("表示順")) %>" style="width:4em"></td>
      <td><input form="<%= tfid %>" type="checkbox" name="on" value="1"<%= IIfS(rs("有効")," checked","") %>></td>
      <td>
        <form id="<%= tfid %>" method="post" action="master.asp">
          <input type="hidden" name="tab" value="task">
          <input type="hidden" name="act" value="tk_upd">
          <input type="hidden" name="id" value="<%= rs("業務項目ID") %>">
          <button class="btn ghost" type="submit">保存</button>
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
<% End If %>
<% PageFoot() : CloseDb %>
