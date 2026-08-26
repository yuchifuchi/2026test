<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/layout.asp"-->
<%
' -----------------------------------------------------------------------------
'  印刷用。現行「日報集計印刷用フォーム.xlsm ＞ 印刷用」シートの体裁を再現する。
'  値の出どころは T_受電 ただ 1 つ。Excel のような二重経路が無いので、
'  この紙とデスクトップ版・集計表の数値は必ず一致する。
' -----------------------------------------------------------------------------
Dim dt, hd, rs, names(14), i, tasks, tk

dt = ParamDate("d", Date())
Ensure日報 dt
Set hd = DbQuery("SELECT * FROM [Q_日報_ヘッダ] WHERE [対象日]=?", Array(dt))

' 出勤者を 2 列 × 7 行に流し込む (現行 F10:O16 と同じ枠)
i = 0
Set rs = DbQuery("SELECT [氏名] FROM [Q_日報_出勤] WHERE [対象日]=? ORDER BY [表示順]", Array(dt))
Do While Not rs.EOF And i < 14
    names(i) = rs("氏名") : i = i + 1
    rs.MoveNext
Loop
rs.Close

PageHead "日報 印刷", ""
%>
<div class="printbar noprint">
  <button class="btn" onclick="window.print()">印刷する</button>
  <a class="btn ghost" href="daily.asp?d=<%= Ymd(dt) %>">日報の編集に戻る</a>
  <input type="date" value="<%= Ymd(dt) %>"
         onchange="location.href='report.asp?d='+this.value">
</div>

<div class="sheet">
  <div class="kanai">【課内限り】</div>
  <div class="topright"><%= H(Wareki(dt)) %></div>
  <h1 class="title">電話応対報告書日報集計表</h1>

  <table>
    <tr>
      <th style="width:22mm">出勤者</th>
      <td style="width:14mm; text-align:center"><%= H(hd("出勤者数")) %></td>
      <td style="width:10mm">名</td>
      <th style="width:52mm">氏名</th>
      <th style="width:52mm">氏名</th>
      <th>備考</th>
    </tr>
<% For i = 0 To 6 %>
    <tr>
      <td class="noborder" style="border:none"></td>
      <td class="noborder" style="border:none"></td>
      <td class="noborder" style="border:none"></td>
      <td><%= H(names(i)) %>&nbsp;</td>
      <td><%= H(names(i + 7)) %>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
<% Next %>
  </table>

  <p style="margin:8px 0">（　<b><%= H(hd("回線数")) %></b>　回線 ）</p>

  <table>
    <tr>
      <th rowspan="2" style="width:26mm">問合せ件数<br><span style="font-size:8pt">（　）内は職員受電数</span></th>
      <td rowspan="2" class="big" style="width:24mm"><%= H(hd("合計")) %> 件</td>
      <th style="text-align:center">申込</th>
      <th style="text-align:center">抽選</th>
      <th style="text-align:center">払込用紙</th>
      <th style="text-align:center">商品発送</th>
      <th style="text-align:center">その他</th>
    </tr>
    <tr>
      <td class="big"><%= H(hd("申込")) %><div class="sub">(<%= H(hd("申込_職員")) %>)</div></td>
      <td class="big"><%= H(hd("抽選")) %><div class="sub">(<%= H(hd("抽選_職員")) %>)</div></td>
      <td class="big"><%= H(hd("払込用紙")) %><div class="sub">(<%= H(hd("払込用紙_職員")) %>)</div></td>
      <td class="big"><%= H(hd("商品発送")) %><div class="sub">(<%= H(hd("商品発送_職員")) %>)</div></td>
      <td class="big"><%= H(hd("その他")) %><div class="sub">(<%= H(hd("その他_職員")) %>)</div></td>
    </tr>
  </table>
  <p style="text-align:right; margin:6px 0">
    内　　交換 <b><%= H(hd("内交換")) %></b> 件　／　内　　返金 <b><%= H(hd("内返金")) %></b> 件
  </p>

  <div class="cap">特記事項（報告書の電話件数だけでは伝わり難い事項など）</div>
  <div class="memo"><%= H(hd("特記事項")) %></div>

  <div class="cap">職員に代わった案件（概要）</div>
  <div class="memo"><%= H(hd("職員代替案件")) %></div>

  <div class="cap">要望（お客様からの問い合わせを減らすための改善提案等）</div>
  <div class="memo"><%= H(hd("要望")) %></div>

  <div class="cap">電話対応以外の業務（データ入力作業等）</div>
  <table>
<%
' ①〜⑬ を左 8 件 / 右 5 件に割る (現行の並びと同じ)
Dim leftN(7), leftV(7), rightN(4), rightV(4), n
n = 0
Set tasks = DbQuery( _
  "SELECT TM.[業務項目ID],TM.[帳票表示名], " & _
  " (SELECT Nz(Sum(W.[件数]),0) FROM [T_業務実績] AS W " & _
  "   WHERE W.[対象日]=? AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] " & _
  "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]", Array(dt))
Do While Not tasks.EOF
    If n < 8 Then
        leftN(n) = tasks("帳票表示名") : leftV(n) = tasks("件数")
    ElseIf n < 13 Then
        rightN(n - 8) = tasks("帳票表示名") : rightV(n - 8) = tasks("件数")
    End If
    n = n + 1
    tasks.MoveNext
Loop
tasks.Close

For i = 0 To 7
%>
    <tr>
      <td style="width:46mm"><%= H(leftN(i)) %>&nbsp;</td>
      <td style="width:14mm; text-align:right"><%= H(leftV(i)) %></td>
      <td style="width:8mm"><%= IIfS(Len("" & leftN(i)) > 0, "件", "") %></td>
      <td style="width:46mm"><%= IIfS(i <= 4, H(rightN(i)), "") %>&nbsp;</td>
      <td style="width:14mm; text-align:right"><%= IIfS(i <= 4, H(rightV(i)), "") %></td>
      <td><%= IIfS(i <= 4 And Len("" & rightN(i)) > 0, "件", "") %></td>
    </tr>
<% Next %>
  </table>
</div>
<% hd.Close : PageFoot() : CloseDb %>
