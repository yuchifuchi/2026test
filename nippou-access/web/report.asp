<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/sheet.asp"-->
<%
RequireStaff                              ' 職員以外はここで止まる

' -----------------------------------------------------------------------------
'  印刷用。現行「日報集計印刷用フォーム.xlsm ＞ 印刷用」シートの体裁を再現する。
'  値の出どころは T_受電 ただ 1 つ。Excel のような二重経路が無いので、
'  この紙とデスクトップ版・集計表の数値は必ず一致する。
' -----------------------------------------------------------------------------
Dim dt

dt = ParamDate("d", Date())

PageHead "日報 印刷", ""
%>
<div class="printbar noprint">
  <a class="btn" href="printpdf.asp?d=<%= Ymd(dt) %>" target="_blank"
     rel="noopener">PDF で印刷する</a>
  <button class="btn ghost" onclick="window.print()">ブラウザで印刷する</button>
  <a class="btn ghost" href="daily.asp?d=<%= Ymd(dt) %>">日報の編集に戻る</a>
  <input type="date" value="<%= Ymd(dt) %>"
         onchange="location.href='report.asp?d='+this.value">
  <p class="legend">
    <b>「PDF で印刷する」がおすすめです。</b>
    別のタブに PDF が開くので、そこで印刷してください。
    日付や URL が紙に入りません。<br>
    「ブラウザで印刷する」でも出せますが、
    印刷設定で「ヘッダーとフッター」を外し、倍率を 100% にしてください。
  </p>
</div>

<%= SheetHtml(dt) %>
<% PageFoot() : CloseDb %>
