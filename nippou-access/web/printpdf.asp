<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/sheet.asp"-->
<!--#include file="include/pdf.asp"-->
<%
' -----------------------------------------------------------------------------
'  帳票を PDF にして返す
'
'  「PDF で印刷する」を押すとここに来る。別タブに PDF が開くので、
'  利用者はいつもどおり Ctrl+P で印刷する。
'  ブラウザが勝手に入れる日付や URL は紙に載らない。
' -----------------------------------------------------------------------------
Dim dt, css, html, pdfPath

RequireStaff                              ' 職員以外はここで止まる

dt = ParamDate("d", Date())

' 帳票のスタイルは画面と同じものを使う。二重管理にしない。
' .css は UTF-8 なので ADODB.Stream で読む (FileSystemObject では化ける)。
css = ReadTextUtf8(Server.MapPath("css/style.css"))

' 変換プログラムに渡す、単独で完結した HTML を組み立てる。
' 画面まわりは入れない。@media print ではなく、最初から帳票だけの紙にする。
html = "<!doctype html><html lang=""ja""><head><meta charset=""utf-8"">" & _
       "<title>電話応対報告書日報集計表</title><style>" & css & _
       " body{background:#fff;margin:0;padding:0}" & _
       " .sheet{width:auto;margin:0;padding:0;border:0;border-radius:0;box-shadow:none}" & _
       "</style></head><body>" & SheetHtml(dt) & "</body></html>"

pdfPath = PdfFromHtml(html)

If Len(pdfPath) > 0 Then
    PdfSendAndDelete pdfPath, "日報_" & Replace(Ymd(dt), "-", "") & ".pdf"
    CloseDb
    Response.End
End If

' --- ここから下は、変換できなかったときだけ表示される ---
CloseDb
PageHead "PDF を作れませんでした", ""
%>
<h1>PDF を作れませんでした</h1>
<p class="notice err">
  帳票は表示できますが、PDF への変換ができませんでした。
  下の「ブラウザで印刷する」からは印刷できます。
</p>

<div class="panel">
  <h2 style="margin-top:0">いますぐ印刷したいとき</h2>
  <p>
    <a class="btn" href="report.asp?d=<%= Ymd(dt) %>">帳票を開いてブラウザで印刷する</a>
  </p>
  <p class="lead" style="margin-top:12px">
    そのときは印刷設定で「<b>ヘッダーとフッター</b>」のチェックを外し、
    倍率を <b>100%</b> にしてください。一度設定すれば次回から覚えています。
  </p>
</div>

<div class="panel">
  <h2 style="margin-top:0">担当者のかたへ</h2>
  <p class="lead">次の内容を情報システム担当にお伝えください。</p>
  <pre style="white-space:pre-wrap; background:#f4f2ee; padding:12px 14px;
              border-radius:6px; font-size:13px"><%= H(gPdfError) %></pre>
  <p class="lead">
    設定は <code>include\pdf.asp</code> の先頭にあります。
    <b>PDF_ENGINE</b>（edge か wkhtmltopdf）、<b>PDF_EDGE_EXE</b>、
    <b>PDF_WORK_DIR</b> を確認してください。
  </p>
</div>
<% PageFoot() %>
