<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>エラーが発生しました</title>
<link rel="stylesheet" href="css/style.css">
</head><body>
<header class="appbar"><a class="brand" href="default.asp">電話応対日報 集計システム</a></header>
<main>
  <h1>画面を表示できませんでした</h1>
<%
' ---------------------------------------------------------------------------
'  ASP がエラーで止まったとき、その中身をここに出す。
'
'  IIS の「ASP → ブラウザーにエラーを送信する」を True にしなくても、
'  この画面に出るようにしてある (web.config で 500.100 をここへ回している)。
'  原因が分かったら、この枠の内容をそのまま担当者にお伝えください。
' ---------------------------------------------------------------------------
Dim e, hasDetail
hasDetail = False
On Error Resume Next
Set e = Server.GetLastError()
If Err.Number = 0 Then
    If Not (e Is Nothing) Then
        If Len("" & e.Description) > 0 Or Len("" & e.File) > 0 Then hasDetail = True
    End If
End If
Err.Clear
On Error GoTo 0

If hasDetail Then
%>
  <div class="panel" style="border-left:5px solid #a32020">
    <h2 style="margin-top:0">エラーの内容（担当者のかたへ）</h2>
    <table style="font-size:14px">
      <tr><th style="width:9em">ファイル</th><td><%= Server.HTMLEncode("" & e.File) %></td></tr>
      <tr><th>行</th><td><%= Server.HTMLEncode("" & e.Line) %></td></tr>
      <tr><th>内容</th><td><b><%= Server.HTMLEncode("" & e.Description) %></b></td></tr>
      <tr><th>種別</th><td><%= Server.HTMLEncode("" & e.Category) %>
          <%= Server.HTMLEncode("" & e.Number) %></td></tr>
      <tr><th>場所</th><td><%= Server.HTMLEncode("" & Request.ServerVariables("URL")) %></td></tr>
    </table>
    <p class="lead" style="margin-bottom:0">
      この 5 行をそのままお知らせいただければ、原因が分かります。
    </p>
  </div>
<%
End If
%>
  <p class="notice err">
    処理の途中で問題が起きました。入力した内容は保存されていない可能性があります。
  </p>

  <div class="panel">
    <h2 style="margin-top:0">まず試すこと</h2>
    <ol>
      <li>ブラウザの「戻る」で前の画面に戻り、もう一度やり直してください。</li>
      <li>それでも直らないときは、いったんメニューに戻ってから開き直してください。</li>
    </ol>
    <p style="margin-top:16px">
      <a class="btn" href="default.asp">メニューに戻る</a>
    </p>
  </div>

  <div class="panel">
    <h2 style="margin-top:0">担当者に連絡するとき</h2>
    <p>次のことを伝えていただけると、原因がすぐ分かります。</p>
    <ul>
      <li>何をしようとしていたか（例：8月25日の受付を登録しようとした）</li>
      <li>どの画面で起きたか（例：受付入力）</li>
      <li>起きた時刻</li>
    </ul>
    <p class="lead">
      くわしいエラーの内容はサーバーのログに記録されています。
      画面には出さない設定にしてあります。
    </p>
  </div>
</main>
<footer>電話応対日報 集計システム</footer>
</body></html>
