<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>エラーが発生しました</title>
<link rel="stylesheet" href="css/style.css">
</head><body>
<header class="appbar"><a class="brand" href="default.asp">電話応対日報 集計システム</a></header>
<main>
  <h1>画面を表示できませんでした</h1>
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
