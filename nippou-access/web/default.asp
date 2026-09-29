<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' -----------------------------------------------------------------------------
'  パート職員のかたの入口
'
'  毎日使うのは受付入力とその他業務の 2 つだけなので、それだけを大きく出す。
'  日報や集計表は職員のかたの担当なので、ここには出さない (staff.asp が入口)。
' -----------------------------------------------------------------------------
 ' 職員用フォルダなら、職員用メニューを出す (config.asp の ROLE が "staff")
If RoleSet() And IsStaff() Then
    Response.Redirect "staff.asp"
End If

Dim today
today = Date()

PageHead "メニュー", "default.asp"
%>
<h1>受付入力メニュー</h1>
<p class="lead">
  <%= H(YmdSlash(today)) %>（<%= H(Youbi(today)) %>）　
  下のどちらかを押してください。入力したものはその場で保存されます。
</p>

<div class="tiles">
  <a class="tile" href="entry.asp"><b>① 受付を入力する</b>
    <span>電話でお受けした問合せの件数を入れます。<br>
          毎日いちばん使う画面です。</span></a>
  <a class="tile" href="tasks.asp"><b>② その他業務を入力する</b>
    <span>受注入力や架電など、電話を受ける以外の作業の件数を入れます。<br>
          1 日の終わりにまとめて入れてください。</span></a>
</div>

<div class="panel" style="margin-top:22px">
  <h2 style="margin-top:0">はじめてお使いのかたへ</h2>
  <ul>
    <li>まず<b>自分の名前</b>を選んでから入力してください。名前を選んだ時点で、その日の出勤者として登録されます。</li>
    <li>「転記」や「更新」のボタンを押す必要はありません。<b>入れたらすぐ反映されます。</b></li>
    <li>間違えたら、その行の「削除」で消せます。</li>
    <li>昨日の分を入れ忘れたときは、<b>対象日を昨日に変えれば</b>あとからでも入れられます。</li>
  </ul>
  <p class="lead" style="margin-bottom:0">
    自分の名前が一覧に無いときや、入れたい区分が無いときは、職員のかたにご連絡ください。
  </p>
</div>

<% If IsStaff() Then %>
<div class="panel" style="margin-top:18px">
  <p style="margin:0">
    日報の作成・集計表・マスタ保守は職員用のメニューにあります。
    <a class="btn" href="staff.asp" style="margin-left:10px">職員用メニューへ</a>
  </p>
</div>
<% End If %>
<% PageFoot() : CloseDb %>
