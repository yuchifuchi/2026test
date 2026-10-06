<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' 職員メニュー
Dim today, d2, daily, state, persons
RequireStaff
today = Date()
d2 = DateAdd("d", 1, today)
daily = DbQuery(SqlDaily(), Array(today, d2))
state = "まだ作っていません"
If UBound(daily) >= 0 Then state = ToStr(daily(0)("状態"))
persons = DbQuery(SqlPersonTotals(), Array(today, d2))
PageHead "職員メニュー"
%>
<div class="totals">今日（<%= DateJa(today) %>）の様子：受付入力をした人 <b><%= UBound(persons) + 1 %></b> 人 ／ 日報 <b><%= H(state) %></b></div>
<div class="menu menu-cols">
  <div class="menu-col">
    <h2>日報</h2>
    <a class="btn" href="daily.asp">日報の作成・確定<small>出勤者・回線数・記述欄</small></a>
    <a class="btn btn-sub" href="report.asp">帳票を見る・PDF にする<small>A4 縦 1 枚</small></a>
    <a class="btn btn-sub" href="check.asp">入力もれチェック<small>入れ忘れている人がいないか</small></a>
  </div>
  <div class="menu-col">
    <h2>集計</h2>
    <a class="btn" href="summary.asp">週の集計表<small>月曜～日曜</small></a>
    <h2>代わりに入力する</h2>
    <a class="btn btn-sub" href="entry.asp">受付入力<small>電話の件数</small></a>
    <a class="btn btn-sub" href="tasks.asp">その他業務<small>①～⑬</small></a>
  </div>
  <div class="menu-col">
    <h2>設定</h2>
    <a class="btn btn-sub" href="master.asp">マスタ保守<small>担当者・区分・製品・業務項目</small></a>
    <a class="btn btn-sub" href="setup_check.asp">設置チェック<small>うまく動かないとき</small></a>
  </div>
</div>
<% PageFoot %>
