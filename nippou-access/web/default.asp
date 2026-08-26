<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Session.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
Dim today, cntToday, missing
today = Date()
cntToday = DbScalar("SELECT Sum([件数]) FROM [T_受電] WHERE [対象日]=?", Array(today), 0)
missing  = DbScalar("SELECT Count(*) FROM [Q_未入力チェック] WHERE [対象日]=?", Array(today), 0)

PageHead "メニュー", "default.asp"
%>
<h1>メニュー</h1>
<p class="lead">
  受付の記録・日報・集計表は、すべて同じ 1 つのデータから作られます。
  日報と集計表の数値が食い違うことはありません。
</p>

<div class="panel">
  <div class="row">
    <div><b><%= H(YmdSlash(today)) %>（<%= H(Youbi(today)) %>）</b> の受付件数：
      <b style="font-size:20px"><%= H(cntToday) %></b> 件</div>
  </div>
<% If missing > 0 Then %>
  <p class="notice warn" style="margin-top:12px">
    出勤登録があるのに実績が 1 件も入っていない担当者が <%= H(missing) %> 名います。
    <a href="check.asp?d=<%= Ymd(today) %>">入力もれチェックを開く</a>
  </p>
<% End If %>
</div>

<div class="tiles">
  <a class="tile" href="entry.asp"><b>① 受付を入力する</b>
    <span>担当者ごとに、その日の問合せ件数を登録します。</span></a>
  <a class="tile" href="tasks.asp"><b>② その他業務を入力する</b>
    <span>受注入力や架電など、電話を受ける以外の作業の件数を入れます。</span></a>
  <a class="tile" href="daily.asp"><b>③ 日報を作る・確認する</b>
    <span>出勤者・回線数・特記事項を入れ、集計値を確認して確定します。</span></a>
  <a class="tile" href="report.asp"><b>④ 日報を印刷する</b>
    <span>現行の「印刷用」シートと同じ体裁で印刷します。</span></a>
  <a class="tile" href="summary.asp"><b>⑤ 集計表を見る</b>
    <span>期間を指定して日別の集計と明細を表示します。</span></a>
  <a class="tile" href="check.asp"><b>入力もれチェック</b>
    <span>実績が 1 件も無い担当者を洗い出します。</span></a>
<% If IsAdmin() Then %>
  <a class="tile" href="master.asp"><b>マスタ保守</b>
    <span>担当者の入退職、製品・区分の追加はここで行います。担当者専用です。</span></a>
<% End If %>
</div>
<% PageFoot() : CloseDb %>
