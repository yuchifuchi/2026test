<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
RequireStaff                              ' パート職員はここで止まる

' -----------------------------------------------------------------------------
'  職員のかたの入口
'
'  日報の作成・確定、帳票の印刷、集計表、入力もれの確認、マスタ保守。
'  パート職員の画面 (受付入力・その他業務) も、代わりに入力できるよう置いてある。
' -----------------------------------------------------------------------------
Dim today, cntToday, missing, state
today = Date()
cntToday = DbScalar("SELECT Sum([件数]) FROM [T_受電] WHERE [対象日]=?", Array(today), 0)
missing  = DbScalar("SELECT Count(*) FROM (" & SQL_未入力チェック() & ") AS C " & _
                    "WHERE C.[対象日]=?", Array(today), 0)
state    = DbScalar("SELECT [状態] FROM [T_日報] WHERE [対象日]=?", Array(today), "未作成")

PageHead "メニュー", "staff.asp"
%>
<h1>職員用メニュー</h1>
<p class="lead">
  日報・集計表・マスタ保守はこちらです。
  受付の入力と日報は同じデータから作られるので、突き合わせは要りません。
</p>

<div class="grid g3" style="margin-bottom:24px">
  <div class="panel">
    <div style="font-size:12px;color:#5b6675;letter-spacing:.08em">本日の受付件数</div>
    <div style="font-size:31px;font-weight:700;line-height:1.25"><%= H(cntToday) %><span style="font-size:14px;font-weight:400;color:#5b6675;margin-left:3px">件</span></div>
    <div style="font-size:12px;color:#5b6675"><%= H(YmdSlash(today)) %>（<%= H(Youbi(today)) %>）</div>
  </div>
  <div class="panel">
    <div style="font-size:12px;color:#5b6675;letter-spacing:.08em">本日の日報</div>
    <div style="font-size:22px;font-weight:700;line-height:1.6">
      <% If state = "確定" Then %><span class="badge fix">確定済み</span>
      <% ElseIf state = "未作成" Then %><span class="badge wip">未作成</span>
      <% Else %><span class="badge wip">入力中</span><% End If %>
    </div>
    <div style="font-size:12px;color:#5b6675">夕方に確定してください</div>
  </div>
  <div class="panel">
    <div style="font-size:12px;color:#5b6675;letter-spacing:.08em">入力もれ</div>
    <div style="font-size:31px;font-weight:700;line-height:1.25<% If missing > 0 Then %>;color:#8a5a12<% End If %>"><%= H(missing) %><span style="font-size:14px;font-weight:400;color:#5b6675;margin-left:3px">名</span></div>
    <div style="font-size:12px;color:#5b6675">出勤登録があって実績が無い人</div>
  </div>
</div>

<% If missing > 0 Then %>
  <p class="notice warn">
    実績が 1 件も入っていない担当者が <%= H(missing) %> 名います。
    確定の前に確認してください。
    <a href="check.asp?d=<%= Ymd(today) %>">入力もれチェックを開く</a>
  </p>
<% End If %>

<h2>日々の作業</h2>
<div class="tiles">
  <a class="tile" href="daily.asp"><b>① 日報を作る・確定する</b>
    <span>出勤者・回線数・特記事項を入れて確定します。件数欄は自動集計です。</span></a>
  <a class="tile" href="report.asp"><b>② 日報を印刷する</b>
    <span>現行の「印刷用」シートと同じ体裁。A4 縦 1 枚で出ます。</span></a>
  <a class="tile" href="check.asp"><b>③ 入力もれをチェックする</b>
    <span>実績が 1 件も無い担当者を洗い出します。確定の前にご確認ください。</span></a>
</div>

<h2>集計・確認</h2>
<div class="tiles">
  <a class="tile" href="summary.asp"><b>集計表を見る・印刷する</b>
    <span>期間を指定して日別の集計と明細を表示します。日報と必ず一致します。</span></a>
</div>

<h2>保守</h2>
<div class="tiles">
  <a class="tile" href="master.asp"><b>マスタ保守</b>
    <span>担当者の入退職、製品・区分の追加や修正。
          実績で使われているものは削除できない安全弁が入っています。</span></a>
</div>

<h2>パート職員の画面（代わりに入力するとき）</h2>
<div class="tiles">
  <a class="tile" href="entry.asp"><b>受付入力</b>
    <span>問合せの件数を入れます。</span></a>
  <a class="tile" href="tasks.asp"><b>その他業務</b>
    <span>受注入力や架電などの件数を入れます。</span></a>
</div>
<% PageFoot() : CloseDb %>
