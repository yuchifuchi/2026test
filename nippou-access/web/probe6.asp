<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!doctype html><html lang="ja"><head><meta charset="utf-8">
<title>切り分け 6</title>
<style>body{font-family:"Meiryo","Yu Gothic UI",system-ui,sans-serif;margin:0;background:#f4f6f8;color:#1b2733}
main{max-width:760px;margin:0 auto;padding:28px 20px}
h1{font-size:20px} .ok{color:#1c7c4a;font-weight:700;font-size:18px}
p{line-height:1.9} code{background:#eef2f6;padding:2px 6px;border-radius:4px}
a{color:#155a8f}</style></head><body><main>
<h1>切り分け 6 ― 役割の判定 (auth.asp)</h1>
<p class="ok">OK ― include\auth.asp を読み込めました。</p>
<p>いまのログオン名：<code><%= Server.HTMLEncode(CurrentUser()) %></code></p>
<p>職員として扱われるか：<b><%= CStr(IsStaff()) %></b>（True＝職員）</p>
<% CloseDb %>
<p style="margin-top:26px"><a href="setup_check.asp">設置チェックに戻る</a></p>
</main></body></html>
