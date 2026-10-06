<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<%
' ============================================================
'  その他業務：電話以外の仕事（①～⑬）の件数を、1 人・1 日ぶん入れる。
'  保存の考え方は受付入力と同じ（今の件数で上書き。0 は行を消す）。
' ============================================================
Dim d, tid, msgNg, saved, items, people, existing, posted, bad, parsed, dd, cookieVal

d = PageDate()
tid = ParamLong(QueryVal("t"), 0)
If tid = 0 Then
    cookieVal = Request.Cookies("nippou_t")
    tid = ParamLong(cookieVal, 0)
End If
msgNg = ""
saved = (QueryVal("saved") = "1")
Set posted = Server.CreateObject("Scripting.Dictionary")
Set bad = Server.CreateObject("Scripting.Dictionary")
Set parsed = Server.CreateObject("Scripting.Dictionary")
items = DbQuery(SqlTaskItemsAll(), Array())

If IsPost() Then
    If FormVal("act") = "go" Then
        If ParseYMD(FormVal("d"), dd) Then
            Go "tasks.asp?d=" & YMD(dd) & "&t=" & ParamLong(FormVal("t"), 0)
        End If
        msgNg = "日付が読めません。2026-08-25 のように入れてください。"
    ElseIf FormVal("act") = "save" Then
        SaveTasks
    End If
End If

Set existing = LoadTaskRows(d, tid)
people = DbQuery(SqlStaffAll(), Array())

Sub SaveTasks()
    Dim t, ex, j, key, raw, n, errCount
    If Not ParseYMD(FormVal("d"), dd) Then
        msgNg = "日付が読めません。"
        Exit Sub
    End If
    d = dd
    t = ParamLong(FormVal("t"), 0)
    tid = t
    If t = 0 Then
        msgNg = "担当者を選んでから保存してください。"
        Exit Sub
    End If
    If IsLocked(d) Then
        msgNg = "この日の日報は確定済みなので、保存できませんでした。直すときは、職員に「確定の取り消し」を頼んでください。"
        Exit Sub
    End If
    Set ex = LoadTaskRows(d, t)
    errCount = 0
    For j = 0 To UBound(items)
        key = CStr(ToLong(items(j)("業務項目ID")))
        If items(j)("有効") Or ex.Exists(key) Then
            raw = Request.Form("g_" & key)
            If Not IsEmpty(raw) Then
                posted(key) = CStr(raw)
                If ParseCount(CStr(raw), n) Then
                    parsed(key) = n
                Else
                    bad(key) = True
                    errCount = errCount + 1
                End If
            End If
        End If
    Next
    If errCount > 0 Then
        msgNg = "赤い欄に、数字ではないもの（または 5 けた以上の数）が入っています。直してから、もう一度「保存する」を押してください。（まだ保存していません）"
        Exit Sub
    End If
    DbBegin
    For Each key In parsed.Keys
        n = parsed(key)
        If ex.Exists(key) Then
            If n = 0 Then
                DbExec SqlTaskDelete(), Array(ex(key)(0))
            ElseIf n <> ex(key)(1) Then
                DbExec SqlTaskUpdate(), Array(n, ex(key)(0))
            End If
        ElseIf n > 0 Then
            DbExec SqlTaskInsert(), Array(d, t, CLng(key), n)
        End If
    Next
    DbCommit
    Response.Cookies("nippou_t") = CStr(t)
    Response.Cookies("nippou_t").Expires = DateAdd("d", 400, Date())
    Response.Cookies("nippou_t").Path = "/"
    Go "tasks.asp?d=" & YMD(d) & "&t=" & t & "&saved=1"
End Sub

' "業務項目ID" → Array(実績ID, 件数)
Function LoadTaskRows(day, t)
    Dim rows, j, dict
    Set dict = Server.CreateObject("Scripting.Dictionary")
    If t <> 0 Then
        rows = DbQuery(SqlTaskRowsOnePerson(), Array(day, DateAdd("d", 1, day), t))
        For j = 0 To UBound(rows)
            dict.Add CStr(ToLong(rows(j)("業務項目ID"))), Array(ToLong(rows(j)("実績ID")), ToLong(rows(j)("件数")))
        Next
    End If
    Set LoadTaskRows = dict
End Function

Function IsLocked(day)
    Dim rows
    rows = DbQuery(SqlDaily(), Array(day, DateAdd("d", 1, day)))
    IsLocked = False
    If UBound(rows) >= 0 Then
        If ToStr(rows(0)("状態")) = "確定" Then IsLocked = True
    End If
End Function

Function PersonInPeriod(row, day)
    PersonInPeriod = row("有効")
    If Not IsNull(row("在籍開始日")) Then
        If row("在籍開始日") > day Then PersonInPeriod = False
    End If
    If Not IsNull(row("在籍終了日")) Then
        If row("在籍終了日") < day Then PersonInPeriod = False
    End If
End Function

Function PersonName(t)
    Dim j
    PersonName = ""
    For j = 0 To UBound(people)
        If ToLong(people(j)("担当者ID")) = t Then PersonName = ToStr(people(j)("氏名"))
    Next
End Function

Sub ShowPersonBar()
    Dim j, label
    Response.Write "<form method=""post"" action=""tasks.asp"" class=""bar noprint"">"
    Response.Write "<input type=""hidden"" name=""act"" value=""go"">"
    Response.Write "<a class=""btn btn-s"" href=""tasks.asp?d=" & YMD(DateAdd("d", -1, d)) & "&amp;t=" & tid & """>&lt; 前の日</a> "
    Response.Write "<label>日付 <input type=""date"" name=""d"" value=""" & YMD(d) & """ class=""in-date""></label> "
    Response.Write "<a class=""btn btn-s"" href=""tasks.asp?d=" & YMD(DateAdd("d", 1, d)) & "&amp;t=" & tid & """>次の日 &gt;</a> "
    Response.Write "<a class=""btn btn-s"" href=""tasks.asp?d=" & YMD(Date()) & "&amp;t=" & tid & """>今日</a> "
    Response.Write "<label>担当者 <select name=""t""><option value=""0"">（選んでください）</option>"
    For j = 0 To UBound(people)
        If PersonInPeriod(people(j), d) Or ToLong(people(j)("担当者ID")) = tid Then
            label = ToStr(people(j)("担当者コード")) & " " & ToStr(people(j)("氏名"))
            Response.Write "<option value=""" & ToLong(people(j)("担当者ID")) & """" & SelAttr(ToLong(people(j)("担当者ID")), tid) & ">" & H(label) & "</option>"
        End If
    Next
    Response.Write "</select></label> "
    Response.Write "<button type=""submit"" class=""btn btn-s"">表示</button>"
    Response.Write "<span class=""bar-date"">" & DateJa(d) & "</span>"
    Response.Write "</form>"
End Sub

Sub ShowList(locked)
    Dim j, key, v, cls, dis, total
    dis = ""
    If locked Then dis = " disabled"
    total = 0
    Response.Write "<table class=""tw""><tr><th class=""c"">番号</th><th>業務</th><th class=""c"">件数</th></tr>"
    For j = 0 To UBound(items)
        key = CStr(ToLong(items(j)("業務項目ID")))
        If items(j)("有効") Or existing.Exists(key) Then
            If posted.Exists(key) Then
                v = posted(key)
            ElseIf existing.Exists(key) Then
                v = CStr(existing(key)(1))
                total = total + existing(key)(1)
            Else
                v = ""
            End If
            cls = "num"
            If bad.Exists(key) Then cls = "num ng"
            If items(j)("有効") Then
                Response.Write "<tr>"
            Else
                Response.Write "<tr class=""off"">"
            End If
            Response.Write "<td class=""c"">" & H(items(j)("番号")) & "</td><td>" & H(items(j)("項目名"))
            If Not items(j)("有効") Then Response.Write " <span class=""note"">（今は使っていない項目）</span>"
            Response.Write "</td><td class=""c""><input type=""text"" class=""" & cls & """ name=""g_" & key & """ value=""" & H(v) & """ inputmode=""numeric"" maxlength=""4"" size=""4""" & dis & "></td></tr>"
        End If
    Next
    Response.Write "<tr class=""sum""><td></td><td>計（保存済み）</td><td class=""n"">" & total & "</td></tr></table>"
End Sub

PageHead "その他業務"
If saved Then ShowMsg "ok", "保存しました。"
If msgNg <> "" Then ShowMsg "ng", msgNg
ShowPersonBar
If UBound(people) < 0 Then
    ShowMsg "note", "担当者がまだ登録されていません。職員に「マスタ保守」で担当者を登録してもらってください。"
ElseIf tid = 0 Then
    ShowMsg "note", "上で「担当者」を選んで「表示」を押してください。"
ElseIf IsLocked(d) Then
    Response.Write "<p>" & H(PersonName(tid)) & " さんの " & DateShort(d) & "</p>"
    ShowMsg "note", "この日の日報は確定済みです。直すときは、職員に「確定の取り消し」を頼んでください。"
    ShowList True
Else
%>
<p><%= H(PersonName(tid)) %> さんの <%= DateShort(d) %> の、電話以外の仕事の件数を入れて「保存する」を押してください。</p>
<form method="post" action="tasks.asp">
<input type="hidden" name="act" value="save">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<input type="hidden" name="t" value="<%= tid %>">
<% ShowList False %>
<div class="actions"><button type="submit" class="btn">保存する</button></div>
</form>
<%
End If
PageFoot
%>
