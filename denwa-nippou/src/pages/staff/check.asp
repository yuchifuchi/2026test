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
'  入力もれチェック（職員用）
'  ・出勤しているのに入力が無い人、入力があるのに出勤に入っていない人を赤で出す。
'    （現行 Excel では、名前のずれたファイルの人が「該当者なし」で黙って飛ばされた。
'      ここでは誰が抜けているかを、人ごとに必ず 1 行出す）
'  ・T_受電 を直接数えた数と、日報の紙に使う SELECT を通した数を突き合わせる。
' ============================================================
Dim d, dd, d2, msgNg, people, att, calls, tasks, daily, i, key, attSet, problems, raw, base, ws, k

RequireStaff
d = PageDate()
msgNg = ""
If IsPost() Then
    If ParseYMD(FormVal("d"), dd) Then Go "check.asp?d=" & YMD(dd)
    msgNg = "日付が読めません。2026-08-25 のように入れてください。"
End If
d2 = DateAdd("d", 1, d)
people = DbQuery(SqlStaffAll(), Array())
att = DbQuery(SqlAttendance(), Array(d, d2))
Set calls = ToDict(DbQuery(SqlPersonTotals(), Array(d, d2)))
Set tasks = ToDict(DbQuery(SqlTaskPersonTotals(), Array(d, d2)))
daily = DbQuery(SqlDaily(), Array(d, d2))
raw = DbQuery(SqlRawCallsCheck(), Array(d, d2))
base = DbQuery(SqlBaseCallsCheck(), Array(d, d2))
Set attSet = Server.CreateObject("Scripting.Dictionary")
For i = 0 To UBound(att)
    attSet(CStr(ToLong(att(i)("担当者ID")))) = True
Next
problems = 0

Function ToDict(rows)
    Dim dict, j
    Set dict = Server.CreateObject("Scripting.Dictionary")
    For j = 0 To UBound(rows)
        dict(CStr(ToLong(rows(j)("担当者ID")))) = ToLong(rows(j)("件数計"))
    Next
    Set ToDict = dict
End Function

Function Got(dict, key2)
    If dict.Exists(key2) Then Got = dict(key2) Else Got = 0
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

Sub ShowPeople()
    Dim j, key2, inP, onDuty, c, t, verdict, cls
    Response.Write "<table class=""tw""><tr><th>担当者</th><th class=""c"">出勤</th><th class=""c"">受付入力</th><th class=""c"">その他業務</th><th>判定</th></tr>"
    For j = 0 To UBound(people)
        key2 = CStr(ToLong(people(j)("担当者ID")))
        inP = PersonInPeriod(people(j), d)
        onDuty = attSet.Exists(key2)
        c = Got(calls, key2)
        t = Got(tasks, key2)
        If inP Or onDuty Or c > 0 Or t > 0 Then
            cls = ""
            If onDuty And c = 0 And t = 0 Then
                verdict = "<span class=""ng"">出勤しているのに、入力がありません</span>"
                problems = problems + 1
            ElseIf Not onDuty And (c > 0 Or t > 0) Then
                verdict = "<span class=""ng"">入力があるのに、出勤者に入っていません</span>"
                problems = problems + 1
            ElseIf Not inP And (c > 0 Or t > 0) Then
                verdict = "<span class=""warn"">在籍期間の外の日に入力があります</span>"
                problems = problems + 1
            ElseIf onDuty Then
                verdict = "<span class=""ok"">OK</span>"
            Else
                verdict = "－（出勤なし・入力なし）"
                cls = " class=""off"""
            End If
            Response.Write "<tr" & cls & "><td>" & H(people(j)("担当者コード")) & " " & H(people(j)("氏名")) & "</td>"
            If onDuty Then
                Response.Write "<td class=""c"">○</td>"
            Else
                Response.Write "<td class=""c"">－</td>"
            End If
            Response.Write "<td class=""n"">" & c & " 件</td><td class=""n"">" & t & " 件</td><td>" & verdict & "</td></tr>"
        End If
    Next
    Response.Write "</table>"
End Sub

Sub ShowDaily()
    If UBound(daily) < 0 Then
        Response.Write "<p><span class=""warn"">日報はまだ作っていません。</span> <a href=""daily.asp?d=" & YMD(d) & """>日報の作成・確定へ</a></p>"
        problems = problems + 1
    Else
        Response.Write "<p>日報: <b>" & H(daily(0)("状態")) & "</b>"
        If IsNull(daily(0)("回線数")) Then
            Response.Write "　<span class=""warn"">回線数が入っていません</span>"
            problems = problems + 1
        Else
            Response.Write "　回線数 " & ToLong(daily(0)("回線数"))
        End If
        If UBound(att) < 0 Then
            Response.Write "　<span class=""warn"">出勤者が入っていません</span>"
            problems = problems + 1
        End If
        Response.Write "</p>"
    End If
End Sub

Sub ShowReconcile()
    Dim r1, r2, n1, n2
    r1 = ToLong(raw(0)("行数"))
    n1 = ToLong(raw(0)("件数計"))
    r2 = ToLong(base(0)("行数"))
    n2 = ToLong(base(0)("件数計"))
    If r1 = r2 And n1 = n2 Then
        Response.Write "<p><span class=""ok"">一致</span>　受付入力の明細 " & r1 & " 行・" & n1 & " 件が、すべて日報の合計に入っています。</p>"
    Else
        Response.Write "<p><span class=""ng"">不一致</span>　明細は " & r1 & " 行・" & n1 & " 件ですが、日報の合計に入ったのは " & r2 & " 行・" & n2 & " 件です。システムの担当に知らせてください。</p>"
        problems = problems + 1
    End If
End Sub

Sub ShowWeek()
    Dim k2, day, rowsA, rowsC, rowsD, nA, nC, st
    ws = WeekStart(d)
    Response.Write "<table class=""tw""><tr><th>日付</th><th class=""c"">出勤者</th><th class=""c"">受付入力をした人</th><th class=""c"">日報</th><th></th></tr>"
    For k2 = 0 To 6
        day = DateAdd("d", k2, ws)
        rowsA = DbQuery(SqlAttendance(), Array(day, DateAdd("d", 1, day)))
        rowsC = DbQuery(SqlPersonTotals(), Array(day, DateAdd("d", 1, day)))
        rowsD = DbQuery(SqlDaily(), Array(day, DateAdd("d", 1, day)))
        nA = UBound(rowsA) + 1
        nC = UBound(rowsC) + 1
        st = "－"
        If UBound(rowsD) >= 0 Then st = ToStr(rowsD(0)("状態"))
        Response.Write "<tr><td>" & DateShort(day) & "</td><td class=""n"">" & nA & " 人</td><td class=""n"">" & nC & " 人</td><td class=""c"">" & H(st) & "</td>"
        Response.Write "<td><a href=""check.asp?d=" & YMD(day) & """>この日を調べる</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

PageHead "入力もれチェック"
If msgNg <> "" Then ShowMsg "ng", msgNg
DateBar "check.asp", d, "", ""
Response.Write "<h2>人ごとの入力</h2>"
ShowPeople
Response.Write "<h2>日報</h2>"
ShowDaily
Response.Write "<h2>数の突き合わせ</h2>"
ShowReconcile
If problems = 0 Then
    ShowMsg "ok", "この日は、気になる所はありません。"
Else
    ShowMsg "note", "上の赤・黄の所を確かめてください（" & problems & " か所）。入力を直すときは、職員メニューの「受付入力」「その他業務」から代わりに入力できます。"
End If
Response.Write "<h2>この週の様子</h2>"
ShowWeek
PageFoot
%>
