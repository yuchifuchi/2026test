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
'  週の集計表（月曜～日曜）。職員用。
'  数はすべて sql.asp の SqlCallsBase（T_受電 ただ 1 つ）から数える。日報の紙と同じ出どころ。
'  週の合計は、日ごとの数を足したものと、週全体を 1 回で数えたものの両方を出し、
'  食い違えば赤で知らせる（構造上は起きないはずの事故を、起きたらすぐ分かるように）。
' ============================================================
Dim d, dd, ws, we, days(6), i, j, msgNg, colsAll, blocks, kubunAll, people
Dim dayCols, dayBrk, dayKubun, dayPerson, attCount, states, weekCols, weekBrk

RequireStaff
d = PageDate()
msgNg = ""
If IsPost() Then
    If ParseYMD(FormVal("d"), dd) Then Go "summary.asp?d=" & YMD(WeekStart(dd))
    msgNg = "日付が読めません。2026-08-25 のように入れてください。"
End If
ws = WeekStart(d)
we = DateAdd("d", 7, ws)
For i = 0 To 6
    days(i) = DateAdd("d", i, ws)
Next

colsAll = DbQuery(SqlColumnsAll(), Array())
blocks = DbQuery(SqlBlocksAll(), Array())
kubunAll = DbQuery(SqlKubunAll(), Array())
people = DbQuery(SqlStaffAll(), Array())
Set dayCols = Server.CreateObject("Scripting.Dictionary")
Set dayBrk = Server.CreateObject("Scripting.Dictionary")
Set dayKubun = Server.CreateObject("Scripting.Dictionary")
Set dayPerson = Server.CreateObject("Scripting.Dictionary")
Set attCount = Server.CreateObject("Scripting.Dictionary")
Set states = Server.CreateObject("Scripting.Dictionary")
LoadAll

Sub LoadAll()
    Dim rows, k, key, v
    ' 同じ日・同じ列の行が 2 つ来ても数を失わないよう、上書きせずに足し込む
    rows = DbQuery(SqlDayColumnTotals(), Array(ws, we))
    For k = 0 To UBound(rows)
        key = YMD(rows(k)("日")) & "|" & ToLong(rows(k)("集計列ID"))
        If dayCols.Exists(key) Then
            v = dayCols(key)
            dayCols(key) = Array(v(0) + ToLong(rows(k)("件数計")), v(1) + ToLong(rows(k)("職員計")))
        Else
            dayCols(key) = Array(ToLong(rows(k)("件数計")), ToLong(rows(k)("職員計")))
        End If
    Next
    rows = DbQuery(SqlDayBreakdownTotals(), Array(ws, we))
    For k = 0 To UBound(rows)
        AddTo dayBrk, YMD(rows(k)("日")) & "|" & ToStr(rows(k)("内訳区分")), ToLong(rows(k)("件数計"))
    Next
    rows = DbQuery(SqlDayKubunTotals(), Array(ws, we))
    For k = 0 To UBound(rows)
        AddTo dayKubun, YMD(rows(k)("日")) & "|" & ToLong(rows(k)("区分ID")), ToLong(rows(k)("件数計"))
    Next
    rows = DbQuery(SqlDayPersonTotals(), Array(ws, we))
    For k = 0 To UBound(rows)
        AddTo dayPerson, YMD(rows(k)("日")) & "|" & ToLong(rows(k)("担当者ID")), ToLong(rows(k)("件数計"))
    Next
    rows = DbQuery(SqlAttendance(), Array(ws, we))
    For k = 0 To UBound(rows)
        AddTo attCount, YMD(rows(k)("対象日")), 1
    Next
    rows = DbQuery(SqlDaily(), Array(ws, we))
    For k = 0 To UBound(rows)
        states(YMD(rows(k)("対象日"))) = ToStr(rows(k)("状態"))
    Next
    weekCols = DbQuery(SqlColumnTotals(), Array(ws, we))
    weekBrk = DbQuery(SqlBreakdownTotals(), Array(ws, we))
End Sub

Sub AddTo(dict, key, n)
    If dict.Exists(key) Then
        dict(key) = dict(key) + n
    Else
        dict(key) = n
    End If
End Sub

Function Got(dict, key)
    If dict.Exists(key) Then Got = dict(key) Else Got = 0
End Function

Function CellCount(n, st)
    If n = 0 Then
        CellCount = "<td class=""n zero"">0</td>"
    Else
        CellCount = "<td class=""n"">" & n & " <span class=""note"">(" & st & ")</span></td>"
    End If
End Function

Function RowClass(day)
    RowClass = ""
    If Weekday(day) = vbSaturday Then RowClass = " class=""sat"""
    If Weekday(day) = vbSunday Then RowClass = " class=""sun"""
End Function

Function KubunRowHasData(kid)
    Dim k
    KubunRowHasData = False
    For k = 0 To 6
        If Got(dayKubun, YMD(days(k)) & "|" & kid) > 0 Then KubunRowHasData = True
    Next
End Function

Function PersonRowHasData(pid)
    Dim k
    PersonRowHasData = False
    For k = 0 To 6
        If Got(dayPerson, YMD(days(k)) & "|" & pid) > 0 Then PersonRowHasData = True
    Next
End Function

Sub ShowDayTable()
    Dim k, c, key, v, tot, totSt, sumTot, sumSt, ex, rf, wt, wst, wex, wrf, st
    Response.Write "<h2>日ごとの件数</h2><p class=""note"">（ ）内は職員が受けた件数</p>"
    Response.Write "<table class=""tw wide""><tr><th>日付</th><th class=""c"">出勤者</th><th class=""c"">合計</th>"
    For c = 0 To UBound(colsAll)
        Response.Write "<th class=""c"">" & H(colsAll(c)("集計列名")) & "</th>"
    Next
    Response.Write "<th class=""c"">内 交換</th><th class=""c"">内 返金</th><th class=""c"">日報</th></tr>"
    sumTot = 0
    sumSt = 0
    For k = 0 To 6
        key = YMD(days(k))
        tot = 0
        totSt = 0
        For c = 0 To UBound(colsAll)
            If dayCols.Exists(key & "|" & ToLong(colsAll(c)("集計列ID"))) Then
                v = dayCols(key & "|" & ToLong(colsAll(c)("集計列ID")))
                tot = tot + v(0)
                totSt = totSt + v(1)
            End If
        Next
        sumTot = sumTot + tot
        sumSt = sumSt + totSt
        Response.Write "<tr" & RowClass(days(k)) & "><td><a href=""daily.asp?d=" & key & """>" & DateShort(days(k)) & "</a></td>"
        Response.Write "<td class=""n"">" & Got(attCount, key) & " 人</td>" & CellCount(tot, totSt)
        For c = 0 To UBound(colsAll)
            If dayCols.Exists(key & "|" & ToLong(colsAll(c)("集計列ID"))) Then
                v = dayCols(key & "|" & ToLong(colsAll(c)("集計列ID")))
                Response.Write CellCount(v(0), v(1))
            Else
                Response.Write CellCount(0, 0)
            End If
        Next
        ex = Got(dayBrk, key & "|交換")
        rf = Got(dayBrk, key & "|返金")
        st = "－"
        If states.Exists(key) Then st = states(key)
        Response.Write "<td class=""n"">" & ex & "</td><td class=""n"">" & rf & "</td><td class=""c"">" & H(st) & "</td></tr>"
    Next
    ' 週の合計（週全体を 1 回で数えたもの）
    wt = 0
    wst = 0
    For c = 0 To UBound(weekCols)
        wt = wt + ToLong(weekCols(c)("件数計"))
        wst = wst + ToLong(weekCols(c)("職員計"))
    Next
    wex = 0
    wrf = 0
    For c = 0 To UBound(weekBrk)
        If ToStr(weekBrk(c)("内訳区分")) = "交換" Then wex = wex + ToLong(weekBrk(c)("件数計"))
        If ToStr(weekBrk(c)("内訳区分")) = "返金" Then wrf = wrf + ToLong(weekBrk(c)("件数計"))
    Next
    Response.Write "<tr class=""sum""><td>週の合計</td><td></td>" & CellCount(wt, wst)
    For c = 0 To UBound(weekCols)
        Response.Write CellCount(ToLong(weekCols(c)("件数計")), ToLong(weekCols(c)("職員計")))
    Next
    Response.Write "<td class=""n"">" & wex & "</td><td class=""n"">" & wrf & "</td><td></td></tr></table>"
    If wt <> sumTot Or wst <> sumSt Then
        ShowMsg "ng", "日ごとの数を足したもの（" & sumTot & "）と、週全体で数えたもの（" & wt & "）が合いません。システムの担当に知らせてください。"
    End If
End Sub

Sub ShowKubunTable()
    Dim b, k, j2, kid, rowTot, dayTot(6), n
    Response.Write "<h2>区分ごとの件数</h2><table class=""tw wide""><tr><th>区分（日報の列）</th>"
    For k = 0 To 6
        Response.Write "<th class=""c"">" & DateShort(days(k)) & "</th>"
        dayTot(k) = 0
    Next
    Response.Write "<th class=""c"">計</th></tr>"
    For b = 0 To UBound(blocks)
        Response.Write "<tr><th colspan=""9"">" & H(blocks(b)("ブロック名")) & "</th></tr>"
        For j2 = 0 To UBound(kubunAll)
            If ToLong(kubunAll(j2)("ブロックID")) = ToLong(blocks(b)("ブロックID")) Then
                kid = ToLong(kubunAll(j2)("区分ID"))
                If kubunAll(j2)("有効") Or KubunRowHasData(kid) Then
                    rowTot = 0
                    Response.Write "<tr><td>" & H(kubunAll(j2)("区分名")) & "</td>"
                    For k = 0 To 6
                        n = Got(dayKubun, YMD(days(k)) & "|" & kid)
                        rowTot = rowTot + n
                        dayTot(k) = dayTot(k) + n
                        If n = 0 Then
                            Response.Write "<td class=""n zero"">0</td>"
                        Else
                            Response.Write "<td class=""n"">" & n & "</td>"
                        End If
                    Next
                    Response.Write "<td class=""n""><b>" & rowTot & "</b></td></tr>"
                End If
            End If
        Next
    Next
    rowTot = 0
    Response.Write "<tr class=""sum""><td>計</td>"
    For k = 0 To 6
        rowTot = rowTot + dayTot(k)
        Response.Write "<td class=""n"">" & dayTot(k) & "</td>"
    Next
    Response.Write "<td class=""n"">" & rowTot & "</td></tr></table>"
End Sub

Sub ShowPersonTable()
    Dim j2, k, pid, rowTot, n, any
    Response.Write "<h2>担当者ごとの件数</h2><table class=""tw wide""><tr><th>担当者</th>"
    For k = 0 To 6
        Response.Write "<th class=""c"">" & DateShort(days(k)) & "</th>"
    Next
    Response.Write "<th class=""c"">計</th></tr>"
    any = False
    For j2 = 0 To UBound(people)
        pid = ToLong(people(j2)("担当者ID"))
        If PersonRowHasData(pid) Then
            any = True
            rowTot = 0
            Response.Write "<tr><td>" & H(people(j2)("担当者コード")) & " " & H(people(j2)("氏名")) & "</td>"
            For k = 0 To 6
                n = Got(dayPerson, YMD(days(k)) & "|" & pid)
                rowTot = rowTot + n
                If n = 0 Then
                    Response.Write "<td class=""n zero"">0</td>"
                Else
                    Response.Write "<td class=""n"">" & n & "</td>"
                End If
            Next
            Response.Write "<td class=""n""><b>" & rowTot & "</b></td></tr>"
        End If
    Next
    If Not any Then Response.Write "<tr><td colspan=""9"">この週の受付入力はまだありません。</td></tr>"
    Response.Write "</table>"
End Sub

PageHead "週の集計表"
If msgNg <> "" Then ShowMsg "ng", msgNg
%>
<form method="post" action="summary.asp" class="bar noprint">
  <a class="btn btn-s" href="summary.asp?d=<%= YMD(DateAdd("d", -7, ws)) %>">&lt; 前の週</a>
  <label>この日を含む週 <input type="date" name="d" value="<%= YMD(ws) %>" class="in-date"></label>
  <button type="submit" class="btn btn-s">表示</button>
  <a class="btn btn-s" href="summary.asp?d=<%= YMD(DateAdd("d", 7, ws)) %>">次の週 &gt;</a>
  <a class="btn btn-s" href="summary.asp?d=<%= YMD(Date()) %>">今週</a>
  <span class="bar-date"><%= DateJa(ws) %> ～ <%= DateJa(days(6)) %></span>
</form>
<%
ShowDayTable
ShowKubunTable
ShowPersonTable
PageFoot
%>
