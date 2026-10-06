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
'  日報の作成・確定（職員用）
'  出勤者・回線数・記述欄 3 つを入れて保存し、よければ「確定」する。
'  確定した日は、受付入力・その他業務の数も直せなくなる（紙に出した数と食い違わないように）。
'  件数そのものは、ここでは入力しない（受付入力の数を T_受電 から数えるだけ）。
' ============================================================

' 記述欄が帳票の A4 1 枚に収まるための決まり（tools/render_sheet.py で、上限いっぱいでも 1 枚に収まることを確かめている）
Const TEXT_PER_LINE = 46      ' 記述欄 1 行に入る全角の文字数
Const TEXT_MIN_LINES = 3      ' 1 つの欄が最低でも取る行数（空でもこの高さの枠がある）
Const TEXT_TOTAL_LINES = 15   ' 3 つの欄の合計の行数の上限
Const ATT_MAX = 14            ' 出勤者の上限（帳票の氏名欄 2 列 × 7 行）
Const ATT_NOTE_MAX = 12       ' 勤務時間と備考を合わせた文字数の上限（帳票の備考欄に収まるように）

Dim d, d2, msgNg, msgOk, people, daily, att, state, hasDaily, lines, t1, t2, t3, attSet, hrs, notes, bad
Dim personCalls, personTasks, i, dd, pid, prefill

d = PageDate()
msgNg = ""
msgOk = ""
Set attSet = Server.CreateObject("Scripting.Dictionary")
Set hrs = Server.CreateObject("Scripting.Dictionary")
Set notes = Server.CreateObject("Scripting.Dictionary")
Set bad = Server.CreateObject("Scripting.Dictionary")

RequireStaff

If QueryVal("saved") = "1" Then msgOk = "保存しました。"
If QueryVal("saved") = "fixed" Then msgOk = "確定しました。帳票（PDF）を出せます。"
If QueryVal("saved") = "unfixed" Then msgOk = "確定を取り消しました。受付入力の数を直せるようになりました。"

people = DbQuery(SqlStaffAll(), Array())

If IsPost() Then
    If FormVal("act") = "go" Then
        If ParseYMD(FormVal("d"), dd) Then Go "daily.asp?d=" & YMD(dd)
        msgNg = "日付が読めません。2026-08-25 のように入れてください。"
    ElseIf FormVal("act") = "save" Or FormVal("act") = "fix" Then
        SaveDaily (FormVal("act") = "fix")
    ElseIf FormVal("act") = "unfix" Then
        Unfix
    End If
End If

d2 = DateAdd("d", 1, d)
LoadDay

Sub LoadDay()
    Dim j, key
    daily = DbQuery(SqlDaily(), Array(d, d2))
    hasDaily = (UBound(daily) >= 0)
    state = "作成中"
    If hasDaily Then state = ToStr(daily(0)("状態"))
    att = DbQuery(SqlAttendance(), Array(d, d2))
    Set personCalls = ToDict(DbQuery(SqlPersonTotals(), Array(d, d2)))
    Set personTasks = ToDict(DbQuery(SqlTaskPersonTotals(), Array(d, d2)))
    If msgNg = "" Then
        ' 保存済みの値を画面に出す（誤りで戻ってきたときは、打った値をそのまま残す）
        If hasDaily Then
            If IsNull(daily(0)("回線数")) Then lines = "" Else lines = CStr(ToLong(daily(0)("回線数")))
            t1 = ToStr(daily(0)("特記事項"))
            t2 = ToStr(daily(0)("職員代替案件"))
            t3 = ToStr(daily(0)("要望"))
        Else
            lines = CStr(ToLong(DbScalar(SqlLastLines(), Array(d), 5)))
            t1 = ""
            t2 = ""
            t3 = ""
        End If
        prefill = False
        If UBound(att) < 0 And Not hasDaily Then
            ' まだ出勤者を入れていない日は、入力のあった人にあらかじめ印を付ける
            For j = 0 To UBound(people)
                key = CStr(ToLong(people(j)("担当者ID")))
                If personCalls.Exists(key) Or personTasks.Exists(key) Then
                    attSet(key) = True
                    prefill = True
                End If
            Next
        End If
        For j = 0 To UBound(att)
            key = CStr(ToLong(att(j)("担当者ID")))
            attSet(key) = True
            hrs(key) = ToStr(att(j)("勤務時間"))
            notes(key) = ToStr(att(j)("備考"))
        Next
    End If
End Sub

' 担当者ID → 件数 の Dictionary にする
Function ToDict(rows)
    Dim dict, j
    Set dict = Server.CreateObject("Scripting.Dictionary")
    For j = 0 To UBound(rows)
        dict(CStr(ToLong(rows(j)("担当者ID")))) = ToLong(rows(j)("件数計"))
    Next
    Set ToDict = dict
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

Function TextBudget(a, b, c)
    TextBudget = MaxOf(TextLines(a, TEXT_PER_LINE), TEXT_MIN_LINES) + MaxOf(TextLines(b, TEXT_PER_LINE), TEXT_MIN_LINES) + MaxOf(TextLines(c, TEXT_PER_LINE), TEXT_MIN_LINES)
End Function

Function MaxOf(a, b)
    If a > b Then MaxOf = a Else MaxOf = b
End Function

Sub SaveDaily(fix)
    Dim j, key, n, cnt, stamp, cur, newState, fixedAt, lineVal
    If Not ParseYMD(FormVal("d"), dd) Then
        msgNg = "日付が読めません。"
        Exit Sub
    End If
    d = dd
    d2 = DateAdd("d", 1, d)
    cur = DbQuery(SqlDaily(), Array(d, d2))
    If UBound(cur) >= 0 Then
        If ToStr(cur(0)("状態")) = "確定" Then
            msgNg = "この日は確定済みです。直すときは、先に「確定を取り消す」を押してください。"
            Exit Sub
        End If
    End If
    ' 送られてきた値を取っておく（誤りがあったら、このまま画面に戻す）
    lines = FormVal("lines")
    t1 = FormText("t1")
    t2 = FormText("t2")
    t3 = FormText("t3")
    cnt = 0
    For j = 0 To UBound(people)
        key = CStr(ToLong(people(j)("担当者ID")))
        If FormVal("att_" & key) = "1" Then
            attSet(key) = True
            cnt = cnt + 1
        End If
        hrs(key) = FormVal("h_" & key)
        notes(key) = FormVal("n_" & key)
        If Len(hrs(key)) + Len(notes(key)) > ATT_NOTE_MAX Then bad("an_" & key) = True
    Next
    If cnt > ATT_MAX Then msgNg = "出勤者は " & ATT_MAX & " 人までです（帳票の氏名欄の数）。いま " & cnt & " 人に印が付いています。"
    If bad.Count > 0 Then msgNg = "赤い欄の「勤務時間」と「備考」は、合わせて " & ATT_NOTE_MAX & " 文字までにしてください（帳票の備考欄に収まるように）。"
    lineVal = Null
    If lines <> "" Then
        If ParseCount(lines, n) Then
            If n > 99 Then
                bad("lines") = True
            Else
                lineVal = n
            End If
        Else
            bad("lines") = True
        End If
        If bad.Exists("lines") Then msgNg = "回線数は 0～99 の数字で入れてください。"
    ElseIf fix Then
        bad("lines") = True
        msgNg = "確定するときは、回線数を入れてください。"
    End If
    If TextBudget(t1, t2, t3) > TEXT_TOTAL_LINES Then
        bad("text") = True
        msgNg = "記述欄が長すぎて、帳票が A4 1 枚に収まりません（今 " & TextBudget(t1, t2, t3) & " 行ぶん、" & TEXT_TOTAL_LINES & " 行まで）。短くしてください。"
    End If
    If msgNg <> "" Then Exit Sub
    stamp = Now()
    If fix Then
        newState = "確定"
        fixedAt = stamp
    Else
        newState = "作成中"
        fixedAt = Null
    End If
    DbBegin
    If UBound(cur) >= 0 Then
        DbExec SqlDailyUpdate(), Array(lineVal, t1, t2, t3, newState, fixedAt, stamp, d, d2)
    Else
        DbExec SqlDailyInsert(), Array(d, lineVal, t1, t2, t3, newState, fixedAt, stamp)
    End If
    DbExec SqlAttendanceDeleteDay(), Array(d, d2)
    For j = 0 To UBound(people)
        key = CStr(ToLong(people(j)("担当者ID")))
        If attSet.Exists(key) Then
            DbExec SqlAttendanceInsert(), Array(d, ToLong(people(j)("担当者ID")), hrs(key), notes(key))
        End If
    Next
    DbCommit
    If fix Then
        Go "daily.asp?d=" & YMD(d) & "&saved=fixed"
    Else
        Go "daily.asp?d=" & YMD(d) & "&saved=1"
    End If
End Sub

Sub Unfix()
    If Not ParseYMD(FormVal("d"), dd) Then
        msgNg = "日付が読めません。"
        Exit Sub
    End If
    d = dd
    DbExec SqlDailySetState(), Array("作成中", Null, Now(), d, DateAdd("d", 1, d))
    Go "daily.asp?d=" & YMD(d) & "&saved=unfixed"
End Sub

Function NgCls(key)
    If bad.Exists(key) Then NgCls = " ng" Else NgCls = ""
End Function

Sub ShowAttendance(locked)
    Dim j, key, dis, row, c, tk, shown
    dis = ""
    If locked Then dis = " disabled"
    Response.Write "<table class=""tw""><tr><th class=""c"">出勤</th><th>担当者</th><th class=""c"">受付入力</th><th class=""c"">その他業務</th><th>勤務時間（例 9:00-13:00）</th><th>備考</th></tr>"
    shown = 0
    For j = 0 To UBound(people)
        Set row = people(j)
        key = CStr(ToLong(row("担当者ID")))
        If PersonInPeriod(row, d) Or attSet.Exists(key) Or personCalls.Exists(key) Or personTasks.Exists(key) Then
            shown = shown + 1
            c = 0
            tk = 0
            If personCalls.Exists(key) Then c = personCalls(key)
            If personTasks.Exists(key) Then tk = personTasks(key)
            Response.Write "<tr><td class=""c""><input type=""checkbox"" name=""att_" & key & """ value=""1""" & ChkAttr(attSet.Exists(key)) & dis & "></td>"
            Response.Write "<td>" & H(row("担当者コード")) & " " & H(row("氏名"))
            If ToStr(row("職員区分")) = "職員" Then Response.Write " <span class=""note"">（職員）</span>"
            Response.Write "</td><td class=""n"">" & c & " 件</td><td class=""n"">" & tk & " 件</td>"
            Response.Write "<td><input type=""text"" name=""h_" & key & """ value=""" & H(hrs(key)) & """ size=""12"" maxlength=""14"" class=""" & Mid(NgCls("an_" & key), 2) & """" & dis & "></td>"
            Response.Write "<td><input type=""text"" name=""n_" & key & """ value=""" & H(notes(key)) & """ size=""16"" maxlength=""14"" class=""" & Mid(NgCls("an_" & key), 2) & """" & dis & "></td></tr>"
        End If
    Next
    Response.Write "</table>"
    If shown = 0 Then ShowMsg "note", "この日に在籍している担当者がいません。「マスタ保守」で担当者を登録してください。"
End Sub

Sub ShowTotals()
    Dim cols, brk, j, tot, s, ex, rf
    cols = DbQuery(SqlColumnTotals(), Array(d, d2))
    brk = DbQuery(SqlBreakdownTotals(), Array(d, d2))
    tot = 0
    s = ""
    For j = 0 To UBound(cols)
        tot = tot + ToLong(cols(j)("件数計"))
        If s <> "" Then s = s & " ／ "
        s = s & H(cols(j)("集計列名")) & " " & ToLong(cols(j)("件数計")) & "（職員 " & ToLong(cols(j)("職員計")) & "）"
    Next
    ex = 0
    rf = 0
    For j = 0 To UBound(brk)
        If ToStr(brk(j)("内訳区分")) = "交換" Then ex = ex + ToLong(brk(j)("件数計"))
        If ToStr(brk(j)("内訳区分")) = "返金" Then rf = rf + ToLong(brk(j)("件数計"))
    Next
    Response.Write "<div class=""totals"">この日の問合せ件数（受付入力の合計）：合計 <b data-col=""total"">" & tot & "</b>　" & s & "　内 交換 " & ex & "・返金 " & rf & "</div>"
End Sub

PageHead "日報の作成・確定"
If msgOk <> "" Then ShowMsg "ok", msgOk
If msgNg <> "" Then ShowMsg "ng", msgNg
DateBar "daily.asp", d, "", ""
If state = "確定" Then
    Response.Write "<p><span class=""stamp stamp-done"">確定済み</span>"
    If hasDaily Then
        If Not IsNull(daily(0)("確定日時")) Then Response.Write "　" & YMDHM(daily(0)("確定日時")) & " に確定"
    End If
    Response.Write "</p>"
Else
    Response.Write "<p><span class=""stamp stamp-draft"">作成中</span>　確定するまでは、受付入力の数を直せます。</p>"
End If
ShowTotals
Response.Write "<p><a class=""btn btn-sub btn-s"" href=""report.asp?d=" & YMD(d) & """>帳票を画面で見る</a> <a class=""btn btn-sub btn-s"" href=""printpdf.asp?d=" & YMD(d) & """>帳票を PDF で出す</a> <a class=""btn btn-sub btn-s"" href=""check.asp?d=" & YMD(d) & """>入力もれチェック</a></p>"
%>
<form method="post" action="daily.asp">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<h2>出勤者</h2>
<% If prefill Then ShowMsg "note", "受付入力・その他業務の入力があった人に、あらかじめ印を付けています。確かめて「保存する」を押してください。" %>
<% ShowAttendance (state = "確定") %>
<h2>回線数</h2>
<p><input type="text" name="lines" value="<%= H(lines) %>" class="num<%= NgCls("lines") %>" size="3" maxlength="2" inputmode="numeric"<% If state = "確定" Then Response.Write " disabled" %>> 回線</p>
<h2>記述欄</h2>
<p class="note">3 つの欄を合わせて、帳票でおよそ <%= TEXT_TOTAL_LINES %> 行（1 行は全角 <%= TEXT_PER_LINE %> 文字）までです。今は <%= TextBudget(t1, t2, t3) %> 行ぶんです。</p>
<p>特記事項<br><textarea name="t1" rows="4" class="<%= Mid(NgCls("text"), 2) %>"<% If state = "確定" Then Response.Write " disabled" %>><%= H(t1) %></textarea></p>
<p>職員に代わった案件（概要）<br><textarea name="t2" rows="4" class="<%= Mid(NgCls("text"), 2) %>"<% If state = "確定" Then Response.Write " disabled" %>><%= H(t2) %></textarea></p>
<p>要望<br><textarea name="t3" rows="4" class="<%= Mid(NgCls("text"), 2) %>"<% If state = "確定" Then Response.Write " disabled" %>><%= H(t3) %></textarea></p>
<div class="actions">
<% If state = "確定" Then %>
<button type="submit" name="act" value="unfix" class="btn btn-warn">確定を取り消す</button>
<% Else %>
<button type="submit" name="act" value="save" class="btn btn-sub">保存する（まだ確定しない）</button>
<button type="submit" name="act" value="fix" class="btn btn-ok">保存して確定する</button>
<% End If %>
</div>
</form>
<% PageFoot %>
