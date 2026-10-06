<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/grid.asp"-->
<%
' ============================================================
'  受付入力：1 人・1 日ぶんの電話を、製品ごと・お問合せ内容ごとに数える。
'
'  表の形は現行 Excel の「記入用フォーム」と同じ（行が製品・列がお問合せ内容。include/grid.asp）。
'  入れ方は 2 通り:
'    ・1 件ずつ数える（初めに開く方）… 欄を押すたびに 1 件足す。Excel の「欄を選んでボタンを押す」と同じ使い方。
'      足し算は Access にさせる（[件数] = [件数] + 1）。続けて 2 回押されたときは、画面の「版」が古いので
'      2 回目は数えない（1 本の電話を 2 件に数えないように）。押し間違いは「1 件戻す」で戻す。
'    ・数をまとめて入れる・直す … 欄に今の件数を書いて保存する（足し算ではなく上書き。0 にした欄は行を消す）。
'
'  ・誰の数かは「担当者ID」で持つ（名前の文字では突き合わせない）。
'    現行 Excel ではファイルの中の名前がずれていて、その人の数が黙って飛ばされた。
'  ・画面の合計は、日報の紙と同じ SQL（sql.asp の SqlCallsBase）で数える。
' ============================================================
Const MEMO_MAX = 400
Dim d, tid, mode, msgNg, msgOk, people, existing, posted, bad, dd, cookieVal, lastKey, op, memoText

d = PageDate()
tid = ParamLong(QueryVal("t"), 0)
If tid = 0 Then
    cookieVal = Request.Cookies("nippou_t")
    tid = ParamLong(cookieVal, 0)
End If
mode = QueryVal("m")
If mode <> "edit" Then mode = "count"
msgNg = ""
msgOk = ""
lastKey = QueryVal("last")
op = QueryVal("op")
memoText = ""
Set posted = Server.CreateObject("Scripting.Dictionary")
Set bad = Server.CreateObject("Scripting.Dictionary")

GridLoad

If IsPost() Then
    If FormVal("act") = "go" Then
        If ParseYMD(FormVal("d"), dd) Then
            Go "entry.asp?d=" & YMD(dd) & "&t=" & ParamLong(FormVal("t"), 0) & "&m=" & mode
        End If
        msgNg = "日付が読めません。2026-08-25 のように入れてください。"
    ElseIf FormVal("act") = "save" Then
        mode = "edit"
        SaveEntry
    ElseIf FormVal("act") = "add" Then
        CountOne
    ElseIf FormVal("act") = "memo" Then
        SaveMemo
    End If
End If

Set existing = GridLoadCalls(d, tid)
people = DbQuery(SqlStaffAll(), Array())

' ---------- 受け取った日・人を確かめる（保存の前に） ----------
Function TakeDayPerson()
    TakeDayPerson = False
    If Not ParseYMD(FormVal("d"), dd) Then
        msgNg = "日付が読めません。"
        Exit Function
    End If
    d = dd
    tid = ParamLong(FormVal("t"), 0)
    If tid = 0 Then
        msgNg = "担当者を選んでから入れてください。"
        Exit Function
    End If
    If IsLocked(d) Then
        msgNg = "この日の日報は確定済みなので、入れられませんでした。直すときは、職員に「確定の取り消し」を頼んでください。"
        Exit Function
    End If
    TakeDayPerson = True
End Function

' ---------- 1 件ずつ数える ----------
Sub CountOne()
    Dim key, delta, ex, valid, stamp, back, parts
    If Not TakeDayPerson() Then Exit Sub
    ' 「1 件戻す」のボタンは name="undo"、欄のボタンは name="k"
    key = FormVal("undo")
    delta = -1
    If key = "" Then
        key = FormVal("k")
        delta = 1
    End If
    Set ex = GridLoadCalls(d, tid)
    Set valid = GridValidCells(d, ex)
    If Not valid.Exists(key) Then
        msgNg = "押された欄が見つかりません。画面を開き直してから、もう一度押してください。"
        Exit Sub
    End If
    back = "entry.asp?d=" & YMD(d) & "&t=" & tid & "&last=" & key
    If FormVal("ver") <> CallVersion(d, tid) Then
        ' 続けて 2 回押された（1 回目で数が変わったので、画面の版が古い）か、別の画面で数が変わった
        RememberPerson tid
        Go back & "&op=dup#b" & GridBlockOfKey(key)
    End If
    stamp = Now()
    DbBegin
    If ex.Exists(key) Then
        If delta < 0 And ex(key)(1) <= 1 Then
            DbExec SqlCallDelete(), Array(ex(key)(0))
        Else
            DbExec SqlCallAddDelta(), Array(delta, stamp, Registrant(), ex(key)(0))
        End If
    ElseIf delta > 0 Then
        parts = Split(key, "_")
        DbExec SqlCallInsert(), Array(d, tid, CLng(parts(0)), CLng(parts(1)), 1, stamp, stamp, Registrant())
    End If
    DbCommit
    RememberPerson tid
    If delta > 0 Then
        Go back & "&op=add#b" & GridBlockOfKey(key)
    Else
        Go back & "&op=sub#b" & GridBlockOfKey(key)
    End If
End Sub

' ---------- 数をまとめて入れる・直す ----------
Sub SaveEntry()
    Dim ex, valid, key, raw, n, parts, stamp, who, errCount, parsed
    If Not TakeDayPerson() Then Exit Sub
    If FormVal("ver") <> CallVersion(d, tid) Then
        msgNg = "この画面を開いている間に、別の画面でこの人のこの日の件数が変わりました。いま保存されている数を出し直しました。確かめて、もう一度入力・保存してください。"
        Exit Sub
    End If
    Set parsed = Server.CreateObject("Scripting.Dictionary")
    Set ex = GridLoadCalls(d, tid)
    Set valid = GridValidCells(d, ex)
    errCount = 0
    For Each key In valid.Keys
        raw = Request.Form("c_" & key)
        If Not IsEmpty(raw) Then
            posted(key) = CStr(raw)
            If ParseCount(CStr(raw), n) Then
                parsed(key) = n
            Else
                bad(key) = True
                errCount = errCount + 1
            End If
        End If
    Next
    If errCount > 0 Then
        msgNg = "赤い欄に、数字ではないもの（または 5 けた以上の数）が入っています。直してから、もう一度「保存する」を押してください。（まだ保存していません）"
        Exit Sub
    End If
    stamp = Now()
    who = Registrant()
    DbBegin
    For Each key In parsed.Keys
        n = parsed(key)
        If ex.Exists(key) Then
            If n = 0 Then
                DbExec SqlCallDelete(), Array(ex(key)(0))
            ElseIf n <> ex(key)(1) Then
                DbExec SqlCallUpdate(), Array(n, stamp, who, ex(key)(0))
            End If
        ElseIf n > 0 Then
            parts = Split(key, "_")
            DbExec SqlCallInsert(), Array(d, tid, CLng(parts(0)), CLng(parts(1)), n, stamp, stamp, who)
        End If
    Next
    DbCommit
    RememberPerson tid
    Go "entry.asp?d=" & YMD(d) & "&t=" & tid & "&m=edit&op=saved"
End Sub

' ---------- 特殊な問合せの内容 ----------
Sub SaveMemo()
    Dim text, rows
    If Not TakeDayPerson() Then Exit Sub
    text = TrimAll(Replace(FormVal("memo"), vbCr, ""))
    If Len(text) > MEMO_MAX Then
        memoText = FormVal("memo")
        msgNg = "特殊な問合せの内容は " & MEMO_MAX & " 文字までにしてください（いま " & Len(text) & " 文字。まだ保存していません）。"
        Exit Sub
    End If
    text = Replace(text, vbLf, vbCrLf)
    rows = DbQuery(SqlMemoOnePerson(), Array(d, DateAdd("d", 1, d), tid))
    DbBegin
    If UBound(rows) >= 0 Then
        If text = "" Then
            DbExec SqlMemoDelete(), Array(ToLong(rows(0)("メモID")))
        Else
            DbExec SqlMemoUpdate(), Array(text, Now(), ToLong(rows(0)("メモID")))
        End If
    ElseIf text <> "" Then
        DbExec SqlMemoInsert(), Array(d, tid, text, Now())
    End If
    DbCommit
    RememberPerson tid
    Go "entry.asp?d=" & YMD(d) & "&t=" & tid & "&m=" & mode & "&op=memo#memo"
End Sub

' ---------- 道具 ----------
Function CallVersion(day, t)
    Dim rows, s
    rows = DbQuery(SqlCallVersion(), Array(day, DateAdd("d", 1, day), t))
    s = ToLong(rows(0)("行数")) & "-" & ToLong(rows(0)("件数計"))
    If Not IsNull(rows(0)("最終")) Then s = s & "-" & YMDHM(rows(0)("最終")) & ":" & Pad2(Second(rows(0)("最終")))
    CallVersion = s
End Function

Function IsLocked(day)
    Dim rows
    rows = DbQuery(SqlDaily(), Array(day, DateAdd("d", 1, day)))
    IsLocked = False
    If UBound(rows) >= 0 Then
        If ToStr(rows(0)("状態")) = "確定" Then IsLocked = True
    End If
End Function

Function Registrant()
    If IsStaff() Then
        Registrant = "職員画面 " & Request.ServerVariables("REMOTE_ADDR")
    Else
        Registrant = "入力画面 " & Request.ServerVariables("REMOTE_ADDR")
    End If
End Function

Sub RememberPerson(t)
    Response.Cookies("nippou_t") = CStr(t)
    Response.Cookies("nippou_t").Expires = DateAdd("d", 400, Date())
    Response.Cookies("nippou_t").Path = "/"
End Sub

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

' 数えたあと・戻したあとの知らせ（押した欄のあるまとまりの見出しの下に出す）
Function CountNotice()
    Dim nm, now1
    CountNotice = ""
    If lastKey = "" Or mode <> "count" Then Exit Function
    If GridBlockOfKey(lastKey) = 0 Then Exit Function
    nm = H(GridCellName(lastKey))
    now1 = GridVal(existing, lastKey)
    If op = "add" Then
        CountNotice = "<div class=""msg msg-ok cnt-msg"">「" & nm & "」を 1 件数えました（いま " & now1 & " 件）。 " & _
                      "<button type=""submit"" name=""undo"" value=""" & lastKey & """ class=""btn btn-s btn-sub"">押し間違い：1 件戻す</button></div>"
    ElseIf op = "sub" Then
        CountNotice = "<div class=""msg msg-note cnt-msg"">「" & nm & "」を 1 件戻しました（いま " & now1 & " 件）。</div>"
    ElseIf op = "dup" Then
        CountNotice = "<div class=""msg msg-note cnt-msg"">画面が古かったので、いま押した分は数えていません（続けて 2 回押したときに、2 件に数えないためです）。「" & nm & "」はいま " & now1 & " 件です。もう 1 件あるときは、もう一度押してください。</div>"
    End If
End Function

' ---------- 画面 ----------
Sub ShowPersonBar()
    Dim j, label
    Response.Write "<form method=""post"" action=""entry.asp?m=" & mode & """ class=""bar noprint"">"
    Response.Write "<input type=""hidden"" name=""act"" value=""go"">"
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(DateAdd("d", -1, d)) & "&amp;t=" & tid & "&amp;m=" & mode & """>&lt; 前の日</a> "
    Response.Write "<label>日付 <input type=""date"" name=""d"" value=""" & YMD(d) & """ class=""in-date""></label> "
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(DateAdd("d", 1, d)) & "&amp;t=" & tid & "&amp;m=" & mode & """>次の日 &gt;</a> "
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(Date()) & "&amp;t=" & tid & "&amp;m=" & mode & """>今日</a> "
    Response.Write "<label>担当者 <select name=""t""><option value=""0"">（選んでください）</option>"
    For j = 0 To UBound(people)
        If PersonInPeriod(people(j), d) Or ToLong(people(j)("担当者ID")) = tid Then
            label = ToStr(people(j)("担当者コード")) & " " & ToStr(people(j)("氏名"))
            If ToStr(people(j)("職員区分")) = "職員" Then label = label & "（職員）"
            Response.Write "<option value=""" & ToLong(people(j)("担当者ID")) & """" & SelAttr(ToLong(people(j)("担当者ID")), tid) & ">" & H(label) & "</option>"
        End If
    Next
    Response.Write "</select></label> "
    Response.Write "<button type=""submit"" class=""btn btn-s"">表示</button>"
    Response.Write "<span class=""bar-date"">" & DateJa(d) & "</span>"
    Response.Write "</form>"
End Sub

' 入れ方の切り替え
Sub ShowModeTabs()
    Dim base
    base = "entry.asp?d=" & YMD(d) & "&amp;t=" & tid
    Response.Write "<p class=""tabs noprint"">"
    If mode = "count" Then
        Response.Write "<span class=""tab on"">1 件ずつ数える</span>"
        Response.Write "<a class=""tab"" href=""" & base & "&amp;m=edit"">数をまとめて入れる・直す</a>"
    Else
        Response.Write "<a class=""tab"" href=""" & base & "&amp;m=count"">1 件ずつ数える</a>"
        Response.Write "<span class=""tab on"">数をまとめて入れる・直す</span>"
    End If
    Response.Write "</p>"
End Sub

' この人のこの日の合計（日報の紙と同じ SQL で数える）
Sub ShowTotals()
    Dim cols, brk, j, tot, s, ex2, rf
    cols = DbQuery(SqlColumnTotalsOnePerson(), Array(d, DateAdd("d", 1, d), tid))
    brk = DbQuery(SqlBreakdownTotalsOnePerson(), Array(d, DateAdd("d", 1, d), tid))
    tot = 0
    s = ""
    For j = 0 To UBound(cols)
        tot = tot + ToLong(cols(j)("件数計"))
        If s <> "" Then s = s & " ／ "
        s = s & H(cols(j)("集計列名")) & " <span data-col=""" & ToLong(cols(j)("集計列ID")) & """>" & ToLong(cols(j)("件数計")) & "</span>"
    Next
    ex2 = 0
    rf = 0
    For j = 0 To UBound(brk)
        If ToStr(brk(j)("内訳区分")) = "交換" Then ex2 = ex2 + ToLong(brk(j)("件数計"))
        If ToStr(brk(j)("内訳区分")) = "返金" Then rf = rf + ToLong(brk(j)("件数計"))
    Next
    Response.Write "<div class=""totals"">" & H(PersonName(tid)) & " さんの " & DateShort(d) & " の件数：総合計 <b data-col=""total"">" & tot & "</b> 件　（日報の列：" & s & "）　内 交換 " & ex2 & "・返金 " & rf & "</div>"
End Sub

Sub ShowMemo(locked)
    Dim rows, v, dis
    v = memoText
    If v = "" Then
        rows = DbQuery(SqlMemoOnePerson(), Array(d, DateAdd("d", 1, d), tid))
        If UBound(rows) >= 0 Then v = ToStr(rows(0)("内容"))
    End If
    dis = ""
    If locked Then dis = " disabled"
    Response.Write "<section class=""gb"" id=""memo""><h2>特殊な問合せの内容（下記のとおり・別添不要のもの）</h2>"
    If op = "memo" Then Response.Write "<div class=""msg msg-ok"">内容を保存しました。</div>"
    Response.Write "<form method=""post"" action=""entry.asp?m=" & mode & """>"
    Response.Write "<input type=""hidden"" name=""act"" value=""memo""><input type=""hidden"" name=""d"" value=""" & YMD(d) & """><input type=""hidden"" name=""t"" value=""" & tid & """>"
    Response.Write "<p class=""note"">「【特殊な問合せ】下記のとおり（別添不要）」に数えたお問合せの中身を書きます（" & MEMO_MAX & " 文字まで）。職員が日報を書くときに、日報の画面で読めます。</p>"
    Response.Write "<textarea name=""memo"" rows=""3"" class=""memo""" & dis & ">" & H(v) & "</textarea>"
    If Not locked Then Response.Write "<div class=""actions""><button type=""submit"" class=""btn btn-sub"">内容を保存する</button></div>"
    Response.Write "</form></section>"
End Sub

PageHead "受付入力"
If op = "saved" Then msgOk = "保存しました。下の欄と、上の合計が、いま保存されている数です。"
If msgOk <> "" Then ShowMsg "ok", msgOk
If msgNg <> "" Then ShowMsg "ng", msgNg
ShowPersonBar
If UBound(people) < 0 Then
    ShowMsg "note", "担当者がまだ登録されていません。職員に「マスタ保守」で担当者を登録してもらってください。"
ElseIf tid = 0 Then
    ShowMsg "note", "上で「担当者」を選んで「表示」を押してください。"
Else
    ShowTotals
    If IsLocked(d) Then
        ShowMsg "note", "この日の日報は確定済みです。直すときは、職員に「確定の取り消し」を頼んでください。"
        Response.Write GridHtml("view", d, existing, posted, bad, "", "")
        ShowMemo True
    ElseIf mode = "count" Then
        ShowModeTabs
%>
<p class="note">電話を 1 本受けるたびに、その製品とお問合せ内容の欄を押してください。押すと 1 件数えて、欄の数が 1 つ増えます（その場で保存されます）。押し間違えたときは、出てくる「1 件戻す」を押してください。</p>
<form method="post" action="entry.asp" class="countform">
<input type="hidden" name="act" value="add">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<input type="hidden" name="t" value="<%= tid %>">
<input type="hidden" name="ver" value="<%= H(CallVersion(d, tid)) %>">
<% Response.Write GridHtml("count", d, existing, posted, bad, lastKey, CountNotice()) %>
</form>
<%
        ShowMemo False
    Else
        ShowModeTabs
%>
<form method="post" action="entry.asp?m=edit">
<input type="hidden" name="act" value="save">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<input type="hidden" name="t" value="<%= tid %>">
<input type="hidden" name="ver" value="<%= H(CallVersion(d, tid)) %>">
<p class="note">欄に今の件数を書いて、いちばん下の「保存する」を押してください。0 件の欄は空のままで構いません。数を直すときは、今の件数（合計）に書き換えます（足し算ではありません）。</p>
<% Response.Write GridHtml("edit", d, existing, posted, bad, "", "") %>
<div class="actions"><button type="submit" class="btn">保存する</button></div>
</form>
<%
        ShowMemo False
    End If
    Response.Write "<p class=""noprint""><a href=""tasks.asp?d=" & YMD(d) & "&amp;t=" & tid & """>電話以外の業務（①～⑬）の件数を入れる（その他業務）</a></p>"
End If
PageFoot
%>
