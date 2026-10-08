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
'  受付入力：1 人・1 日ぶんの電話の件数を、製品ごと・お問合せ内容ごとに、数字で直接入れる。
'
'  表の形は現行 Excel の「記入用フォーム」と同じ（行が製品・列がお問合せ内容。include/grid.asp）。
'  特殊な問合せの内容（Excel の「下記のとおり」の欄）も同じ画面で書き、「保存する」1 つで件数と一緒に保存する。
'
'  ・保存は「その欄の今の件数」で上書きする（足し込まない）。0 にした欄は行を消す。
'    前の日の数が残ったり二重に数えたりしないように。
'  ・誤りが 1 つでもあれば、何も保存しない（件数の一部やメモだけが保存されて、数が合わなくなることがないように）。
'  ・誰の数かは「担当者ID」で持つ（名前の文字では突き合わせない）。
'    現行 Excel ではファイルの中の名前がずれていて、その人の数が黙って飛ばされた。
'  ・画面の合計は、日報の紙と同じ SQL（sql.asp の SqlCallsBase）で数える。
' ============================================================
Const MEMO_MAX = 400
Dim d, tid, msgNg, msgOk, people, existing, posted, bad, dd, cookieVal, op, memoText, memoPosted

d = PageDate()
tid = ParamLong(QueryVal("t"), 0)
If tid = 0 Then
    cookieVal = Request.Cookies("nippou_t")
    tid = ParamLong(cookieVal, 0)
End If
msgNg = ""
msgOk = ""
op = QueryVal("op")
memoText = ""
memoPosted = False
Set posted = Server.CreateObject("Scripting.Dictionary")
Set bad = Server.CreateObject("Scripting.Dictionary")

GridLoad

If IsPost() Then
    If FormVal("act") = "go" Then
        If ParseYMD(FormVal("d"), dd) Then
            Go "entry.asp?d=" & YMD(dd) & "&t=" & ParamLong(FormVal("t"), 0)
        End If
        msgNg = "日付が読めません。2026-08-25 のように入れてください。"
    ElseIf FormVal("act") = "save" Then
        SaveEntry
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

' ---------- 保存（件数と、特殊な問合せの内容） ----------
Sub SaveEntry()
    Dim ex, valid, key, raw, n, parts, stamp, who, errCount, parsed, text, memoRows, memoGiven
    If Not TakeDayPerson() Then Exit Sub
    ' 誤りで戻ったときに、書いた内容を消さずに出すため
    memoText = FormVal("memo")
    memoPosted = True
    If FormVal("ver") <> CallVersion(d, tid) Then
        msgNg = "この画面を開いている間に、別の画面でこの人のこの日の件数が変わりました。いま保存されている数を出し直しました。確かめて、もう一度入力・保存してください。"
        memoPosted = False
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
    ' 内容の欄が送られてこなかったとき（確定した日の画面など）は、内容には手を付けない
    memoGiven = Not IsEmpty(Request.Form("memo"))
    text = TrimAll(Replace(FormVal("memo"), vbCr, ""))
    If errCount > 0 Then
        msgNg = "赤い欄に、数字ではないもの（または 5 けた以上の数）が入っています。直してから、もう一度「保存する」を押してください。（まだ保存していません）"
        Exit Sub
    End If
    If Len(text) > MEMO_MAX Then
        msgNg = "特殊な問合せの内容は " & MEMO_MAX & " 文字までにしてください（いま " & Len(text) & " 文字）。（まだ保存していません）"
        Exit Sub
    End If
    text = Replace(text, vbLf, vbCrLf)
    stamp = Now()
    who = Registrant()
    memoRows = DbQuery(SqlMemoOnePerson(), Array(d, DateAdd("d", 1, d), tid))
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
    If Not memoGiven Then
        ' 何もしない
    ElseIf UBound(memoRows) >= 0 Then
        If text = "" Then
            DbExec SqlMemoDelete(), Array(ToLong(memoRows(0)("メモID")))
        ElseIf text <> ToStr(memoRows(0)("内容")) Then
            DbExec SqlMemoUpdate(), Array(text, stamp, ToLong(memoRows(0)("メモID")))
        End If
    ElseIf text <> "" Then
        DbExec SqlMemoInsert(), Array(d, tid, text, stamp)
    End If
    DbCommit
    RememberPerson tid
    Go "entry.asp?d=" & YMD(d) & "&t=" & tid & "&op=saved"
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

' ---------- 画面 ----------
Sub ShowPersonBar()
    Dim j, label
    Response.Write "<form method=""post"" action=""entry.asp"" class=""bar noprint"">"
    Response.Write "<input type=""hidden"" name=""act"" value=""go"">"
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(DateAdd("d", -1, d)) & "&amp;t=" & tid & """>&lt; 前の日</a> "
    Response.Write "<label>日付 <input type=""date"" name=""d"" value=""" & YMD(d) & """ class=""in-date""></label> "
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(DateAdd("d", 1, d)) & "&amp;t=" & tid & """>次の日 &gt;</a> "
    Response.Write "<a class=""btn btn-s"" href=""entry.asp?d=" & YMD(Date()) & "&amp;t=" & tid & """>今日</a> "
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

' 特殊な問合せの内容（現行 Excel の記入用フォームで「下記のとおり（別添不要）」の中身を書いていた欄）。
' 件数と同じフォームの中に置き、「保存する」1 つで一緒に保存する。
Sub ShowMemo(locked)
    Dim rows, v, dis
    If memoPosted Then
        v = memoText
    Else
        v = ""
        rows = DbQuery(SqlMemoOnePerson(), Array(d, DateAdd("d", 1, d), tid))
        If UBound(rows) >= 0 Then v = ToStr(rows(0)("内容"))
    End If
    dis = ""
    If locked Then dis = " disabled"
    Response.Write "<section class=""gb"" id=""memo""><h2>特殊な問合せの内容（下記のとおり・別添不要のもの）</h2>"
    Response.Write "<p class=""note"">「【特殊な問合せ】下記のとおり（別添不要）」に入れたお問合せの中身を書きます（" & MEMO_MAX & " 文字まで）。職員が日報を書くときに、日報の画面で読めます。</p>"
    Response.Write "<textarea name=""memo"" rows=""3"" class=""memo""" & dis & ">" & H(v) & "</textarea>"
    Response.Write "</section>"
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
        Response.Write GridHtml("view", d, existing, posted, bad)
        ShowMemo True
    Else
%>
<form method="post" action="entry.asp">
<input type="hidden" name="act" value="save">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<input type="hidden" name="t" value="<%= tid %>">
<input type="hidden" name="ver" value="<%= H(CallVersion(d, tid)) %>">
<p class="note">製品とお問合せ内容の欄に、その日の件数を数字で入れて「保存する」を押してください（「保存する」は上と下の 2 か所にあります。どちらを押しても同じです）。0 件の欄は空のままで構いません。数を直すときは、今の件数（合計）に書き換えます（足し算ではありません）。別の日・別の人に移る前に、保存してください。</p>
<div class="actions"><button type="submit" class="btn">保存する</button></div>
<% Response.Write GridHtml("edit", d, existing, posted, bad) %>
<% ShowMemo False %>
<div class="actions"><button type="submit" class="btn">保存する</button></div>
</form>
<%
    End If
    Response.Write "<p class=""noprint""><a href=""tasks.asp?d=" & YMD(d) & "&amp;t=" & tid & """>電話以外の業務（①～⑬）の件数を入れる（その他業務）</a></p>"
End If
PageFoot
%>
