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
'  受付入力：1 人・1 日ぶんの電話の件数を、区分（と製品）ごとに入れる。
'
'  ・誰の数かは「担当者ID」で持つ（名前の文字では突き合わせない）。
'    現行 Excel ではファイルの中の名前がずれていて、その人の数が黙って飛ばされた。
'  ・保存は「その欄の今の件数」で上書きする（足し込まない）。0 にした欄は行を消す。
'    前の日の数が残ったり二重に数えたりしないように。
'  ・画面の合計は、日報の紙と同じ SQL（sql.asp の SqlCallsBase）で数える。
' ============================================================
Dim d, tid, msgNg, saved, blocks, kubunAll, prodAll, colsAll, people, existing, posted, bad, parsed
Dim i, dd, cookieVal

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

blocks = DbQuery(SqlBlocksAll(), Array())
kubunAll = DbQuery(SqlKubunAll(), Array())
prodAll = DbQuery(SqlProductsAll(), Array())
colsAll = DbQuery(SqlColumnsAll(), Array())

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

Set existing = LoadExisting(d, tid)
people = DbQuery(SqlStaffAll(), Array())

' ---------- 保存 ----------
Sub SaveEntry()
    Dim t, ex, valid, key, raw, n, parts, stamp, who, errCount
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
    If FormVal("ver") <> CallVersion(d, t) Then
        msgNg = "この画面を開いている間に、別の画面でこの人のこの日の件数が保存されました。いま保存されている数を出し直しました。確かめて、もう一度入力・保存してください。"
        Exit Sub
    End If
    Set ex = LoadExisting(d, t)
    Set valid = ValidCells(d, ex)
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
            DbExec SqlCallInsert(), Array(d, t, CLng(parts(0)), CLng(parts(1)), n, stamp, stamp, who)
        End If
    Next
    DbCommit
    RememberPerson t
    Go "entry.asp?d=" & YMD(d) & "&t=" & t & "&saved=1"
End Sub

' ---------- 道具 ----------
Function CellKey(k, p)
    CellKey = CStr(ToLong(k)) & "_" & CStr(ToLong(p))
End Function

' 1 人・1 日ぶんの今の値。"区分ID_製品ID" → Array(受電ID, 件数)
Function LoadExisting(day, t)
    Dim rows, j, dict
    Set dict = Server.CreateObject("Scripting.Dictionary")
    If t <> 0 Then
        rows = DbQuery(SqlCallRowsOnePerson(), Array(day, DateAdd("d", 1, day), t))
        For j = 0 To UBound(rows)
            dict.Add CellKey(rows(j)("区分ID"), rows(j)("製品ID")), Array(ToLong(rows(j)("受電ID")), ToLong(rows(j)("件数")))
        Next
    End If
    Set LoadExisting = dict
End Function

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

Function InPeriod(row, day)
    InPeriod = True
    If Not IsNull(row("適用開始日")) Then
        If row("適用開始日") > day Then InPeriod = False
    End If
    If Not IsNull(row("適用終了日")) Then
        If row("適用終了日") < day Then InPeriod = False
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

' 区分 k にこの人のこの日の数があるか（製品は問わない）
Function HasKubunData(ex, k)
    Dim key
    HasKubunData = False
    For Each key In ex.Keys
        If Left(key, Len(CStr(k)) + 1) = CStr(k) & "_" Then HasKubunData = True
    Next
End Function

' ブロックに出す区分（使っている区分と、使わなくなったが数が残っている区分）
Function BlockKubun(blockId, ex)
    Dim out, n, j, k
    ReDim out(-1)
    n = -1
    For j = 0 To UBound(kubunAll)
        If ToLong(kubunAll(j)("ブロックID")) = blockId Then
            k = ToLong(kubunAll(j)("区分ID"))
            If kubunAll(j)("有効") Or HasKubunData(ex, k) Then
                n = n + 1
                ReDim Preserve out(n)
                Set out(n) = kubunAll(j)
            End If
        End If
    Next
    BlockKubun = out
End Function

' ブロックに出す製品の列。Array(製品ID, 名前, 今も使っているか) の配列。
' 製品別でないブロックは「件数」の 1 列（製品ID = 0）。
Function BlockProducts(brow, day, ex, kubunList)
    Dim out, n, j, pid, k, hasZero, used
    ReDim out(-1)
    n = -1
    If Not brow("製品別") Then
        ReDim out(0)
        out(0) = Array(0, "件数", True)
        BlockProducts = out
        Exit Function
    End If
    For j = 0 To UBound(prodAll)
        If ToLong(prodAll(j)("ブロックID")) = ToLong(brow("ブロックID")) Then
            pid = ToLong(prodAll(j)("製品ID"))
            used = False
            If prodAll(j)("有効") Then
                If InPeriod(prodAll(j), day) Then used = True
            End If
            If used Or ProductHasData(ex, kubunList, pid) Then
                n = n + 1
                ReDim Preserve out(n)
                out(n) = Array(pid, ToStr(prodAll(j)("製品名")), used)
            End If
        End If
    Next
    hasZero = ProductHasData(ex, kubunList, 0)
    If n = -1 Or hasZero Then
        n = n + 1
        ReDim Preserve out(n)
        out(n) = Array(0, "（製品指定なし）", True)
    End If
    BlockProducts = out
End Function

Function ProductHasData(ex, kubunList, pid)
    Dim j
    ProductHasData = False
    For j = 0 To UBound(kubunList)
        If ex.Exists(CellKey(kubunList(j)("区分ID"), pid)) Then ProductHasData = True
    Next
End Function

' 画面に出す欄の一覧（保存のときは、これに入っている欄だけを受け付ける）
Function ValidCells(day, ex)
    Dim dict, b, kl, pl, j, m
    Set dict = Server.CreateObject("Scripting.Dictionary")
    For b = 0 To UBound(blocks)
        kl = BlockKubun(ToLong(blocks(b)("ブロックID")), ex)
        pl = BlockProducts(blocks(b), day, ex, kl)
        For j = 0 To UBound(kl)
            For m = 0 To UBound(pl)
                dict(CellKey(kl(j)("区分ID"), pl(m)(0))) = True
            Next
        Next
    Next
    Set ValidCells = dict
End Function

Function ColumnName(colId)
    Dim j
    ColumnName = ""
    For j = 0 To UBound(colsAll)
        If ToLong(colsAll(j)("集計列ID")) = colId Then ColumnName = ToStr(colsAll(j)("集計列名"))
    Next
End Function

' 欄の値（誤りで戻ってきたときは、打った文字をそのまま出す）
Function CellValue(key)
    If posted.Exists(key) Then
        CellValue = posted(key)
    ElseIf existing.Exists(key) Then
        CellValue = CStr(existing(key)(1))
    Else
        CellValue = ""
    End If
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

Function PersonName(t)
    Dim j
    PersonName = ""
    For j = 0 To UBound(people)
        If ToLong(people(j)("担当者ID")) = t Then PersonName = ToStr(people(j)("氏名"))
    Next
End Function

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
    Response.Write "<div class=""totals"">" & H(PersonName(tid)) & " さんの " & DateShort(d) & " の件数（保存済みの数）：合計 <b data-col=""total"">" & tot & "</b>　（" & s & "）　内 交換 " & ex2 & "・返金 " & rf & "</div>"
End Sub

Sub ShowGrid(locked)
    Dim b, kl, pl, j, m, key, rowSum, colSum(), v, cls, dis, kname, grand
    dis = ""
    If locked Then dis = " disabled"
    For b = 0 To UBound(blocks)
        kl = BlockKubun(ToLong(blocks(b)("ブロックID")), existing)
        If UBound(kl) >= 0 Then
            pl = BlockProducts(blocks(b), d, existing, kl)
            ReDim colSum(UBound(pl))
            For m = 0 To UBound(pl)
                colSum(m) = 0
            Next
            grand = 0
            Response.Write "<h2>" & H(blocks(b)("ブロック名")) & "</h2>"
            Response.Write "<table class=""tw grid""><tr><th class=""kname"">区分（日報の列）</th>"
            For m = 0 To UBound(pl)
                If pl(m)(2) Then
                    Response.Write "<th class=""c"">" & H(pl(m)(1)) & "</th>"
                Else
                    Response.Write "<th class=""c"">" & H(pl(m)(1)) & "<br><span class=""note"">（今は使っていない製品）</span></th>"
                End If
            Next
            If UBound(pl) > 0 Then Response.Write "<th class=""c"">計</th>"
            Response.Write "</tr>"
            For j = 0 To UBound(kl)
                rowSum = 0
                kname = H(kl(j)("区分名")) & " <span class=""note"">→" & H(ColumnName(ToLong(kl(j)("集計列ID"))))
                If ToStr(kl(j)("内訳区分")) <> "" Then kname = kname & "・内 " & H(kl(j)("内訳区分"))
                kname = kname & "</span>"
                If Not kl(j)("有効") Then
                    Response.Write "<tr class=""off""><th class=""kname"">" & kname & "<br><span class=""note"">（今は使っていない区分）</span></th>"
                Else
                    Response.Write "<tr><th class=""kname"">" & kname & "</th>"
                End If
                For m = 0 To UBound(pl)
                    key = CellKey(kl(j)("区分ID"), pl(m)(0))
                    v = CellValue(key)
                    cls = "num"
                    If bad.Exists(key) Then cls = "num ng"
                    Response.Write "<td><input type=""text"" class=""" & cls & """ name=""c_" & key & """ value=""" & H(v) & """ inputmode=""numeric"" maxlength=""4"" size=""4""" & dis & "></td>"
                    If existing.Exists(key) Then
                        rowSum = rowSum + existing(key)(1)
                        colSum(m) = colSum(m) + existing(key)(1)
                        grand = grand + existing(key)(1)
                    End If
                Next
                If UBound(pl) > 0 Then Response.Write "<td class=""n"">" & rowSum & "</td>"
                Response.Write "</tr>"
            Next
            Response.Write "<tr class=""sum""><td>計（保存済み）</td>"
            For m = 0 To UBound(pl)
                Response.Write "<td class=""n"">" & colSum(m) & "</td>"
            Next
            If UBound(pl) > 0 Then Response.Write "<td class=""n"">" & grand & "</td>"
            Response.Write "</tr></table>"
        End If
    Next
End Sub

PageHead "受付入力"
If saved Then ShowMsg "ok", "保存しました。下の「件数」と、上の合計が、いま保存されている数です。"
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
        ShowGrid True
    Else
%>
<form method="post" action="entry.asp">
<input type="hidden" name="act" value="save">
<input type="hidden" name="d" value="<%= YMD(d) %>">
<input type="hidden" name="t" value="<%= tid %>">
<input type="hidden" name="ver" value="<%= H(CallVersion(d, tid)) %>">
<p class="note">その日に受けた電話の件数を入れて、いちばん下の「保存する」を押してください。0 件の欄は空のままで構いません。数を直すときは、今の件数（合計）に書き換えます（足し算ではありません）。</p>
<% ShowGrid False %>
<div class="actions"><button type="submit" class="btn">保存する</button></div>
</form>
<%
    End If
End If
PageFoot
%>
