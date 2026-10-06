<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
%><!--#include file="include/config.asp"-->
<!--#include file="include/sql.asp"-->
<!--#include file="include/db.asp"-->
<!--#include file="include/auth.asp"-->
<!--#include file="include/layout.asp"-->
<!--#include file="include/sheet.asp"-->
<%
' ============================================================
'  マスタ保守（職員用）
'  担当者・区分・製品・ブロック・業務項目・集計列（日報の 5 列の名前）を登録・修正する。
'
'  ・消す機能は付けない。担当者は「在籍終了日」、ほかは「有効」を外すと入力画面の候補から消える。
'    過去の日報・集計表は、その担当者・区分のまま残る（参照整合性でも消せないようにしてある）。
'  ・区分の「日報の列」「内訳」を変えると、過去の日報の数も変わる（数は毎回 T_受電 から数え直すため）。
'    入力が既にある区分では、確認の印を付けないと変えられないようにしている。
' ============================================================
Dim tab, idParam, msgNg, msgOk, vals, bad, blocks, colsAll, j

RequireStaff
tab = QueryVal("t")
If tab <> "kubun" And tab <> "seihin" And tab <> "block" And tab <> "gyomu" And tab <> "col" And tab <> "stamp" Then tab = "tanto"
idParam = QueryVal("id")
msgNg = ""
msgOk = ""
If QueryVal("saved") = "1" Then msgOk = "保存しました。"
Set vals = Server.CreateObject("Scripting.Dictionary")
Set bad = Server.CreateObject("Scripting.Dictionary")
blocks = DbQuery(SqlBlocksAll(), Array())
colsAll = DbQuery(SqlColumnsAll(), Array())

If IsPost() Then
    idParam = FormVal("id")
    If tab = "tanto" Then SaveTanto
    If tab = "kubun" Then SaveKubun
    If tab = "seihin" Then SaveSeihin
    If tab = "block" Then SaveBlock
    If tab = "gyomu" Then SaveGyomu
    If tab = "col" Then SaveCol
    If tab = "stamp" Then SaveStamp
End If

' ============================================================
'  入力欄の道具
' ============================================================
Function FVal(name)
    If vals.Exists(name) Then FVal = vals(name) Else FVal = ""
End Function

Function NgC(name)
    If bad.Exists(name) Then NgC = " class=""ng""" Else NgC = ""
End Function

Sub Need(name, label, lim)
    If FVal(name) = "" Then
        bad(name) = True
        AddNg label & " を入れてください。"
    ElseIf Len(FVal(name)) > lim Then
        bad(name) = True
        AddNg label & " は " & lim & " 文字までです。"
    End If
End Sub

Sub MaxLen(name, label, lim)
    If Len(FVal(name)) > lim Then
        bad(name) = True
        AddNg label & " は " & lim & " 文字までです。"
    End If
End Sub

Sub AddNg(text)
    If msgNg <> "" Then msgNg = msgNg & " "
    msgNg = msgNg & text
End Sub

' 日付の欄：空なら Null、読めなければ誤り
Function DateOrNull(name, label)
    Dim dt
    DateOrNull = Null
    If FVal(name) <> "" Then
        If ParseYMD(FVal(name), dt) Then
            DateOrNull = dt
        Else
            bad(name) = True
            AddNg label & " が日付として読めません（例 2026-04-01）。"
        End If
    End If
End Function

Function OrderValue(name, dflt)
    Dim n
    OrderValue = dflt
    If FVal(name) <> "" Then
        If ParseCount(FVal(name), n) Then
            OrderValue = n
        Else
            bad(name) = True
            AddNg "表示順は 0～9999 の数字で入れてください。"
        End If
    End If
End Function

Function DateText(v)
    If IsNull(v) Then DateText = "" Else DateText = YMD(v)
End Function

Function BoolText(v)
    If v Then BoolText = "1" Else BoolText = ""
End Function

Function NextId(sql)
    NextId = ToLong(DbScalar(sql, Array(), 0)) + 1
End Function

Sub FormStart(title)
    Response.Write "<h2>" & H(title) & "</h2><form method=""post"" action=""master.asp?t=" & tab & """>"
    Response.Write "<input type=""hidden"" name=""id"" value=""" & H(idParam) & """><table class=""tw"">"
End Sub

Sub FormEnd()
    Response.Write "</table><div class=""actions""><button type=""submit"" class=""btn"">保存する</button> "
    Response.Write "<a class=""btn btn-sub"" href=""master.asp?t=" & tab & """>やめる（一覧へ）</a></div></form>"
End Sub

Sub TextRow(label, name, size, lim, note)
    Response.Write "<tr><th>" & H(label) & "</th><td><input type=""text"" name=""" & name & """ value=""" & H(FVal(name)) & """ size=""" & size & """ maxlength=""" & lim & """" & NgC(name) & ">"
    If note <> "" Then Response.Write " <span class=""note"">" & H(note) & "</span>"
    Response.Write "</td></tr>"
End Sub

Sub DateRow(label, name, note)
    Response.Write "<tr><th>" & H(label) & "</th><td><input type=""date"" name=""" & name & """ value=""" & H(FVal(name)) & """ class=""in-date""" & NgC(name) & ">"
    If note <> "" Then Response.Write " <span class=""note"">" & H(note) & "</span>"
    Response.Write "</td></tr>"
End Sub

Sub CheckRow(label, name, note)
    Response.Write "<tr><th>" & H(label) & "</th><td><label><input type=""checkbox"" name=""" & name & """ value=""1""" & ChkAttr(FVal(name) = "1") & "> " & H(note) & "</label></td></tr>"
End Sub

Sub SelectRow(label, name, opts)
    Dim k
    Response.Write "<tr><th>" & H(label) & "</th><td><select name=""" & name & """" & NgC(name) & ">"
    For k = 0 To UBound(opts)
        Response.Write "<option value=""" & H(opts(k)(0)) & """" & SelAttr(opts(k)(0), FVal(name)) & ">" & H(opts(k)(1)) & "</option>"
    Next
    Response.Write "</select></td></tr>"
End Sub

Function BlockOptions(onlyProductBlocks)
    Dim out, n, k
    ReDim out(-1)
    n = -1
    For k = 0 To UBound(blocks)
        If blocks(k)("製品別") Or Not onlyProductBlocks Then
            n = n + 1
            ReDim Preserve out(n)
            out(n) = Array(CStr(ToLong(blocks(k)("ブロックID"))), ToStr(blocks(k)("ブロック名")))
        End If
    Next
    BlockOptions = out
End Function

Function ColumnOptions()
    Dim out, k
    ReDim out(UBound(colsAll))
    For k = 0 To UBound(colsAll)
        out(k) = Array(CStr(ToLong(colsAll(k)("集計列ID"))), ToStr(colsAll(k)("集計列名")))
    Next
    ColumnOptions = out
End Function

Function BlockName(bid)
    Dim k
    BlockName = ""
    For k = 0 To UBound(blocks)
        If ToLong(blocks(k)("ブロックID")) = bid Then BlockName = ToStr(blocks(k)("ブロック名"))
    Next
End Function

Function ColumnName(cid)
    Dim k
    ColumnName = ""
    For k = 0 To UBound(colsAll)
        If ToLong(colsAll(k)("集計列ID")) = cid Then ColumnName = ToStr(colsAll(k)("集計列名"))
    Next
End Function

Function IsEdit()
    IsEdit = IsDigits(idParam)
End Function

Function EditId()
    EditId = ParamLong(idParam, 0)
End Function

' ============================================================
'  担当者
' ============================================================
Sub SaveTanto()
    Dim kind, st, en, ord, nm, newId, f
    For Each f In Array("code", "sei", "mei", "shimei", "kana", "logon", "kind", "start", "end", "order", "active", "note")
        vals(f) = FormVal(f)
    Next
    vals("note") = FormText("note")
    Need "code", "担当者コード（席番号）", 10
    Need "sei", "姓", 50
    MaxLen "mei", "名", 50
    MaxLen "kana", "カナ", 100
    MaxLen "logon", "ログオン名", 50
    nm = FVal("shimei")
    If nm = "" And FVal("sei") <> "" Then
        nm = FVal("sei")
        If FVal("mei") <> "" Then nm = nm & "　" & FVal("mei")
        vals("shimei") = nm
    End If
    MaxLen "shimei", "氏名（帳票に出す名前）", SHEET_NAME_MAX
    kind = FVal("kind")
    If kind <> "パート" And kind <> "職員" Then
        bad("kind") = True
        AddNg "職員区分を選んでください。"
    End If
    st = DateOrNull("start", "在籍開始日")
    en = DateOrNull("end", "在籍終了日")
    If Not IsNull(st) And Not IsNull(en) Then
        If en < st Then
            bad("end") = True
            AddNg "在籍終了日が在籍開始日より前です。"
        End If
    End If
    ord = OrderValue("order", -1)
    If msgNg <> "" Then Exit Sub
    DbBegin
    If IsEdit() Then
        If ord = -1 Then ord = 0
        DbExec SqlStaffUpdate(), Array(FVal("code"), FVal("sei"), FVal("mei"), nm, FVal("kana"), FVal("logon"), kind, st, en, ord, (FVal("active") = "1"), FVal("note"), EditId())
    Else
        newId = NextId(SqlStaffMaxId())
        If ord = -1 Then ord = newId * 10
        DbExec SqlStaffInsert(), Array(newId, FVal("code"), FVal("sei"), FVal("mei"), nm, FVal("kana"), FVal("logon"), kind, st, en, ord, (FVal("active") = "1"), FVal("note"))
    End If
    DbCommit
    Go "master.asp?t=tanto&saved=1"
End Sub

Sub ShowTanto()
    Dim rows, k, r, cls, today, dupNote, k2
    rows = DbQuery(SqlStaffAll(), Array())
    today = Date()
    If idParam <> "" Then
        If msgNg = "" Then
            If IsEdit() Then
                For k = 0 To UBound(rows)
                    If ToLong(rows(k)("担当者ID")) = EditId() Then
                        Set r = rows(k)
                        vals("code") = ToStr(r("担当者コード"))
                        vals("sei") = ToStr(r("姓"))
                        vals("mei") = ToStr(r("名"))
                        vals("shimei") = ToStr(r("氏名"))
                        vals("kana") = ToStr(r("カナ"))
                        vals("logon") = ToStr(r("ログオン名"))
                        vals("kind") = ToStr(r("職員区分"))
                        vals("start") = DateText(r("在籍開始日"))
                        vals("end") = DateText(r("在籍終了日"))
                        vals("order") = CStr(ToLong(r("表示順")))
                        vals("active") = BoolText(r("有効"))
                        vals("note") = ToStr(r("備考"))
                    End If
                Next
            Else
                vals("kind") = "パート"
                vals("active") = "1"
                vals("start") = YMD(today)
            End If
        End If
        If IsEdit() Then FormStart "担当者を直す" Else FormStart "担当者を追加する"
        TextRow "担当者コード（席番号）", "code", 6, 10, "例 001。人が替わったら同じ番号を使って構いません"
        TextRow "姓", "sei", 10, 50, ""
        TextRow "名", "mei", 10, 50, ""
        TextRow "氏名（帳票に出す名前）", "shimei", 16, SHEET_NAME_MAX, "空なら「姓　名」にします。帳票の欄に収まるよう " & SHEET_NAME_MAX & " 文字まで"
        TextRow "カナ", "kana", 16, 100, ""
        SelectRow "職員区分", "kind", Array(Array("パート", "パート（電話オペレーター）"), Array("職員", "職員（正規職員）"))
        DateRow "在籍開始日", "start", ""
        DateRow "在籍終了日", "end", "辞めた人はここに最後の日を入れます（消さないでください）"
        TextRow "表示順", "order", 4, 4, "小さい順に並びます。空なら最後に"
        CheckRow "有効", "active", "入力画面の候補に出す"
        TextRow "ログオン名", "logon", 16, 50, "ふだんは空のまま（将来 Windows 認証に切り替えたとき用）"
        Response.Write "<tr><th>備考</th><td><textarea name=""note"" rows=""2"">" & H(FVal("note")) & "</textarea></td></tr>"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p><a class=""btn"" href=""master.asp?t=tanto&amp;id=new"">担当者を追加する</a></p>"
    Response.Write "<p class=""note"">辞めた人は「直す」から在籍終了日を入れてください。消すと過去の日報が作れなくなるため、消す機能はありません。</p>"
    Response.Write "<table class=""tw""><tr><th>コード</th><th>氏名</th><th>カナ</th><th>職員区分</th><th>在籍</th><th class=""c"">表示順</th><th class=""c"">有効</th><th></th></tr>"
    For k = 0 To UBound(rows)
        Set r = rows(k)
        cls = ""
        If Not r("有効") Then cls = " class=""off"""
        If Not IsNull(r("在籍終了日")) Then
            If r("在籍終了日") < today Then cls = " class=""off"""
        End If
        dupNote = ""
        For k2 = 0 To UBound(rows)
            If k2 <> k And ToStr(rows(k2)("担当者コード")) = ToStr(r("担当者コード")) And rows(k2)("有効") And r("有効") Then
                If IsNull(rows(k2)("在籍終了日")) And IsNull(r("在籍終了日")) Then dupNote = " <span class=""warn"">（同じコードの在籍者がいます）</span>"
            End If
        Next
        Response.Write "<tr" & cls & "><td>" & H(r("担当者コード")) & dupNote & "</td><td>" & H(r("氏名")) & "</td><td>" & H(r("カナ")) & "</td><td>" & H(r("職員区分")) & "</td>"
        Response.Write "<td>" & DateText(r("在籍開始日")) & " ～ " & DateText(r("在籍終了日")) & "</td><td class=""n"">" & ToLong(r("表示順")) & "</td>"
        Response.Write "<td class=""c"">" & BoolMark(r("有効")) & "</td><td><a href=""master.asp?t=tanto&amp;id=" & ToLong(r("担当者ID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
    If UBound(rows) < 0 Then ShowMsg "note", "まだ担当者がいません。「担当者を追加する」から登録してください。"
End Sub

Function BoolMark(v)
    If v Then BoolMark = "○" Else BoolMark = "－"
End Function

' ============================================================
'  区分
' ============================================================
Sub SaveKubun()
    Dim f, ord, bid, cid, cur, k, used, rows, newId
    For Each f In Array("block", "name", "col", "uchi", "order", "active", "old", "ok_past")
        vals(f) = FormVal(f)
    Next
    Need "name", "区分名", 100
    MaxLen "old", "旧転記名", 100
    bid = ParamLong(FVal("block"), 0)
    If BlockName(bid) = "" Then
        bad("block") = True
        AddNg "ブロックを選んでください。"
    End If
    cid = ParamLong(FVal("col"), 0)
    If ColumnName(cid) = "" Then
        bad("col") = True
        AddNg "日報の列を選んでください。"
    End If
    If FVal("uchi") <> "" And FVal("uchi") <> "交換" And FVal("uchi") <> "返金" Then
        bad("uchi") = True
        AddNg "内訳は「なし」「交換」「返金」から選んでください。"
    End If
    ord = OrderValue("order", -1)
    If IsEdit() And msgNg = "" Then
        ' 入力がある区分で、日報の列・内訳を変えるときは確認の印を求める（過去の日報の数が変わるため）
        rows = DbQuery(SqlKubunAll(), Array())
        For k = 0 To UBound(rows)
            If ToLong(rows(k)("区分ID")) = EditId() Then
                If ToLong(rows(k)("集計列ID")) <> cid Or ToStr(rows(k)("内訳区分")) <> FVal("uchi") Then
                    used = ToLong(DbScalar(SqlKubunUseCount(), Array(EditId()), 0))
                    If used > 0 And FVal("ok_past") <> "1" Then
                        bad("col") = True
                        AddNg "この区分には、もう " & used & " 行の入力があります。日報の列や内訳を変えると、過去の日報・集計表の数も変わります。それでよければ、下の「過去の数が変わってもよい」に印を付けて、もう一度保存してください。"
                    End If
                End If
            End If
        Next
    End If
    If msgNg <> "" Then Exit Sub
    DbBegin
    If IsEdit() Then
        If ord = -1 Then ord = 0
        DbExec SqlKubunUpdate(), Array(bid, FVal("name"), cid, FVal("uchi"), ord, (FVal("active") = "1"), FVal("old"), EditId())
    Else
        newId = NextId(SqlKubunMaxId())
        If ord = -1 Then ord = newId
        DbExec SqlKubunInsert(), Array(newId, bid, FVal("name"), cid, FVal("uchi"), ord, (FVal("active") = "1"), FVal("old"))
    End If
    DbCommit
    Go "master.asp?t=kubun&saved=1"
End Sub

Sub ShowKubun()
    Dim rows, k, r, b, cls
    rows = DbQuery(SqlKubunAll(), Array())
    If idParam <> "" Then
        If msgNg = "" Then
            If IsEdit() Then
                For k = 0 To UBound(rows)
                    If ToLong(rows(k)("区分ID")) = EditId() Then
                        Set r = rows(k)
                        vals("block") = CStr(ToLong(r("ブロックID")))
                        vals("name") = ToStr(r("区分名"))
                        vals("col") = CStr(ToLong(r("集計列ID")))
                        vals("uchi") = ToStr(r("内訳区分"))
                        vals("order") = CStr(ToLong(r("表示順")))
                        vals("active") = BoolText(r("有効"))
                        vals("old") = ToStr(r("旧転記名"))
                    End If
                Next
            Else
                vals("active") = "1"
            End If
        End If
        If IsEdit() Then FormStart "区分を直す" Else FormStart "区分を追加する"
        SelectRow "ブロック（入力画面のまとまり）", "block", BlockOptions(False)
        TextRow "区分名", "name", 30, 100, ""
        SelectRow "日報の列", "col", ColumnOptions()
        SelectRow "内訳", "uchi", Array(Array("", "なし"), Array("交換", "内 交換 に数える"), Array("返金", "内 返金 に数える"))
        TextRow "表示順", "order", 4, 4, "ブロックの中で小さい順に並びます"
        CheckRow "有効", "active", "入力画面に出す"
        TextRow "旧転記名", "old", 30, 100, "現行 Excel での書き方（過去データの突き合わせ用。空でも構いません）"
        If IsEdit() Then CheckRow "確認", "ok_past", "過去の数が変わってもよい（日報の列・内訳を変えるときだけ）"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p><a class=""btn"" href=""master.asp?t=kubun&amp;id=new"">区分を追加する</a></p>"
    Response.Write "<p class=""note"">「日報の列」は、その区分の件数を日報のどの列（申込・抽選…）に積むかです。使わなくなった区分は「有効」を外してください。</p>"
    For b = 0 To UBound(blocks)
        Response.Write "<h2>" & H(blocks(b)("ブロック名")) & "</h2>"
        Response.Write "<table class=""tw""><tr><th>区分名</th><th>日報の列</th><th>内訳</th><th class=""c"">表示順</th><th class=""c"">有効</th><th>旧転記名</th><th></th></tr>"
        For k = 0 To UBound(rows)
            Set r = rows(k)
            If ToLong(r("ブロックID")) = ToLong(blocks(b)("ブロックID")) Then
                cls = ""
                If Not r("有効") Then cls = " class=""off"""
                Response.Write "<tr" & cls & "><td>" & H(r("区分名")) & "</td><td>" & H(ColumnName(ToLong(r("集計列ID")))) & "</td><td>" & H(r("内訳区分")) & "</td>"
                Response.Write "<td class=""n"">" & ToLong(r("表示順")) & "</td><td class=""c"">" & BoolMark(r("有効")) & "</td><td>" & H(r("旧転記名")) & "</td>"
                Response.Write "<td><a href=""master.asp?t=kubun&amp;id=" & ToLong(r("区分ID")) & """>直す</a></td></tr>"
            End If
        Next
        Response.Write "</table>"
    Next
End Sub

' ============================================================
'  製品
' ============================================================
Sub SaveSeihin()
    Dim f, ord, bid, st, en, newId
    For Each f In Array("block", "name", "start", "end", "order", "active")
        vals(f) = FormVal(f)
    Next
    Need "name", "製品名", 100
    bid = ParamLong(FVal("block"), 0)
    If BlockName(bid) = "" Then
        bad("block") = True
        AddNg "ブロックを選んでください（「製品別」のブロックだけ選べます）。"
    End If
    st = DateOrNull("start", "販売開始日")
    en = DateOrNull("end", "販売終了日")
    If Not IsNull(st) And Not IsNull(en) Then
        If en < st Then
            bad("end") = True
            AddNg "販売終了日が販売開始日より前です。"
        End If
    End If
    ord = OrderValue("order", -1)
    If msgNg <> "" Then Exit Sub
    DbBegin
    If IsEdit() Then
        If ord = -1 Then ord = 0
        DbExec SqlProductUpdate(), Array(bid, FVal("name"), st, en, ord, (FVal("active") = "1"), EditId())
    Else
        newId = NextId(SqlProductMaxId())
        If ord = -1 Then ord = newId
        DbExec SqlProductInsert(), Array(newId, bid, FVal("name"), st, en, ord, (FVal("active") = "1"))
    End If
    DbCommit
    Go "master.asp?t=seihin&saved=1"
End Sub

Sub ShowSeihin()
    Dim rows, k, r, cls, opts
    rows = DbQuery(SqlProductsAll(), Array())
    opts = BlockOptions(True)
    If idParam <> "" Then
        If UBound(opts) < 0 Then
            ShowMsg "note", "「製品別」のブロックがありません。先に「ブロック」で「製品別」に印を付けてください。"
            Exit Sub
        End If
        If msgNg = "" Then
            If IsEdit() Then
                For k = 0 To UBound(rows)
                    If ToLong(rows(k)("製品ID")) = EditId() Then
                        Set r = rows(k)
                        vals("block") = CStr(ToLong(r("ブロックID")))
                        vals("name") = ToStr(r("製品名"))
                        vals("start") = DateText(r("適用開始日"))
                        vals("end") = DateText(r("適用終了日"))
                        vals("order") = CStr(ToLong(r("表示順")))
                        vals("active") = BoolText(r("有効"))
                    End If
                Next
            Else
                vals("active") = "1"
            End If
        End If
        If IsEdit() Then FormStart "製品を直す" Else FormStart "製品を追加する"
        SelectRow "ブロック", "block", opts
        TextRow "製品名", "name", 30, 100, ""
        DateRow "販売開始日", "start", "空ならいつでも"
        DateRow "販売終了日", "end", "この日を過ぎると入力画面の列から消えます"
        TextRow "表示順", "order", 4, 4, "左から小さい順に並びます"
        CheckRow "有効", "active", "入力画面に出す"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p><a class=""btn"" href=""master.asp?t=seihin&amp;id=new"">製品を追加する</a></p>"
    Response.Write "<p class=""note"">製品は、「製品別」のブロックの入力画面で、列として並びます。販売が終わった製品は「販売終了日」を入れるか「有効」を外してください。</p>"
    Response.Write "<table class=""tw""><tr><th>ブロック</th><th>製品名</th><th>販売期間</th><th class=""c"">表示順</th><th class=""c"">有効</th><th></th></tr>"
    For k = 0 To UBound(rows)
        Set r = rows(k)
        cls = ""
        If Not r("有効") Then cls = " class=""off"""
        Response.Write "<tr" & cls & "><td>" & H(BlockName(ToLong(r("ブロックID")))) & "</td><td>" & H(r("製品名")) & "</td>"
        Response.Write "<td>" & DateText(r("適用開始日")) & " ～ " & DateText(r("適用終了日")) & "</td><td class=""n"">" & ToLong(r("表示順")) & "</td>"
        Response.Write "<td class=""c"">" & BoolMark(r("有効")) & "</td><td><a href=""master.asp?t=seihin&amp;id=" & ToLong(r("製品ID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

' ============================================================
'  ブロック
' ============================================================
Sub SaveBlock()
    Dim f, ord, newId, cnt
    For Each f In Array("name", "byproduct", "order")
        vals(f) = FormVal(f)
    Next
    Need "name", "ブロック名", 50
    ord = OrderValue("order", -1)
    If IsEdit() And FVal("byproduct") <> "1" And msgNg = "" Then
        cnt = ToLong(DbScalar(SqlBlockProductRows(), Array(EditId()), 0))
        If cnt > 0 Then
            bad("byproduct") = True
            AddNg "このブロックには、製品を指定した入力が " & cnt & " 行あります。「製品別」を外すと、その数が入力画面に出せなくなるので外せません。"
        End If
    End If
    If msgNg <> "" Then Exit Sub
    DbBegin
    If IsEdit() Then
        If ord = -1 Then ord = 0
        DbExec SqlBlockUpdate(), Array(FVal("name"), (FVal("byproduct") = "1"), ord, EditId())
    Else
        newId = NextId(SqlBlockMaxId())
        If ord = -1 Then ord = newId
        DbExec SqlBlockInsert(), Array(newId, FVal("name"), (FVal("byproduct") = "1"), ord)
    End If
    DbCommit
    Go "master.asp?t=block&saved=1"
End Sub

Sub ShowBlock()
    Dim k, r
    If idParam <> "" Then
        If msgNg = "" Then
            If IsEdit() Then
                For k = 0 To UBound(blocks)
                    If ToLong(blocks(k)("ブロックID")) = EditId() Then
                        Set r = blocks(k)
                        vals("name") = ToStr(r("ブロック名"))
                        vals("byproduct") = BoolText(r("製品別"))
                        vals("order") = CStr(ToLong(r("表示順")))
                    End If
                Next
            End If
        End If
        If IsEdit() Then FormStart "ブロックを直す" Else FormStart "ブロックを追加する"
        TextRow "ブロック名", "name", 30, 50, "入力画面の見出しになります"
        CheckRow "製品別", "byproduct", "製品ごとに件数を入れる（製品が列として並ぶ）"
        TextRow "表示順", "order", 4, 4, "上から小さい順に並びます"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p><a class=""btn"" href=""master.asp?t=block&amp;id=new"">ブロックを追加する</a></p>"
    Response.Write "<table class=""tw""><tr><th>ブロック名</th><th class=""c"">製品別</th><th class=""c"">表示順</th><th></th></tr>"
    For k = 0 To UBound(blocks)
        Set r = blocks(k)
        Response.Write "<tr><td>" & H(r("ブロック名")) & "</td><td class=""c"">" & BoolMark(r("製品別")) & "</td><td class=""n"">" & ToLong(r("表示順")) & "</td>"
        Response.Write "<td><a href=""master.asp?t=block&amp;id=" & ToLong(r("ブロックID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

' ============================================================
'  業務項目（①～⑬）
' ============================================================
Sub SaveGyomu()
    Dim f, ord, newId
    For Each f In Array("no", "name", "label", "order", "active")
        vals(f) = FormVal(f)
    Next
    Need "no", "番号", 4
    Need "name", "項目名", 100
    If FVal("label") = "" Then vals("label") = FVal("name")
    MaxLen "label", "帳票に出す名前", SHEET_TASK_LABEL_MAX
    ord = OrderValue("order", -1)
    If msgNg <> "" Then Exit Sub
    DbBegin
    If IsEdit() Then
        If ord = -1 Then ord = 0
        DbExec SqlTaskItemUpdate(), Array(FVal("no"), FVal("name"), FVal("label"), ord, (FVal("active") = "1"), EditId())
    Else
        newId = NextId(SqlTaskItemMaxId())
        If ord = -1 Then ord = newId
        DbExec SqlTaskItemInsert(), Array(newId, FVal("no"), FVal("name"), FVal("label"), ord, (FVal("active") = "1"))
    End If
    DbCommit
    Go "master.asp?t=gyomu&saved=1"
End Sub

Sub ShowGyomu()
    Dim rows, k, r, cls, active
    rows = DbQuery(SqlTaskItemsAll(), Array())
    If idParam <> "" Then
        If msgNg = "" Then
            If IsEdit() Then
                For k = 0 To UBound(rows)
                    If ToLong(rows(k)("業務項目ID")) = EditId() Then
                        Set r = rows(k)
                        vals("no") = ToStr(r("番号"))
                        vals("name") = ToStr(r("項目名"))
                        vals("label") = ToStr(r("帳票表示名"))
                        vals("order") = CStr(ToLong(r("表示順")))
                        vals("active") = BoolText(r("有効"))
                    End If
                Next
            Else
                vals("active") = "1"
            End If
        End If
        If IsEdit() Then FormStart "業務項目を直す" Else FormStart "業務項目を追加する"
        TextRow "番号", "no", 3, 4, "例 ⑭"
        TextRow "項目名", "name", 30, 100, "入力画面に出す名前"
        TextRow "帳票に出す名前", "label", 30, SHEET_TASK_LABEL_MAX, "空なら項目名と同じ。帳票の欄に収まるよう " & SHEET_TASK_LABEL_MAX & " 文字まで"
        TextRow "表示順", "order", 4, 4, ""
        CheckRow "有効", "active", "入力画面と帳票に出す"
        FormEnd
        Exit Sub
    End If
    active = 0
    For k = 0 To UBound(rows)
        If rows(k)("有効") Then active = active + 1
    Next
    Response.Write "<p><a class=""btn"" href=""master.asp?t=gyomu&amp;id=new"">業務項目を追加する</a></p>"
    If active > 13 Then ShowMsg "ng", "有効な業務項目が " & active & " 個あります。帳票の欄は 13 個（①～⑬）までなので、14 個目からは帳票に出ません。"
    Response.Write "<table class=""tw""><tr><th class=""c"">番号</th><th>項目名</th><th>帳票に出す名前</th><th class=""c"">表示順</th><th class=""c"">有効</th><th></th></tr>"
    For k = 0 To UBound(rows)
        Set r = rows(k)
        cls = ""
        If Not r("有効") Then cls = " class=""off"""
        Response.Write "<tr" & cls & "><td class=""c"">" & H(r("番号")) & "</td><td>" & H(r("項目名")) & "</td><td>" & H(r("帳票表示名")) & "</td>"
        Response.Write "<td class=""n"">" & ToLong(r("表示順")) & "</td><td class=""c"">" & BoolMark(r("有効")) & "</td>"
        Response.Write "<td><a href=""master.asp?t=gyomu&amp;id=" & ToLong(r("業務項目ID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

' ============================================================
'  集計列（日報の 5 列の名前だけ直せる）
' ============================================================
Sub SaveCol()
    vals("name") = FormVal("name")
    Need "name", "列の名前", SHEET_COL_NAME_MAX
    If Not IsEdit() Then AddNg "日報の列は 5 つで決まっています。名前を直すことだけができます。"
    If msgNg <> "" Then Exit Sub
    DbExec SqlColumnRename(), Array(FVal("name"), EditId())
    Go "master.asp?t=col&saved=1"
End Sub

Sub ShowCol()
    Dim k, r
    If idParam <> "" And IsEdit() Then
        If msgNg = "" Then
            For k = 0 To UBound(colsAll)
                If ToLong(colsAll(k)("集計列ID")) = EditId() Then vals("name") = ToStr(colsAll(k)("集計列名"))
            Next
        End If
        FormStart "日報の列の名前を直す"
        TextRow "列の名前", "name", 12, SHEET_COL_NAME_MAX, "帳票の見出しになります（" & SHEET_COL_NAME_MAX & " 文字まで）"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p class=""note"">日報の問合せ件数の 5 列です。列を増やしたり減らしたりはできません（帳票の形が決まっているため）。名前だけ直せます。</p>"
    Response.Write "<table class=""tw""><tr><th class=""c"">順</th><th>列の名前</th><th></th></tr>"
    For k = 0 To UBound(colsAll)
        Set r = colsAll(k)
        Response.Write "<tr><td class=""n"">" & ToLong(r("表示順")) & "</td><td>" & H(r("集計列名")) & "</td><td><a href=""master.asp?t=col&amp;id=" & ToLong(r("集計列ID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

' ============================================================
'  回覧の欄（帳票の上の押印欄の名前だけ直せる）
' ============================================================
Sub SaveStamp()
    vals("name") = FormVal("name")
    MaxLen "name", "押印欄の名前", SHEET_STAMP_NAME_MAX
    If Not IsEdit() Then AddNg "回覧の欄は 8 つ（左 4・右 4）で決まっています。名前を直すことだけができます。"
    If msgNg <> "" Then Exit Sub
    DbExec SqlStampRename(), Array(FVal("name"), EditId())
    Go "master.asp?t=stamp&saved=1"
End Sub

Sub ShowStamp()
    Dim rows, k, r
    rows = DbQuery(SqlStampsAll(), Array())
    If idParam <> "" And IsEdit() Then
        If msgNg = "" Then
            For k = 0 To UBound(rows)
                If ToLong(rows(k)("回覧ID")) = EditId() Then vals("name") = ToStr(rows(k)("表示名"))
            Next
        End If
        FormStart "回覧の欄の名前を直す"
        TextRow "押印欄の名前", "name", 10, SHEET_STAMP_NAME_MAX, "例 藤本課長（" & SHEET_STAMP_NAME_MAX & " 文字まで。空にすると枠だけ出ます）"
        FormEnd
        Exit Sub
    End If
    Response.Write "<p class=""note"">日報の上の「回覧」の押印欄です。左に 4 つ、右に 4 つ並びます。人が替わったら名前を直してください。</p>"
    Response.Write "<table class=""tw""><tr><th class=""c"">場所</th><th>押印欄の名前</th><th></th></tr>"
    For k = 0 To UBound(rows)
        Set r = rows(k)
        Response.Write "<tr><td>" & StampPlace(k) & "</td><td>" & HCell(r("表示名")) & "</td><td><a href=""master.asp?t=stamp&amp;id=" & ToLong(r("回覧ID")) & """>直す</a></td></tr>"
    Next
    Response.Write "</table>"
End Sub

Function StampPlace(k)
    If k < 4 Then
        StampPlace = "左の " & (k + 1) & " 番目"
    Else
        StampPlace = "右の " & (k - 3) & " 番目"
    End If
End Function

Sub Tabs()
    Dim names, labels, k
    names = Array("tanto", "kubun", "seihin", "block", "gyomu", "col", "stamp")
    labels = Array("担当者", "区分", "製品", "ブロック", "業務項目（①～⑬）", "日報の列", "回覧の欄")
    Response.Write "<p class=""noprint"">"
    For k = 0 To UBound(names)
        If names(k) = tab Then
            Response.Write "<a class=""btn btn-s"" href=""master.asp?t=" & names(k) & """>" & labels(k) & "</a> "
        Else
            Response.Write "<a class=""btn btn-s btn-sub"" href=""master.asp?t=" & names(k) & """>" & labels(k) & "</a> "
        End If
    Next
    Response.Write "</p>"
End Sub

PageHead "マスタ保守"
If msgOk <> "" Then ShowMsg "ok", msgOk
If msgNg <> "" Then ShowMsg "ng", msgNg & "（まだ保存していません）"
Tabs
If tab = "tanto" Then ShowTanto
If tab = "kubun" Then ShowKubun
If tab = "seihin" Then ShowSeihin
If tab = "block" Then ShowBlock
If tab = "gyomu" Then ShowGyomu
If tab = "col" Then ShowCol
If tab = "stamp" Then ShowStamp
PageFoot
%>
