<%
' ============================================================
'  受付の表（製品 × お問合せ内容）を組み立てる部品。
'
'  受付入力（entry.asp）と、個人別の受付表（detail.asp・PDF）が、同じこの部品を使う。
'  並びは現行 Excel の「記入用フォーム」と同じにしてある（使う人が迷わないように）:
'    ・製品別のまとまり … 行が製品、列がお問合せ内容。右端が件数、下が計。
'    ・製品別でないまとまり … お問合せ内容を横に並べ、その下に件数。
'    ・区分名の「／」の前が同じ区分が続くときは、見出しを 2 段にする
'      （例「払込用紙／再発行（可）」「払込用紙／発送状況」→ 上の段に「払込用紙」を 1 つ）。
'  表の数は、画面を開くたびに T_受電 から読み直したもの（Excel のように欄に数を持ち続けない）。
'
'  mode:
'    "edit"  … 欄が入力欄（受付入力。数字を直接入れる）
'    "view"  … 数だけ（確定した日の受付入力・個人別の受付表の画面）
'    "print" … 数だけ・日報の列の案内なし（個人別の受付表の PDF）
' ============================================================
Dim gBlocks, gKubun, gProd, gCols

Sub GridLoad()
    gBlocks = DbQuery(SqlBlocksAll(), Array())
    gKubun = DbQuery(SqlKubunAll(), Array())
    gProd = DbQuery(SqlProductsAll(), Array())
    gCols = DbQuery(SqlColumnsAll(), Array())
End Sub

Function GridKey(k, p)
    GridKey = CStr(ToLong(k)) & "_" & CStr(ToLong(p))
End Function

' 1 人・1 日ぶんの今の値。"区分ID_製品ID" → Array(受電ID, 件数)
Function GridLoadCalls(day, t)
    Dim rows, j, dict
    Set dict = Server.CreateObject("Scripting.Dictionary")
    If t <> 0 Then
        rows = DbQuery(SqlCallRowsOnePerson(), Array(day, DateAdd("d", 1, day), t))
        For j = 0 To UBound(rows)
            dict.Add GridKey(rows(j)("区分ID"), rows(j)("製品ID")), Array(ToLong(rows(j)("受電ID")), ToLong(rows(j)("件数")))
        Next
    End If
    Set GridLoadCalls = dict
End Function

Function GridVal(ex, key)
    GridVal = 0
    If ex.Exists(key) Then GridVal = ex(key)(1)
End Function

Function GridInPeriod(row, day)
    GridInPeriod = True
    If Not IsNull(row("適用開始日")) Then
        If row("適用開始日") > day Then GridInPeriod = False
    End If
    If Not IsNull(row("適用終了日")) Then
        If row("適用終了日") < day Then GridInPeriod = False
    End If
End Function

' 区分 k にこの人のこの日の数があるか（製品は問わない）
Function GridKubunHasData(ex, k)
    Dim key
    GridKubunHasData = False
    For Each key In ex.Keys
        If Left(key, Len(CStr(k)) + 1) = CStr(k) & "_" Then GridKubunHasData = True
    Next
End Function

' まとまりに出す区分（使っている区分と、使わなくなったが数が残っている区分）
Function GridBlockKubun(blockId, ex)
    Dim out, n, j, k
    ReDim out(-1)
    n = -1
    For j = 0 To UBound(gKubun)
        If ToLong(gKubun(j)("ブロックID")) = blockId Then
            k = ToLong(gKubun(j)("区分ID"))
            If gKubun(j)("有効") Or GridKubunHasData(ex, k) Then
                n = n + 1
                ReDim Preserve out(n)
                Set out(n) = gKubun(j)
            End If
        End If
    Next
    GridBlockKubun = out
End Function

Function GridProductHasData(ex, kubunList, pid)
    Dim j
    GridProductHasData = False
    For j = 0 To UBound(kubunList)
        If ex.Exists(GridKey(kubunList(j)("区分ID"), pid)) Then GridProductHasData = True
    Next
End Function

' 製品別のまとまりに出す製品の行。Array(製品ID, 名前, 今も使っているか) の配列。
' 製品が 1 つも無いまとまりと、製品ID = 0 に数があるときは「（製品指定なし）」の行を出す。
Function GridBlockProducts(brow, day, ex, kubunList)
    Dim out, n, j, pid, used
    ReDim out(-1)
    n = -1
    For j = 0 To UBound(gProd)
        If ToLong(gProd(j)("ブロックID")) = ToLong(brow("ブロックID")) Then
            pid = ToLong(gProd(j)("製品ID"))
            used = False
            If gProd(j)("有効") Then
                If GridInPeriod(gProd(j), day) Then used = True
            End If
            If used Or GridProductHasData(ex, kubunList, pid) Then
                n = n + 1
                ReDim Preserve out(n)
                out(n) = Array(pid, ToStr(gProd(j)("製品名")), used)
            End If
        End If
    Next
    If n = -1 Or GridProductHasData(ex, kubunList, 0) Then
        n = n + 1
        ReDim Preserve out(n)
        out(n) = Array(0, "（製品指定なし）", True)
    End If
    GridBlockProducts = out
End Function

' 画面に出す欄の一覧（保存のときは、これに入っている欄だけを受け付ける）
Function GridValidCells(day, ex)
    Dim dict, b, kl, pl, j, m
    Set dict = Server.CreateObject("Scripting.Dictionary")
    For b = 0 To UBound(gBlocks)
        kl = GridBlockKubun(ToLong(gBlocks(b)("ブロックID")), ex)
        If gBlocks(b)("製品別") Then
            pl = GridBlockProducts(gBlocks(b), day, ex, kl)
            For j = 0 To UBound(kl)
                For m = 0 To UBound(pl)
                    dict(GridKey(kl(j)("区分ID"), pl(m)(0))) = True
                Next
            Next
        Else
            For j = 0 To UBound(kl)
                dict(GridKey(kl(j)("区分ID"), 0)) = True
            Next
        End If
    Next
    Set GridValidCells = dict
End Function

Function GridColumnName(colId)
    Dim j
    GridColumnName = ""
    For j = 0 To UBound(gCols)
        If ToLong(gCols(j)("集計列ID")) = colId Then GridColumnName = ToStr(gCols(j)("集計列名"))
    Next
End Function

' 欄の名前（「製品・お問合せ内容」）。入力欄の説明（マウスを重ねたときに出る）に使う。
Function GridCellName(key)
    Dim parts, j, kname, pname
    parts = Split(key, "_")
    kname = ""
    pname = ""
    For j = 0 To UBound(gKubun)
        If ToLong(gKubun(j)("区分ID")) = CLng(parts(0)) Then kname = Replace(ToStr(gKubun(j)("区分名")), "／", " ")
    Next
    For j = 0 To UBound(gProd)
        If ToLong(gProd(j)("製品ID")) = CLng(parts(1)) Then pname = ToStr(gProd(j)("製品名"))
    Next
    If pname = "" Then
        GridCellName = kname
    Else
        GridCellName = pname & "・" & kname
    End If
End Function

' 区分名を上の段と下の段に分ける。「／」が無ければ上の段は ""。
Sub GridSplitName(nm, top, lower)
    Dim p
    p = InStr(nm, "／")
    If p > 0 Then
        top = Left(nm, p - 1)
        lower = Mid(nm, p + 1)
    Else
        top = ""
        lower = nm
    End If
End Sub

Function GridKubunLabel(row, lower, showHint)
    Dim s
    s = H(lower)
    If showHint Then
        s = s & "<span class=""to"">→" & H(GridColumnName(ToLong(row("集計列ID"))))
        If ToStr(row("内訳区分")) <> "" Then s = s & "・内 " & H(row("内訳区分"))
        s = s & "</span>"
    End If
    If Not row("有効") Then s = s & "<span class=""note"">（今は使っていない）</span>"
    GridKubunLabel = s
End Function

' 見出しの行（1 段か 2 段）
Function GridHeadHtml(kl, showHint, leadLabel, tailLabel)
    Dim j, k, n, two, top, lower, tops(), lowers(), s1, s2, rs, span
    n = UBound(kl)
    ReDim tops(n)
    ReDim lowers(n)
    two = False
    For j = 0 To n
        GridSplitName ToStr(kl(j)("区分名")), top, lower
        tops(j) = top
        lowers(j) = lower
        If top <> "" Then two = True
    Next
    rs = ""
    If two Then rs = " rowspan=""2"""
    s1 = "<tr>"
    If leadLabel <> "" Then s1 = s1 & "<th class=""pname""" & rs & ">" & H(leadLabel) & "</th>"
    s2 = ""
    j = 0
    Do While j <= n
        If tops(j) = "" Then
            s1 = s1 & "<th class=""kh""" & rs & ">" & GridKubunLabel(kl(j), lowers(j), showHint) & "</th>"
            j = j + 1
        Else
            span = 1
            Do While j + span <= n
                If tops(j + span) <> tops(j) Then Exit Do
                span = span + 1
            Loop
            s1 = s1 & "<th class=""kg"" colspan=""" & span & """>" & H(tops(j)) & "</th>"
            For k = j To j + span - 1
                s2 = s2 & "<th class=""kh"">" & GridKubunLabel(kl(k), lowers(k), showHint) & "</th>"
            Next
            j = j + span
        End If
    Loop
    If tailLabel <> "" Then s1 = s1 & "<th class=""tot""" & rs & ">" & H(tailLabel) & "</th>"
    s1 = s1 & "</tr>"
    If two Then s1 = s1 & "<tr>" & s2 & "</tr>"
    GridHeadHtml = s1
End Function

Function GridNum(v)
    If v = 0 Then
        GridNum = "&nbsp;"
    Else
        GridNum = CStr(v)
    End If
End Function

' 1 つの欄
Function GridCellHtml(mode, key, ex, posted, bad)
    Dim v, cls, raw
    v = GridVal(ex, key)
    If mode = "edit" Then
        If posted.Exists(key) Then
            raw = posted(key)
        ElseIf v = 0 Then
            raw = ""
        Else
            raw = CStr(v)
        End If
        cls = "num"
        If bad.Exists(key) Then cls = "num ng"
        GridCellHtml = "<td class=""c""><input type=""text"" class=""" & cls & """ name=""c_" & key & """ value=""" & H(raw) & """ inputmode=""numeric"" maxlength=""4"" size=""3"" title=""" & H(GridCellName(key)) & """></td>"
    Else
        GridCellHtml = "<td class=""n"">" & GridNum(v) & "</td>"
    End If
End Function

' すべてのまとまりの表
Function GridHtml(mode, day, ex, posted, bad)
    Dim b, bid, kl, pl, j, m, key, v, s, out, rowSum, colSum(), grand, showHint, pnote
    showHint = (mode <> "print")
    out = ""
    For b = 0 To UBound(gBlocks)
        bid = ToLong(gBlocks(b)("ブロックID"))
        kl = GridBlockKubun(bid, ex)
        If UBound(kl) >= 0 Then
            s = "<section class=""gb"" id=""b" & bid & """><h2>" & H(gBlocks(b)("ブロック名")) & "</h2>"
            s = s & "<div class=""gscroll""><table class=""tw grid"">"
            ReDim colSum(UBound(kl))
            For j = 0 To UBound(kl)
                colSum(j) = 0
            Next
            grand = 0
            If gBlocks(b)("製品別") Then
                pl = GridBlockProducts(gBlocks(b), day, ex, kl)
                s = s & GridHeadHtml(kl, showHint, "製品", "件数")
                For m = 0 To UBound(pl)
                    rowSum = 0
                    pnote = ""
                    If Not pl(m)(2) Then pnote = "<span class=""note"">（今は使っていない製品）</span>"
                    s = s & "<tr><th class=""pname"">" & H(pl(m)(1)) & pnote & "</th>"
                    For j = 0 To UBound(kl)
                        key = GridKey(kl(j)("区分ID"), pl(m)(0))
                        s = s & GridCellHtml(mode, key, ex, posted, bad)
                        v = GridVal(ex, key)
                        rowSum = rowSum + v
                        colSum(j) = colSum(j) + v
                        grand = grand + v
                    Next
                    s = s & "<td class=""n tot"">" & GridNum(rowSum) & "</td></tr>"
                Next
                s = s & "<tr class=""sum""><th class=""pname"">計</th>"
                For j = 0 To UBound(kl)
                    s = s & "<td class=""n"">" & colSum(j) & "</td>"
                Next
                s = s & "<td class=""n tot"" data-block=""" & bid & """>" & grand & "</td></tr>"
            Else
                s = s & GridHeadHtml(kl, showHint, "", "件数")
                s = s & "<tr>"
                For j = 0 To UBound(kl)
                    key = GridKey(kl(j)("区分ID"), 0)
                    s = s & GridCellHtml(mode, key, ex, posted, bad)
                    grand = grand + GridVal(ex, key)
                Next
                s = s & "<td class=""n tot"" data-block=""" & bid & """>" & grand & "</td></tr>"
            End If
            out = out & s & "</table></div></section>"
        End If
    Next
    GridHtml = out
End Function

' ------------------------------------------------------------
'  個人別の受付表（1 人 1 枚）。現行 Excel の「印刷」マクロが、入力のあった人の記入用フォームを
'  1 枚ずつ印刷していたものにあたる。画面（detail.asp）も PDF も、この関数で作る。
' ------------------------------------------------------------
Function PersonSheetHtml(day, person)
    Dim t, ex, empty1, att, hours, tasks, items, j, k, done, s, cols, tot, memo, line, n
    t = ToLong(person("担当者ID"))
    Set ex = GridLoadCalls(day, t)
    Set empty1 = Server.CreateObject("Scripting.Dictionary")
    hours = ""
    att = DbQuery(SqlAttendance(), Array(day, DateAdd("d", 1, day)))
    For j = 0 To UBound(att)
        If ToLong(att(j)("担当者ID")) = t Then hours = TrimAll(ToStr(att(j)("勤務時間")) & " " & ToStr(att(j)("備考")))
    Next
    s = "<div class=""psheet"">"
    s = s & "<div class=""ps-head""><span class=""ps-who"">担当者　" & H(person("担当者コード")) & "　" & H(person("氏名"))
    If ToStr(person("職員区分")) = "職員" Then s = s & "（職員）"
    s = s & "</span><span class=""ps-date"">" & Wareki(day) & "</span></div>"
    If hours <> "" Then s = s & "<div class=""ps-hours"">勤務時間・備考　" & H(hours) & "</div>"
    s = s & GridHtml("print", day, ex, empty1, empty1)
    ' 日報の 5 列に入る数（日報と同じ SQL で数える）
    cols = DbQuery(SqlColumnTotalsOnePerson(), Array(day, DateAdd("d", 1, day), t))
    tot = 0
    line = ""
    For j = 0 To UBound(cols)
        tot = tot + ToLong(cols(j)("件数計"))
        If line <> "" Then line = line & "　"
        line = line & H(cols(j)("集計列名")) & " " & ToLong(cols(j)("件数計"))
    Next
    s = s & "<div class=""ps-total"">総合計 <b data-total=""" & t & """>" & tot & "</b> 件　（" & line & "）</div>"
    memo = DbQuery(SqlMemoOnePerson(), Array(day, DateAdd("d", 1, day), t))
    If UBound(memo) >= 0 Then
        If ToStr(memo(0)("内容")) <> "" Then s = s & "<div class=""ps-memo""><b>特殊な問合せの内容</b><div class=""pre"">" & H(memo(0)("内容")) & "</div></div>"
    End If
    ' ※電話応対以外の業務（①～⑬）
    items = DbQuery(SqlTaskItemsAll(), Array())
    tasks = DbQuery(SqlTaskRowsOnePerson(), Array(day, DateAdd("d", 1, day), t))
    s = s & "<table class=""tw ps-tasks""><tr><th colspan=""6"">※電話応対以外の業務</th></tr>"
    n = 0
    For j = 0 To UBound(items)
        If items(j)("有効") Then
            If n Mod 3 = 0 Then s = s & "<tr>"
            done = 0
            For k = 0 To UBound(tasks)
                If ToLong(tasks(k)("業務項目ID")) = ToLong(items(j)("業務項目ID")) Then done = ToLong(tasks(k)("件数"))
            Next
            s = s & "<td>" & H(items(j)("番号")) & " " & H(items(j)("項目名")) & "</td><td class=""n"">" & GridNum(done) & " 件</td>"
            If n Mod 3 = 2 Then s = s & "</tr>"
            n = n + 1
        End If
    Next
    Do While n Mod 3 <> 0
        s = s & "<td>&nbsp;</td><td>&nbsp;</td>"
        If n Mod 3 = 2 Then s = s & "</tr>"
        n = n + 1
    Loop
    s = s & "</table></div>"
    PersonSheetHtml = s
End Function

' その日に入力のあった全員ぶん（PDF では 1 人 1 枚）
Function PersonSheetsHtml(day)
    Dim d2, people, j, s
    d2 = DateAdd("d", 1, day)
    people = DbQuery(SqlPeopleWithInput(), Array(day, d2, day, d2, day, d2))
    s = ""
    For j = 0 To UBound(people)
        s = s & PersonSheetHtml(day, people(j))
    Next
    PersonSheetsHtml = s
End Function
%>
