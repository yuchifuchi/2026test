<%
' ============================================================
'  帳票「電話応対報告書日報集計表」（A4 縦 1 枚）
'
'  画面で見る（report.asp）のも、PDF にする（printpdf.asp）のも、必ずこの SheetHtml(dt) を呼ぶ。
'  片方だけ直して食い違う事故を、作りの上で起こさないため。
'
'  罫線の決まり:
'    空の行・空の欄も、必ず &nbsp; を入れた <td> を置く（枠が途切れないように）。
'    罫線は css/style.css の .sheet の規則だけで引く。ここでは style 属性を使わない。
'
'  数はすべて sql.asp の SqlCallsBase（T_受電 ただ 1 つ）から計算したもの。
' ============================================================

Const SHEET_ATT_ROWS = 7      ' 出勤者の氏名欄の行数（2 列 × 7 行 ＝ 14 名まで）
Const SHEET_TASK_ROWS = 8     ' 電話対応以外の業務：左の段の行数（右の段は 5 行だけ使う）
Const SHEET_TASK_USED = 13    ' 電話対応以外の業務：使う欄の数（①～⑬）
' 帳票に出す文字の長さの上限（マスタ保守で止める）。文字を縮めて収めるのではなく、長さで収める。
' 値を変えたら tools/render_sheet.py で「上限いっぱいでも A4 1 枚・はみ出し無し」を確かめ直すこと。
Const SHEET_NAME_MAX = 12          ' 出勤者の氏名
Const SHEET_TASK_LABEL_MAX = 16    ' 電話対応以外の業務の名前
Const SHEET_COL_NAME_MAX = 6       ' 問合せ件数の列の名前

Function SheetHtml(dt)
    Dim d2, daily, att, cols, brk, tasks, s
    d2 = DateAdd("d", 1, dt)
    daily = DbQuery(SqlDaily(), Array(dt, d2))
    att = DbQuery(SqlAttendance(), Array(dt, d2))
    cols = DbQuery(SqlColumnTotals(), Array(dt, d2))
    brk = DbQuery(SqlBreakdownTotals(), Array(dt, d2))
    tasks = DbQuery(SqlTaskTotals(), Array(dt, d2))
    s = "<div class=""sheet"">"
    s = s & "<div class=""sh-limit"">【課内限り】</div>"
    s = s & "<div class=""sh-date"">" & H(Wareki(dt)) & "</div>"
    s = s & "<div class=""sh-title"">電話応対報告書日報集計表</div>"
    s = s & SheetAttendance(att)
    s = s & SheetLines(daily)
    s = s & SheetCounts(cols)
    s = s & SheetBreakdown(brk)
    s = s & SheetBox("特記事項", SheetDailyText(daily, "特記事項"))
    s = s & SheetBox("職員に代わった案件（概要）", SheetDailyText(daily, "職員代替案件"))
    s = s & SheetBox("要望", SheetDailyText(daily, "要望"))
    s = s & SheetTasks(tasks)
    s = s & "</div>"
    SheetHtml = s
End Function

' 出勤者：左 3 列（出勤者 / 人数 / 名）を縦に結合し、右に 氏名 2 列 ＋ 備考
Function SheetAttendance(att)
    Dim nm(13), nt(13), n, i, r, s, note
    n = UBound(att) + 1
    For i = 0 To 13
        nm(i) = ""
        nt(i) = ""
    Next
    For i = 0 To UBound(att)
        If i <= 13 Then
            nm(i) = ToStr(att(i)("氏名"))
            nt(i) = SheetAttNote(att(i))
        End If
    Next
    s = "<table class=""sh-att"">"
    s = s & "<colgroup><col class=""w-a1""><col class=""w-a2""><col class=""w-a3""><col class=""w-a4""><col class=""w-a4""><col class=""w-a5""></colgroup>"
    s = s & "<tr><th rowspan=""" & (SHEET_ATT_ROWS + 1) & """ class=""a-lbl"">出<br>勤<br>者</th>"
    s = s & "<td rowspan=""" & (SHEET_ATT_ROWS + 1) & """ class=""a-num"">" & n & "</td>"
    s = s & "<td rowspan=""" & (SHEET_ATT_ROWS + 1) & """ class=""a-mei"">名</td>"
    s = s & "<th class=""a-h"">氏　　名</th><th class=""a-h"">氏　　名</th><th class=""a-h"">備　　考</th></tr>"
    For r = 0 To SHEET_ATT_ROWS - 1
        ' 備考は、左の人と右の人を 1 行ずつ（2 人ぶんで 2 行。行の高さは 2 行ぶんに固定してある）
        note = H(nt(r))
        If nt(r + SHEET_ATT_ROWS) <> "" Then
            If note <> "" Then note = note & "<br>"
            note = note & H(nt(r + SHEET_ATT_ROWS))
        End If
        If note = "" Then note = "&nbsp;"
        s = s & "<tr><td class=""a-name"">" & HCell(nm(r)) & "</td>"
        s = s & "<td class=""a-name"">" & HCell(nm(r + SHEET_ATT_ROWS)) & "</td>"
        s = s & "<td class=""a-note"">" & note & "</td></tr>"
    Next
    s = s & "</table>"
    SheetAttendance = s
End Function

' 備考欄に出す 1 人ぶん（姓：勤務時間 備考）
Function SheetAttNote(row)
    Dim t
    t = TrimAll(ToStr(row("勤務時間")) & " " & ToStr(row("備考")))
    If t <> "" Then t = ToStr(row("姓")) & "：" & t
    SheetAttNote = t
End Function

Function SheetLines(daily)
    Dim n
    n = "　　"
    If UBound(daily) >= 0 Then
        If Not IsNull(daily(0)("回線数")) Then n = CStr(ToLong(daily(0)("回線数")))
    End If
    SheetLines = "<div class=""sh-lines"">（ " & H(n) & " 回線 ）</div>"
End Function

' 問合せ件数：左に合計を大きく、右に 5 列。各欄は大きい数字の下に（職員が受けた数）。
Function SheetCounts(cols)
    Dim i, tot, totSt, s, head, body
    tot = 0
    totSt = 0
    head = ""
    body = ""
    For i = 0 To UBound(cols)
        tot = tot + ToLong(cols(i)("件数計"))
        totSt = totSt + ToLong(cols(i)("職員計"))
        head = head & "<th class=""c-col"">" & HCell(cols(i)("集計列名")) & "</th>"
        body = body & "<td class=""c-col"" data-col=""" & H(cols(i)("集計列ID")) & """>" & SheetNum(ToLong(cols(i)("件数計")), ToLong(cols(i)("職員計"))) & "</td>"
    Next
    s = "<table class=""sh-cnt"">"
    s = s & "<tr><th rowspan=""2"" class=""c-q"">問合せ<br>件数</th><th class=""c-tot"">合　計</th>" & head & "</tr>"
    s = s & "<tr><td class=""c-tot"" data-col=""total"">" & SheetNum(tot, totSt) & "</td>" & body & "</tr>"
    s = s & "</table>"
    s = s & "<div class=""sh-cnt-note"">（ ）内は職員が受けた件数</div>"
    SheetCounts = s
End Function

Function SheetNum(n, st)
    SheetNum = "<span class=""big"">" & n & "</span><span class=""sub"">(" & st & ")</span>"
End Function

Function SheetBreakdown(brk)
    Dim i, ex, rf
    ex = 0
    rf = 0
    For i = 0 To UBound(brk)
        If ToStr(brk(i)("内訳区分")) = "交換" Then ex = ex + ToLong(brk(i)("件数計"))
        If ToStr(brk(i)("内訳区分")) = "返金" Then rf = rf + ToLong(brk(i)("件数計"))
    Next
    SheetBreakdown = "<div class=""sh-uchi"">内　交換　<span data-k=""交換"">" & ex & "</span> 件　／　内　返金　<span data-k=""返金"">" & rf & "</span> 件</div>"
End Function

Function SheetDailyText(daily, col)
    If UBound(daily) >= 0 Then
        SheetDailyText = ToStr(daily(0)(col))
    Else
        SheetDailyText = ""
    End If
End Function

Function SheetBox(label, text)
    Dim body
    body = HBr(text)
    If body = "" Then body = "&nbsp;"
    SheetBox = "<table class=""sh-box""><tr><th class=""b-h"">" & H(label) & "</th></tr><tr><td class=""b-txt"">" & body & "</td></tr></table>"
End Function

' 電話対応以外の業務：左 8 行 × 右 5 行の 2 段組。
' 2 段 × 8 行 ＝ 16 の枠に詰め、右の段の 6～8 行目（⑭～⑯ にあたる所）は空の枠にする。
Function SheetTasks(tasks)
    Dim no(15), nm(15), cnt(15), used, i, r, s
    used = 0
    For i = 0 To 15
        no(i) = ""
        nm(i) = ""
        cnt(i) = ""
    Next
    For i = 0 To UBound(tasks)
        If used < SHEET_TASK_USED Then
            If tasks(i)("有効") Or ToLong(tasks(i)("件数計")) > 0 Then
                no(used) = ToStr(tasks(i)("番号"))
                nm(used) = ToStr(tasks(i)("帳票表示名"))
                cnt(used) = ToLong(tasks(i)("件数計")) & " 件"
                used = used + 1
            End If
        End If
    Next
    s = "<table class=""sh-task"">"
    s = s & "<colgroup><col class=""w-t1""><col class=""w-t2""><col class=""w-t3""><col class=""w-t1""><col class=""w-t2""><col class=""w-t3""></colgroup>"
    s = s & "<tr><th colspan=""6"" class=""t-head"">電話対応以外の業務（データ入力作業等）</th></tr>"
    For r = 0 To SHEET_TASK_ROWS - 1
        s = s & "<tr>"
        s = s & "<td class=""t-no"">" & HCell(no(r)) & "</td><td class=""t-nm"">" & HCell(nm(r)) & "</td><td class=""t-n"">" & HCell(cnt(r)) & "</td>"
        s = s & "<td class=""t-no"">" & HCell(no(r + SHEET_TASK_ROWS)) & "</td><td class=""t-nm"">" & HCell(nm(r + SHEET_TASK_ROWS)) & "</td><td class=""t-n"">" & HCell(cnt(r + SHEET_TASK_ROWS)) & "</td>"
        s = s & "</tr>"
    Next
    s = s & "</table>"
    SheetTasks = s
End Function
%>
