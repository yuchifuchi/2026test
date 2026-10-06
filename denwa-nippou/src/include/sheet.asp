<%
' ============================================================
'  帳票「電話応対報告書日報集計表」（A4 縦 1 枚）
'
'  様式は、課の「日報集計印刷用フォーム.xlsm」の印刷用シートに合わせてある。
'  Excel の 1 行（16 ピクセル）を 5.65mm として、罫線の位置を写した。
'
'  画面で見る（report.asp）のも、PDF にする（printpdf.asp）のも、必ずこの SheetHtml(dt) を呼ぶ。
'  片方だけ直して食い違う事故を、作りの上で起こさないため。
'
'  罫線の決まり:
'    ・6 つのまとまり（出勤者・問合せ件数・特記事項・職員に代わった案件・要望・電話対応以外の業務）は、
'      それぞれ外枠の閉じた表にする。外枠は表そのものの罫線で引く（欄の罫線に頼らない）。
'    ・空の欄も、必ず &nbsp; を入れた <td> を置く（枠や点線が途切れないように）。
'    ・罫線は css/style.css の .sheet の規則だけで引く。ここでは style 属性を使わない。
'  記述欄は、文字をここで行に分けて、点線 1 本に 1 行ずつ載せる（ブラウザーに折り返させると、行の数が変わって
'  A4 1 枚に収まらなくなるため）。入りきらない長さは、日報の画面で保存する前に止めている。
'
'  数はすべて sql.asp の SqlCallsBase（T_受電 ただ 1 つ）から計算したもの。
' ============================================================

Const SHEET_ATT_ROWS = 7          ' 出勤者の氏名欄の行数（2 列 × 7 行 ＝ 14 名まで。左の列から縦に埋める）
Const SHEET_TASK_ROWS = 8         ' 電話対応以外の業務：左の段の行数（右の段は 5 行だけ使う）
Const SHEET_TASK_USED = 13        ' 電話対応以外の業務：使う欄の数（①～⑬）
Const SHEET_STAMPS = 8            ' 回覧の押印欄の数（左 4・右 4）

' 帳票に出す文字の長さの上限（マスタ保守・日報の画面で止める）。文字を縮めて収めるのではなく、長さで収める。
' 値を変えたら tools/render_sheet.py で「上限いっぱいでも A4 1 枚・はみ出し無し」を確かめ直すこと。
Const SHEET_NAME_MAX = 9          ' 出勤者の氏名
Const SHEET_TASK_LABEL_MAX = 13   ' 電話対応以外の業務の名前
Const SHEET_COL_NAME_MAX = 5      ' 問合せ件数の列の名前
Const SHEET_STAMP_NAME_MAX = 6    ' 回覧の押印欄の名前

' 記述欄の罫線の数と、1 行に入る幅（全角 1 文字を 2、半角 1 文字を 1 と数える）
Const SHEET_MEMO1_ROWS = 4        ' 特記事項
Const SHEET_MEMO2_ROWS = 5        ' 職員に代わった案件（概要）
Const SHEET_MEMO3_ROWS = 6        ' 要望
Const SHEET_LINE_WIDE = 92        ' 幅いっぱいの行（全角 46 文字）
Const SHEET_LINE_NARROW = 68      ' 特記事項の 1 行目（右に「内 返金」があるので狭い。全角 34 文字）
Const SHEET_NOTE_ROWS = 8         ' 出勤者の「備考」の欄に入る行の数
Const SHEET_NOTE_WIDE = 36        ' 「備考」1 行の幅（全角 18 文字）

Function SheetHtml(dt)
    Dim d2, daily, att, cols, brk, tasks, stamps, s, ex, rf, i
    d2 = DateAdd("d", 1, dt)
    daily = DbQuery(SqlDaily(), Array(dt, d2))
    att = DbQuery(SqlAttendance(), Array(dt, d2))
    cols = DbQuery(SqlColumnTotals(), Array(dt, d2))
    brk = DbQuery(SqlBreakdownTotals(), Array(dt, d2))
    tasks = DbQuery(SqlTaskTotals(), Array(dt, d2))
    stamps = DbQuery(SqlStampsAll(), Array())
    ex = 0
    rf = 0
    For i = 0 To UBound(brk)
        If ToStr(brk(i)("内訳区分")) = "交換" Then ex = ex + ToLong(brk(i)("件数計"))
        If ToStr(brk(i)("内訳区分")) = "返金" Then rf = rf + ToLong(brk(i)("件数計"))
    Next
    s = "<div class=""sheet"">"
    s = s & "<div class=""s-top""><span class=""s-limit"">【課内限り】</span>" & SheetDate(dt) & "</div>"
    s = s & "<div class=""s-kairan"">回覧</div>"
    s = s & SheetStamps(stamps)
    s = s & "<div class=""s-title"">電話応対報告書日報集計表</div>"
    s = s & SheetAttendance(att, daily)
    s = s & SheetCounts(cols)
    s = s & SheetMemo1(SheetDailyText(daily, "特記事項"), ex, rf)
    s = s & SheetMemo("職員に代わった案件（概要）", SheetDailyText(daily, "職員代替案件"), SHEET_MEMO2_ROWS)
    s = s & SheetMemo("要望（お客様からの問い合わせを減らすための改善提案等）", SheetDailyText(daily, "要望"), SHEET_MEMO3_ROWS)
    s = s & SheetTasks(tasks)
    s = s & "</div>"
    SheetHtml = s
End Function

' 右上の日付。令和　8年　　8月　　25日（火）。数は右にそろえる（現行の紙と同じ）。
Function SheetDate(d)
    Dim era, y
    If d >= DateSerial(2019, 5, 1) Then
        era = "令和"
        y = Year(d) - 2018
    Else
        era = "平成"
        y = Year(d) - 1988
    End If
    SheetDate = "<span class=""s-date"">" & era & "<span class=""d2"">" & y & "</span>年<span class=""d2"">" & Month(d) & _
                "</span>月<span class=""d2"">" & Day(d) & "</span>日（" & WeekdayJa(d) & "）</span>"
End Function

' 回覧の押印欄（左 4・右 4）。名前の無い欄も枠は出す。
Function SheetStamps(stamps)
    Dim nm(7), i, s, side, k
    For i = 0 To 7
        nm(i) = ""
    Next
    For i = 0 To UBound(stamps)
        If i <= 7 Then nm(i) = ToStr(stamps(i)("表示名"))
    Next
    s = "<div class=""s-stamps"">"
    For side = 0 To 1
        If side = 0 Then
            s = s & "<table class=""s-stamp s-stamp-l""><tr>"
        Else
            s = s & "<table class=""s-stamp s-stamp-r""><tr>"
        End If
        For k = 0 To 3
            s = s & "<th>" & HCell(nm(side * 4 + k)) & "</th>"
        Next
        s = s & "</tr><tr>"
        For k = 0 To 3
            s = s & "<td>&nbsp;</td>"
        Next
        s = s & "</tr></table>"
    Next
    SheetStamps = s & "</div>"
End Function

' 出勤者：左に「出勤者／○ 名／（ ○ 回線 ）」、中に氏名 2 列 × 7 行、右に備考の大きな欄
Function SheetAttendance(att, daily)
    Dim nm(13), i, r, s, n, lines, notes, k, nlines, lineCount
    n = UBound(att) + 1
    For i = 0 To 13
        nm(i) = ""
    Next
    For i = 0 To UBound(att)
        If i <= 13 Then nm(i) = ToStr(att(i)("氏名"))
    Next
    lineCount = "　"
    If UBound(daily) >= 0 Then
        If Not IsNull(daily(0)("回線数")) Then lineCount = CStr(ToLong(daily(0)("回線数")))
    End If
    ' 備考：1 人ずつ「姓：勤務時間 備考」を行に分けて並べる
    notes = ""
    nlines = 0
    For i = 0 To UBound(att)
        lines = SheetWrap(SheetAttNote(att(i)), SHEET_NOTE_WIDE, SHEET_NOTE_WIDE)
        For k = 0 To UBound(lines)
            If notes <> "" Then notes = notes & "<br>"
            notes = notes & H(lines(k))
            nlines = nlines + 1
        Next
    Next
    If notes = "" Then notes = "&nbsp;"
    s = "<table class=""s-blk s-att""><colgroup><col class=""w-lbl""><col class=""w-nm""><col class=""w-nm""><col></colgroup>"
    s = s & "<tr><td class=""lbl"">出勤者</td><th class=""ln"">氏名</th><th class=""ln"">氏名</th><th class=""ln"">備考</th></tr>"
    For r = 0 To SHEET_ATT_ROWS - 1
        s = s & "<tr>"
        If r = 0 Then
            s = s & "<td class=""lbl""><span class=""num a-num"">" & n & "</span> 名</td>"
        ElseIf r = SHEET_ATT_ROWS - 1 Then
            s = s & "<td class=""lbl"">（<span class=""num"">" & H(lineCount) & "</span> 回線）</td>"
        Else
            s = s & "<td class=""lbl"">&nbsp;</td>"
        End If
        s = s & "<td class=""ln a-name"">" & HCell(nm(r)) & "</td><td class=""ln a-name"">" & HCell(nm(r + SHEET_ATT_ROWS)) & "</td>"
        If r = 0 Then s = s & "<td class=""a-note"" rowspan=""" & SHEET_ATT_ROWS & """>" & notes & "</td>"
        s = s & "</tr>"
    Next
    SheetAttendance = s & "</table>"
End Function

' 備考に出す 1 人ぶん（姓：勤務時間 備考）。何も無ければ ""。
Function SheetAttNote(row)
    Dim t
    t = TrimAll(ToStr(row("勤務時間")) & " " & ToStr(row("備考")))
    If t <> "" Then t = ToStr(row("姓")) & "：" & t
    SheetAttNote = t
End Function

' 問合せ件数：左に合計、右に 5 列。どの欄も (職員が受けた件数) を数字の左に小さく添える。
Function SheetCounts(cols)
    Dim i, tot, totSt, s, head, body
    tot = 0
    totSt = 0
    head = ""
    body = ""
    For i = 0 To UBound(cols)
        tot = tot + ToLong(cols(i)("件数計"))
        totSt = totSt + ToLong(cols(i)("職員計"))
        head = head & "<th class=""ln c-col"">" & HCell(cols(i)("集計列名")) & "</th>"
        body = body & "<td class=""ln c-col"" data-col=""" & H(cols(i)("集計列ID")) & """><span class=""st"">(" & ToLong(cols(i)("職員計")) & _
               ")</span><span class=""v"">" & ToLong(cols(i)("件数計")) & "</span></td>"
    Next
    s = "<table class=""s-blk s-cnt""><colgroup><col class=""w-lbl""><col class=""w-c""><col class=""w-c""><col class=""w-c""><col class=""w-c""><col class=""w-c""><col></colgroup>"
    s = s & "<tr><td class=""lbl"">問合せ件数</td>" & head & "<td class=""c-note"" rowspan=""2"">(　)内は<br>職員受電数</td></tr>"
    s = s & "<tr><td class=""lbl c-tot"" data-col=""total""><span class=""num"">" & tot & "</span> 件<span class=""st st-r"">(" & totSt & ")</span></td>" & body & "</tr>"
    SheetCounts = s & "</table>"
End Function

Function SheetDailyText(daily, col)
    If UBound(daily) >= 0 Then
        SheetDailyText = ToStr(daily(0)(col))
    Else
        SheetDailyText = ""
    End If
End Function

' 特記事項：右上に「内 交換／内 返金」。1 行目だけ右が使われているので狭い。
Function SheetMemo1(text, ex, rf)
    Dim lines, rows, r, s, ln
    lines = SheetWrap(text, SHEET_LINE_NARROW, SHEET_LINE_WIDE)
    rows = SHEET_MEMO1_ROWS
    If UBound(lines) + 1 > rows Then rows = UBound(lines) + 1
    s = "<table class=""s-blk s-memo""><colgroup><col class=""w-memo1""><col></colgroup>"
    s = s & "<tr class=""rl""><td class=""m-h"">特記事項（報告書の電話件数だけでは伝わり難い事項など）</td>"
    s = s & "<td class=""m-uchi"">内　交換<span class=""num u-num"" data-k=""交換"">" & ex & "</span> 件</td></tr>"
    For r = 0 To rows - 1
        ln = ""
        If r <= UBound(lines) Then ln = lines(r)
        If r = 0 Then
            s = s & "<tr class=""rl""><td class=""m-line"">" & HCell(ln) & "</td><td class=""m-uchi"">内　返金<span class=""num u-num"" data-k=""返金"">" & rf & "</span> 件</td></tr>"
        Else
            s = s & "<tr class=""rl""><td class=""m-line"" colspan=""2"">" & HCell(ln) & "</td></tr>"
        End If
    Next
    SheetMemo1 = s & "</table>"
End Function

' 記述欄（見出し ＋ 点線の行）
Function SheetMemo(title, text, minRows)
    Dim lines, rows, r, s, ln
    lines = SheetWrap(text, SHEET_LINE_WIDE, SHEET_LINE_WIDE)
    rows = minRows
    If UBound(lines) + 1 > rows Then rows = UBound(lines) + 1
    s = "<table class=""s-blk s-memo""><tr class=""rl""><td class=""m-h"">" & H(title) & "</td></tr>"
    For r = 0 To rows - 1
        ln = ""
        If r <= UBound(lines) Then ln = lines(r)
        s = s & "<tr class=""rl""><td class=""m-line"">" & HCell(ln) & "</td></tr>"
    Next
    SheetMemo = s & "</table>"
End Function

' 電話対応以外の業務：左 8 行 × 右 5 行の 2 段組。上と下に 1 行ずつ空きの行（現行の紙と同じ）。
' 2 段 × 8 行 ＝ 16 の枠に詰め、右の段の 6～8 行目は空の行にする。
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
                cnt(used) = CStr(ToLong(tasks(i)("件数計")))
                used = used + 1
            End If
        End If
    Next
    s = "<table class=""s-blk s-task""><colgroup><col class=""w-tno""><col class=""w-tnm1""><col class=""w-tn""><col class=""w-tgap""><col class=""w-tno""><col class=""w-tnm2""><col></colgroup>"
    s = s & "<tr class=""rl""><td class=""m-h"" colspan=""7"">電話対応以外の業務（データ入力作業等）</td></tr>"
    s = s & "<tr class=""rl""><td colspan=""7"">&nbsp;</td></tr>"
    For r = 0 To SHEET_TASK_ROWS - 1
        s = s & "<tr class=""rl"">" & SheetTaskCells(no(r), nm(r), cnt(r)) & "<td>&nbsp;</td>"
        s = s & SheetTaskCells(no(r + SHEET_TASK_ROWS), nm(r + SHEET_TASK_ROWS), cnt(r + SHEET_TASK_ROWS)) & "</tr>"
    Next
    s = s & "<tr class=""rl""><td colspan=""7"">&nbsp;</td></tr>"
    SheetTasks = s & "</table>"
End Function

Function SheetTaskCells(no, nm, cnt)
    If no = "" And nm = "" Then
        SheetTaskCells = "<td class=""t-no"">&nbsp;</td><td class=""t-nm"">&nbsp;</td><td class=""t-n"">&nbsp;</td>"
    Else
        SheetTaskCells = "<td class=""t-no"">" & HCell(no) & "</td><td class=""t-nm"">" & HCell(nm) & "</td><td class=""t-n""><span class=""num"">" & H(cnt) & "</span> 件</td>"
    End If
End Function

' ---------- 記述欄の行分け（日報の画面の「入りきるか」の判定にも使う） ----------
' 文字の幅：半角（英数字・半角カナ）は 1、それ以外は 2。
' AscW は &H8000 以上の文字で負の数を返すので、65536 を足して直す。範囲は 10 進で書く（&HFF61 は負の数になるため）。
Function SheetCharWidth(c)
    Dim u
    u = AscW(c)
    If u < 0 Then u = u + 65536
    If u < 128 Or (u >= 65377 And u <= 65439) Then
        SheetCharWidth = 1
    Else
        SheetCharWidth = 2
    End If
End Function

' 文字を、改行と幅で行に分ける。1 行目だけ幅 firstWidth、2 行目からは width。
Function SheetWrap(text, firstWidth, width)
    Dim out, n, paras, i, j, c, w, cw, cur, lim
    ReDim out(-1)
    n = -1
    If text = "" Then
        SheetWrap = out
        Exit Function
    End If
    paras = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(paras)
        cur = ""
        w = 0
        If n = -1 Then lim = firstWidth Else lim = width
        For j = 1 To Len(paras(i))
            c = Mid(paras(i), j, 1)
            cw = SheetCharWidth(c)
            If w + cw > lim Then
                n = n + 1
                ReDim Preserve out(n)
                out(n) = cur
                cur = ""
                w = 0
                lim = width
            End If
            cur = cur & c
            w = w + cw
        Next
        n = n + 1
        ReDim Preserve out(n)
        out(n) = cur
    Next
    SheetWrap = out
End Function

' 日報の画面から呼ぶ：記述欄と備考が帳票の罫線に入りきるか。入りきらなければ理由、入れば ""。
Function SheetFitProblem(t1, t2, t3, attNotes)
    Dim msg, n
    msg = ""
    n = UBound(SheetWrap(t1, SHEET_LINE_NARROW, SHEET_LINE_WIDE)) + 1
    If n > SHEET_MEMO1_ROWS Then msg = msg & "特記事項が " & n & " 行になります（帳票の罫線は " & SHEET_MEMO1_ROWS & " 行。1 行目は全角 " & (SHEET_LINE_NARROW \ 2) & " 文字、2 行目からは全角 " & (SHEET_LINE_WIDE \ 2) & " 文字まで）。"
    n = UBound(SheetWrap(t2, SHEET_LINE_WIDE, SHEET_LINE_WIDE)) + 1
    If n > SHEET_MEMO2_ROWS Then msg = msg & "職員に代わった案件が " & n & " 行になります（帳票の罫線は " & SHEET_MEMO2_ROWS & " 行。1 行は全角 " & (SHEET_LINE_WIDE \ 2) & " 文字まで）。"
    n = UBound(SheetWrap(t3, SHEET_LINE_WIDE, SHEET_LINE_WIDE)) + 1
    If n > SHEET_MEMO3_ROWS Then msg = msg & "要望が " & n & " 行になります（帳票の罫線は " & SHEET_MEMO3_ROWS & " 行。1 行は全角 " & (SHEET_LINE_WIDE \ 2) & " 文字まで）。"
    If attNotes > SHEET_NOTE_ROWS Then msg = msg & "出勤者の勤務時間・備考が、帳票の「備考」の欄で " & attNotes & " 行になります（" & SHEET_NOTE_ROWS & " 行まで。1 行は全角 " & (SHEET_NOTE_WIDE \ 2) & " 文字）。"
    SheetFitProblem = msg
End Function

' 日報の画面から呼ぶ：1 人ぶんの「備考」が帳票で何行になるか
Function SheetNoteLines(sei, hours, note)
    Dim t
    t = TrimAll(hours & " " & note)
    If t = "" Then
        SheetNoteLines = 0
    Else
        SheetNoteLines = UBound(SheetWrap(sei & "：" & t, SHEET_NOTE_WIDE, SHEET_NOTE_WIDE)) + 1
    End If
End Function
%>
