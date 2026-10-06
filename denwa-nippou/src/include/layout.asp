<%
' ============================================================
'  画面の枠（上の青い帯・メニューへのリンク）と、どの画面でも使う小さな道具
'
'  日付は必ずここの YMD / DateJa / Wareki で文字にする。
'  VBScript で日付を & でつなぐと、サーバーの地域設定によって書式が変わるため。
' ============================================================

Sub PageHead(title)
    ' 古い画面が出ないように（IE モードは、同じアドレスの画面を前に見たまま出し直すことがある）
    Response.Expires = -1
    Response.AddHeader "Cache-Control", "no-cache, no-store"
    Response.AddHeader "Pragma", "no-cache"
    Response.Write "<!DOCTYPE html>" & vbCrLf & "<html lang=""ja""><head><meta charset=""utf-8"">"
    ' 課内のサイトは IE モードの「互換表示」で開かれることがあるので、最新の表示方法を指定する
    Response.Write "<meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">"
    Response.Write "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    Response.Write "<title>" & H(title) & " - " & H(APP_NAME) & "</title>"
    Response.Write "<link rel=""stylesheet"" href=""css/style.css""></head><body>"
    Response.Write "<div class=""band noprint""><div class=""band-in"">"
    Response.Write "<span class=""app"">" & H(APP_NAME) & "</span>"
    If IsStaff() Then
        Response.Write "<span class=""role role-staff"">職員用</span>"
    Else
        Response.Write "<span class=""role"">入力用</span>"
    End If
    Response.Write "<a class=""home"" href=""default.asp"">メニューへ</a>"
    Response.Write "</div></div>"
    Response.Write "<div class=""wrap""><h1 class=""noprint"">" & H(title) & "</h1>"
End Sub

Sub PageFoot()
    Response.Write "</div></body></html>"
    DbClose
End Sub

' ---------- 文字 ----------
' HTML に埋め込む文字を安全にする。Null と Empty は空にする。
' Server.HTMLEncode を使わないのは、' を置き換えないことと、Null を渡したときの動きを確かめていないため。
Function H(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        H = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

' 改行を <br> にして HTML に埋め込む（記述欄）
Function HBr(v)
    Dim s
    s = H(v)
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    HBr = Replace(s, vbLf, "<br>")
End Function

' 空なら &nbsp; にする（表の枠が途切れないように、空のセルにも必ず中身を入れる）
Function HCell(v)
    Dim s
    s = H(v)
    If s = "" Then s = "&nbsp;"
    HCell = s
End Function

Function U(s)
    U = Server.URLEncode(s)
End Function

' 前後の空白（半角・全角・タブ）を取る。VBScript の Trim は半角空白しか取らない。
Function TrimAll(s)
    Dim t, c
    t = s
    Do While Len(t) > 0
        c = Left(t, 1)
        If c = " " Or c = "　" Or c = vbTab Or c = vbCr Or c = vbLf Then
            t = Mid(t, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(t) > 0
        c = Right(t, 1)
        If c = " " Or c = "　" Or c = vbTab Or c = vbCr Or c = vbLf Then
            t = Left(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimAll = t
End Function

' 全角の数字と記号を半角にする（日本語入力のまま数字を打っても受け付けるため）。
' AscW は &H8000 以上の文字で負の数を返すので、文字コードの計算はせず対応表で置き換える。
Function ToHalfWidth(s)
    Dim i, c, p, full, half, t
    full = "０１２３４５６７８９－ー―／．：　"
    half = "0123456789---/.: "
    t = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        p = InStr(full, c)
        If p > 0 Then c = Mid(half, p, 1)
        t = t & c
    Next
    ToHalfWidth = t
End Function

Function IsDigits(s)
    Dim i
    IsDigits = (Len(s) > 0)
    For i = 1 To Len(s)
        If InStr("0123456789", Mid(s, i, 1)) = 0 Then IsDigits = False
    Next
End Function

' 件数の欄を読む。空は 0。数字以外や 5 けた以上なら False を返す。
Function ParseCount(s, ByRef n)
    Dim t
    t = ToHalfWidth(TrimAll(s))
    ParseCount = False
    n = 0
    If t = "" Then
        ParseCount = True
    ElseIf IsDigits(t) And Len(t) <= 4 Then
        n = CLng(t)
        ParseCount = True
    End If
End Function

' データベースから来た数（Null・Empty・小数の型で来ることがある）を Long にする
Function ToLong(v)
    If IsNull(v) Or IsEmpty(v) Then
        ToLong = 0
    Else
        ToLong = CLng(v)
    End If
End Function

' データベースから来た文字（Null がある）を文字にする
Function ToStr(v)
    If IsNull(v) Or IsEmpty(v) Then
        ToStr = ""
    Else
        ToStr = CStr(v)
    End If
End Function

' ---------- 送られてきた値 ----------
Function IsPost()
    IsPost = (Request.ServerVariables("REQUEST_METHOD") = "POST")
End Function

Function FormVal(name)
    Dim v
    v = Request.Form(name)
    If IsEmpty(v) Then
        FormVal = ""
    Else
        FormVal = TrimAll(CStr(v))
    End If
End Function

' 記述欄（改行を残す。前後の空白だけ取る）
Function FormText(name)
    Dim v
    v = Request.Form(name)
    If IsEmpty(v) Then
        FormText = ""
    Else
        FormText = TrimAll(Replace(CStr(v), vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf))
    End If
End Function

Function QueryVal(name)
    Dim v
    v = Request.QueryString(name)
    If IsEmpty(v) Then
        QueryVal = ""
    Else
        QueryVal = TrimAll(CStr(v))
    End If
End Function

' 数字だけの値（ID など）を読む。読めなければ dflt。
Function ParamLong(s, dflt)
    Dim t
    t = ToHalfWidth(s)
    If IsDigits(t) And Len(t) <= 9 Then
        ParamLong = CLng(t)
    Else
        ParamLong = dflt
    End If
End Function

' ---------- 日付 ----------
' 「2026-08-25」「2026/8/25」「20260825」を日付にする。読めなければ False。
' CDate を使わないのは、文字の日付の読み方がサーバーの地域設定で変わるため。
Function ParseYMD(s, ByRef d)
    Dim t, parts, y, m, dd
    ParseYMD = False
    t = Replace(Replace(ToHalfWidth(TrimAll(s)), "/", "-"), ".", "-")
    If Len(t) = 8 And IsDigits(t) Then t = Left(t, 4) & "-" & Mid(t, 5, 2) & "-" & Right(t, 2)
    parts = Split(t, "-")
    If UBound(parts) <> 2 Then Exit Function
    If Not IsDigits(parts(0)) Then Exit Function
    If Not IsDigits(parts(1)) Then Exit Function
    If Not IsDigits(parts(2)) Then Exit Function
    If Len(parts(0)) <> 4 Or Len(parts(1)) > 2 Or Len(parts(2)) > 2 Then Exit Function
    y = CLng(parts(0))
    m = CLng(parts(1))
    dd = CLng(parts(2))
    If y < 2000 Or y > 2099 Or m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    d = DateSerial(y, m, dd)
    If Month(d) <> m Then Exit Function
    ParseYMD = True
End Function

' 画面で選ばれた日付（?d=...）。無ければ今日。
Function PageDate()
    Dim d
    If ParseYMD(QueryVal("d"), d) Then
        PageDate = d
    Else
        PageDate = Date()
    End If
End Function

Function Pad2(n)
    Pad2 = Right("0" & n, 2)
End Function

Function YMD(d)
    YMD = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

Function YMDHM(d)
    YMDHM = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & " " & Pad2(Hour(d)) & ":" & Pad2(Minute(d))
End Function

Function WeekdayJa(d)
    WeekdayJa = Mid("日月火水木金土", Weekday(d), 1)
End Function

' 2026年8月25日（火）
Function DateJa(d)
    DateJa = Year(d) & "年" & Month(d) & "月" & Day(d) & "日（" & WeekdayJa(d) & "）"
End Function

' 8/25（火）
Function DateShort(d)
    DateShort = Month(d) & "/" & Day(d) & "（" & WeekdayJa(d) & "）"
End Function

' 帳票の日付。令和 8年 8月25日（火）。1 けたの数の前は半角の空白でそろえる（現行の紙と同じ）。
Function Wareki(d)
    Dim era, y, ys
    If d >= DateSerial(2019, 5, 1) Then
        era = "令和"
        y = Year(d) - 2018
    Else
        era = "平成"
        y = Year(d) - 1988
    End If
    If y = 1 Then
        ys = "元"
    Else
        ys = Right(" " & y, 2)
    End If
    Wareki = era & ys & "年" & Right(" " & Month(d), 2) & "月" & Right(" " & Day(d), 2) & "日（" & WeekdayJa(d) & "）"
End Function

' その週の月曜日
Function WeekStart(d)
    WeekStart = DateAdd("d", 1 - Weekday(d, vbMonday), d)
End Function

' ---------- 画面の部品 ----------
Sub ShowMsg(kind, text)
    Response.Write "<p class=""msg msg-" & kind & """>" & H(text) & "</p>"
End Sub

Function SelAttr(a, b)
    If CStr(a) = CStr(b) Then
        SelAttr = " selected"
    Else
        SelAttr = ""
    End If
End Function

Function ChkAttr(b)
    If b Then
        ChkAttr = " checked"
    Else
        ChkAttr = ""
    End If
End Function

' 画面を別の URL に移す（保存のあとに使う。ページはここで止まる）。
' On Error Resume Next が生きている所で呼ばないこと（止まらなくなる。手順書 7-2）。
Sub Go(url)
    DbClose
    Response.Redirect url
End Sub

' 日付を選ぶ欄（前の日・次の日つき）。押すと POST で送り、受け取った側が ?d= に移す。
Sub DateBar(page, d, extraName, extraVal)
    Dim ex
    ex = ""
    If extraName <> "" Then ex = "&amp;" & extraName & "=" & H(extraVal)
    Response.Write "<form method=""post"" action=""" & page & """ class=""bar noprint"">"
    Response.Write "<input type=""hidden"" name=""act"" value=""go"">"
    If extraName <> "" Then Response.Write "<input type=""hidden"" name=""" & extraName & """ value=""" & H(extraVal) & """>"
    Response.Write "<a class=""btn btn-s"" href=""" & page & "?d=" & YMD(DateAdd("d", -1, d)) & ex & """>&lt; 前の日</a> "
    Response.Write "<label>日付 <input type=""date"" name=""d"" value=""" & YMD(d) & """ class=""in-date""></label> "
    Response.Write "<button type=""submit"" class=""btn btn-s"">表示</button> "
    Response.Write "<a class=""btn btn-s"" href=""" & page & "?d=" & YMD(DateAdd("d", 1, d)) & ex & """>次の日 &gt;</a> "
    Response.Write "<a class=""btn btn-s"" href=""" & page & "?d=" & YMD(Date()) & ex & """>今日</a>"
    Response.Write "<span class=""bar-date"">" & DateJa(d) & "</span>"
    Response.Write "</form>"
End Sub

' 記述欄が帳票（A4 1 枚）の何行ぶんになるか。全角 1 文字を 2、半角を 1 と数える。
Function TextLines(s, perLine)
    Dim lines, i, w, j, c, n
    n = 0
    If s = "" Then
        TextLines = 0
        Exit Function
    End If
    lines = Split(Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        w = 0
        For j = 1 To Len(lines(i))
            c = Mid(lines(i), j, 1)
            If AscW(c) >= 32 And AscW(c) < 127 Then
                w = w + 1
            Else
                w = w + 2
            End If
        Next
        If w = 0 Then
            n = n + 1
        Else
            n = n + ((w + perLine * 2 - 1) \ (perLine * 2))
        End If
    Next
    TextLines = n
End Function
%>
