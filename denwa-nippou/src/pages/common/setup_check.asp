<%@ LANGUAGE="VBScript" CODEPAGE="65001" %><%
Option Explicit
Response.CharSet = "utf-8" : Response.CodePage = 65001
' ============================================================
'  設置チェック
'
'  ・#include を 1 つも使わない。include が壊れていても、このページだけは開けるように。
'    （そのため config.asp は「文字として」読み、db.asp と同じ場所の解き方をここにも書いている。
'      2 つが食い違わないことは tools/test_pages.py で確かめている）
'  ・上から順に調べて、表にする。いちばん上の赤い行が原因であることが多い。
'  ・最後に、切り分け用の probe1～7 と各画面を、このサーバー自身から HTTP で開いて結果を並べる。
' ============================================================
Server.ScriptTimeout = 300

Dim res(), nres, cfg, cfgOk, folder, roleVal, roleSeen, baseUrl, dbRaw, dbPath, dbOk, conn, i, n, msg, v
Dim fso, sh, net, okCount, ngCount, warnCount
ReDim res(-1)
nres = -1

Sub Add(sec, item, st, text)
    nres = nres + 1
    ReDim Preserve res(nres)
    res(nres) = Array(sec, item, st, text)
End Sub

Function HE(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        HE = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HE = s
End Function

' ---------- 道具（失敗しても先へ進めるよう、どれも小さく守る） ----------
Function TryCreate(progId, ByRef obj)
    Dim errNo
    On Error Resume Next
    Set obj = Server.CreateObject(progId)
    errNo = Err.Number
    On Error GoTo 0
    TryCreate = (errNo = 0)
End Function

Function ReadUtf8(path, ByRef ok)
    Dim st, t, errNo
    t = ""
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    t = st.ReadText
    errNo = Err.Number
    st.Close
    On Error GoTo 0
    ok = (errNo = 0)
    ReadUtf8 = t
End Function

Function ReadBytes(path, ByRef ok)
    Dim st, b, errNo
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile path
    b = st.Read(32)
    errNo = Err.Number
    st.Close
    On Error GoTo 0
    ok = (errNo = 0)
    ReadBytes = b
End Function

Function GetEnv(name)
    Dim e, errNo
    GetEnv = ""
    On Error Resume Next
    e = sh.Environment("PROCESS")(name)
    errNo = Err.Number
    On Error GoTo 0
    If errNo = 0 Then GetEnv = CStr(e)
End Function

Function NetValue(what)
    Dim v2, errNo
    NetValue = "（調べられませんでした）"
    On Error Resume Next
    If what = "user" Then v2 = net.UserName Else v2 = net.ComputerName
    errNo = Err.Number
    On Error GoTo 0
    If errNo = 0 Then NetValue = CStr(v2)
End Function

' config.asp の Const を読む（名前は小文字で持つ）。書き方が壊れている行は cfgBad に入れる。
Function ParseConfig(text, ByRef bad)
    Dim d, lines, j, ln, p, name, val, q
    Set d = Server.CreateObject("Scripting.Dictionary")
    bad = ""
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For j = 0 To UBound(lines)
        ln = Trim(Replace(lines(j), vbTab, " "))
        If LCase(Left(ln, 6)) = "const " Then
            p = InStr(ln, "=")
            If p = 0 Then
                bad = bad & (j + 1) & " 行目（= がありません）" & vbCrLf
            Else
                name = LCase(Trim(Mid(ln, 7, p - 7)))
                val = Trim(Mid(ln, p + 1))
                If Left(val, 1) = """" Then
                    q = InStr(2, val, """")
                    If q = 0 Then
                        bad = bad & (j + 1) & " 行目（"" が閉じていません）" & vbCrLf
                    Else
                        d(name) = Mid(val, 2, q - 2)
                    End If
                Else
                    If InStr(val, "'") > 0 Then val = Trim(Left(val, InStr(val, "'") - 1))
                    d(name) = val
                End If
            End If
        End If
    Next
    Set ParseConfig = d
End Function

Function CfgVal(name)
    If cfg.Exists(name) Then
        CfgVal = cfg(name)
    Else
        CfgVal = ""
    End If
End Function

' db.asp の DbResolvePath と同じ解き方（失敗したら "" と msg）
Function ResolveDb(p0, ByRef why)
    Dim p, root, u, j, mapped, errNo, errDesc
    why = ""
    ResolveDb = ""
    p = Trim(p0)
    If p = "" Then
        why = "DB_PATH が空です"
        Exit Function
    End If
    If Left(p, 2) = "\\" Or Mid(p, 2, 1) = ":" Then
        ResolveDb = p
        Exit Function
    End If
    p = Replace(p, "\", "/")
    If Left(p, 1) <> "/" Then
        u = Request.ServerVariables("SCRIPT_NAME")
        j = InStrRev(u, "/")
        If j > 0 Then u = Left(u, j - 1)
        j = InStrRev(u, "/")
        If j > 0 Then root = Left(u, j - 1) Else root = ""
        If InStr(p, "/") = 0 Then p = root & "/data/" & p Else p = root & "/" & p
    End If
    On Error Resume Next
    mapped = Server.MapPath(p)
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo <> 0 Then
        why = "Server.MapPath で場所に直せませんでした（" & errDesc & "）"
    Else
        ResolveDb = mapped
    End If
End Function

Function TryWrite(dir, ByRef why)
    Dim p, tsf, errNo, errDesc
    p = dir & "\_setup_check_" & CStr(Int(Timer() * 100)) & ".tmp"
    On Error Resume Next
    Set tsf = fso.CreateTextFile(p, True)
    tsf.Write "test"
    tsf.Close
    fso.DeleteFile p
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    why = errDesc
    TryWrite = (errNo = 0)
End Function

Function TryConnect(path, ByRef c, ByRef prov, ByRef why)
    Dim errNo, errDesc
    prov = "Microsoft.ACE.OLEDB.12.0"
    On Error Resume Next
    Set c = Server.CreateObject("ADODB.Connection")
    c.Open "Provider=" & prov & ";Data Source=" & path & ";"
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    If errNo = -2146824582 Then
        prov = "Microsoft.ACE.OLEDB.16.0"
        On Error Resume Next
        c.Open "Provider=" & prov & ";Data Source=" & path & ";"
        errNo = Err.Number
        errDesc = Err.Description
        On Error GoTo 0
    End If
    why = "番号 " & Hex(errNo) & "：" & errDesc
    If errNo = -2146824582 Then why = "ACE（Microsoft Access データベース エンジン）が見つかりません。入っていないか、32 ビット版と 64 ビット版が合っていません（" & why & "）"
    TryConnect = (errNo = 0)
End Function

Function TryCount(c, ByRef cnt, ByRef why)
    Dim rs, errNo, errDesc
    cnt = -1
    On Error Resume Next
    Set rs = c.Execute("SELECT Count(T.[担当者ID]) AS [件数] FROM [M_担当者] AS T")
    cnt = CLng(rs.Fields(0).Value)
    rs.Close
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    why = errDesc
    TryCount = (errNo = 0)
End Function

Function TryColumn(c, tbl, col)
    Dim rs, errNo, found
    found = False
    On Error Resume Next
    Set rs = c.OpenSchema(4, Array(Empty, Empty, tbl, col))
    found = Not rs.EOF
    rs.Close
    errNo = Err.Number
    On Error GoTo 0
    TryColumn = (errNo = 0 And found)
End Function

' 自分自身のサーバーに HTTP で取りに行く
Function TryHttp(url, ms, ByRef st, ByRef ctype, ByRef body, ByRef bin, ByRef why)
    Dim x, errNo, errDesc
    st = 0
    ctype = ""
    body = ""
    On Error Resume Next
    Set x = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    x.setTimeouts 5000, 5000, ms, ms
    x.setProxy 1
    x.open "GET", url, False
    x.send
    st = x.status
    ctype = x.getResponseHeader("Content-Type")
    body = x.responseText
    bin = x.responseBody
    errNo = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    why = errDesc
    TryHttp = (errNo = 0)
End Function

' XML の注釈（<!-- ... -->）を取り除く。web.config の説明文の中の <asp> を「設定」と見誤らないように。
Function NoXmlComments(t)
    Dim a, z, r
    r = t
    a = InStr(r, "<!--")
    Do While a > 0
        z = InStr(a, r, "-->")
        If z = 0 Then Exit Do
        r = Left(r, a - 1) & Mid(r, z + 3)
        a = InStr(r, "<!--")
    Loop
    NoXmlComments = r
End Function

Function StLabel(st)
    If st = "ok" Then
        StLabel = "よい"
    ElseIf st = "ng" Then
        StLabel = "だめ"
    ElseIf st = "warn" Then
        StLabel = "注意"
    Else
        StLabel = "－"
    End If
End Function

Function Has(s, part)
    Has = (InStr(1, s, part, vbTextCompare) > 0)
End Function

' 本文から文字だけを取り出して短くする（赤い行の説明に使う）
Function Gist(body)
    Dim t, j, c, inTag, out
    t = Replace(Replace(body, vbCr, " "), vbLf, " ")
    out = ""
    inTag = False
    For j = 1 To Len(t)
        c = Mid(t, j, 1)
        If c = "<" Then
            inTag = True
        ElseIf c = ">" Then
            inTag = False
            out = out & " "
        ElseIf Not inTag Then
            out = out & c
        End If
        If Len(out) > 600 Then Exit For
    Next
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    Gist = Trim(out)
End Function

' ---------- 1. サーバー ----------
If Not TryCreate("Scripting.FileSystemObject", fso) Then Add "1. サーバー", "ファイルの部品", "ng", "Scripting.FileSystemObject を作れません"
If Not TryCreate("WScript.Shell", sh) Then Add "1. サーバー", "WScript.Shell", "warn", "作れません（PDF を作れません）"
If Not TryCreate("WScript.Network", net) Then Add "1. サーバー", "WScript.Network", "warn", "作れません"
Add "1. サーバー", "サーバー名", "info", Request.ServerVariables("SERVER_NAME") & "（コンピューター名 " & NetValue("name") & "）"
Add "1. サーバー", "IIS", "info", Request.ServerVariables("SERVER_SOFTWARE")
Add "1. サーバー", "ASP", "ok", "動いています（VBScript " & ScriptEngineMajorVersion() & "." & ScriptEngineMinorVersion() & "）"
Add "1. サーバー", "アプリケーション プール", "info", Request.ServerVariables("APP_POOL_ID") & "（実行アカウント IIS AppPool\" & Request.ServerVariables("APP_POOL_ID") & "）"
Add "1. サーバー", "ファイルを読み書きする名前", "info", NetValue("user") & "（データベースのフォルダには、この名前と IIS_IUSRS に「変更」の許可が要ります）"
v = GetEnv("PROCESSOR_ARCHITECTURE")
If v = "x86" Then
    Add "1. サーバー", "32 / 64 ビット", "info", "32 ビットで動いています（ACE も 32 ビット版が必要）"
Else
    Add "1. サーバー", "32 / 64 ビット", "info", "64 ビットで動いています（ACE も 64 ビット版が必要）[" & v & "]"
End If

' ---------- 2. 設定と役割 ----------
Dim cfgText, cfgBad, url, parts
url = Request.ServerVariables("URL")
parts = Split(url, "/")
folder = ""
If UBound(parts) >= 1 Then folder = LCase(parts(UBound(parts) - 1))
cfgText = ReadUtf8(Server.MapPath("include/config.asp"), cfgOk)
If Not cfgOk Then
    Add "2. 設定", "config.asp", "ng", "include\config.asp が読めません（ありません）"
    Set cfg = Server.CreateObject("Scripting.Dictionary")
Else
    Set cfg = ParseConfig(cfgText, cfgBad)
    If cfgBad <> "" Then
        Add "2. 設定", "config.asp の書き方", "ng", "書き方が壊れている行があります（このままでは全部の画面が開けません）:" & vbCrLf & cfgBad
    Else
        Add "2. 設定", "config.asp の書き方", "ok", "読めました"
    End If
End If
roleVal = LCase(CfgVal("role"))
Add "2. 設定", "ROLE", "info", """" & CfgVal("role") & """"
If roleVal = "staff" Then
    roleSeen = "職員用"
ElseIf roleVal = "part" Then
    roleSeen = "パート職員用"
ElseIf roleVal = "" Then
    roleSeen = "ログオン名で判定（今のログオン名: """ & Request.ServerVariables("LOGON_USER") & """）"
Else
    roleSeen = "パート職員用（ROLE が part / staff 以外なので、狭い方に倒しています）"
End If
If (folder = "part" Or folder = "staff") And roleVal <> folder Then
    Add "2. 設定", "いまの役割", "ng", roleSeen & "。このフォルダは「" & folder & "」なのに ROLE が """ & CfgVal("role") & """ です。part と staff の config.asp を取り違えていませんか"
Else
    Add "2. 設定", "いまの役割", "ok", roleSeen
End If
Dim wc, wcOk
wc = NoXmlComments(ReadUtf8(Server.MapPath("web.config"), wcOk))
If Not wcOk Then
    Add "2. 設定", "web.config", "ng", "このフォルダに web.config がありません（エラーの説明画面が出なくなります）"
ElseIf Has(wc, "<asp ") Or Has(wc, "<asp>") Or Has(wc, "<asp/") Or Has(wc, "<authentication") Or Has(wc, "<ipSecurity") Then
    Add "2. 設定", "web.config", "ng", "IIS の既定で書けない設定（<asp> / <authentication> / <ipSecurity>）が入っています。サイト全体が 500.19 になります"
ElseIf Not Has(wc, "path=""/nippou/" & folder & "/error.asp""") Then
    Add "2. 設定", "web.config", "warn", "エラーの説明画面の場所が、このフォルダ（" & folder & "）を指していません。part と staff の web.config を取り違えていませんか"
Else
    Add "2. 設定", "web.config", "ok", "エラーの説明画面: /nippou/" & folder & "/error.asp"
End If

' ---------- 3. データベース ----------
Dim why, fb, fbOk, prov, cnt
dbRaw = CfgVal("db_path")
Add "3. データベース", "DB_PATH", "info", """" & dbRaw & """"
dbPath = ResolveDb(dbRaw, why)
dbOk = False
If dbPath = "" Then
    Add "3. データベース", "場所", "ng", why
Else
    Add "3. データベース", "場所", "info", dbPath
    If Not fso.FileExists(dbPath) Then
        Add "3. データベース", "ファイル", "ng", "ファイルがありません。置き場所か DB_PATH を確かめてください"
    Else
        Add "3. データベース", "ファイル", "ok", "あります（" & fso.GetFile(dbPath).Size & " バイト）"
        fb = ReadBytes(dbPath, fbOk)
        If fbOk Then
            If AscB(MidB(fb, 21, 1)) = 2 Then
                Add "3. データベース", "ファイルの形式", "ok", "Access 2007-2016 形式"
            ElseIf AscB(MidB(fb, 21, 1)) = 5 Then
                Add "3. データベース", "ファイルの形式", "ng", "Access 2016 専用形式です。課の Access で開けません（作り直しが必要です）"
            Else
                Add "3. データベース", "ファイルの形式", "warn", "想定外の形式です（" & AscB(MidB(fb, 21, 1)) & "）"
            End If
        End If
        If TryWrite(fso.GetParentFolderName(dbPath), why) Then
            Add "3. データベース", "フォルダへの書き込み", "ok", "書けます"
        Else
            Add "3. データベース", "フォルダへの書き込み", "ng", "書けません（" & why & "）。データベースのフォルダに、IUSR と IIS_IUSRS の「変更」の許可を付けてください（手順書 01）"
        End If
        v = GetEnv("TEMP")
        If v <> "" Then
            If fso.FolderExists(v) Then
                If TryWrite(v, why) Then
                    Add "3. データベース", "一時フォルダ（ACE の作業用）", "ok", v & "（書けます）"
                Else
                    Add "3. データベース", "一時フォルダ（ACE の作業用）", "warn", v & " に書けません（" & why & "）。データが増えたときに「未定義のエラー」で止まることがあります。このフォルダに " & NetValue("user") & " の「変更」の許可を付けてください（手順書 03）"
                End If
            End If
        End If
        If TryConnect(dbPath, conn, prov, why) Then
            Add "3. データベース", "ACE で接続", "ok", "接続できました（" & prov & "）"
            dbOk = True
            If TryCount(conn, cnt, why) Then
                If cnt = 0 Then
                    Add "3. データベース", "M_担当者", "warn", "0 人です。職員メニューの「マスタ保守」で担当者を登録してください"
                Else
                    Add "3. データベース", "M_担当者", "ok", cnt & " 人"
                End If
            Else
                Add "3. データベース", "M_担当者", "ng", "数えられません（" & why & "）"
            End If
            If TryColumn(conn, "M_担当者", "ログオン名") Then
                Add "3. データベース", "[ログオン名] 列", "ok", "あります"
            Else
                Add "3. データベース", "[ログオン名] 列", "ng", "ありません（古いデータベースを置いていませんか）"
            End If
            conn.Close
        Else
            Add "3. データベース", "ACE で接続", "ng", why
        End If
    End If
End If

' ---------- 4. PDF ----------
Dim edge, work
edge = CfgVal("pdf_edge_exe")
If fso.FileExists(edge) Then
    Add "4. PDF", "Microsoft Edge", "ok", edge
ElseIf fso.FileExists("C:\Program Files\Microsoft\Edge\Application\msedge.exe") Then
    Add "4. PDF", "Microsoft Edge", "ng", "PDF_EDGE_EXE の場所にありません。C:\Program Files\Microsoft\Edge\Application\msedge.exe にあります。config.asp の PDF_EDGE_EXE を書き換えてください"
Else
    Add "4. PDF", "Microsoft Edge", "ng", "見つかりません: " & edge
End If
work = CfgVal("pdf_work_dir")
If work = "" And dbPath <> "" Then
    If Left(dbPath, 2) = "\\" Then
        work = sh.ExpandEnvironmentStrings("%TEMP%") & "\nippou_pdf"
    Else
        work = Left(dbPath, InStrRev(dbPath, "\")) & "pdfwork"
    End If
End If
If work = "" Then
    Add "4. PDF", "作業フォルダ", "warn", "データベースの場所が分からないので決められません"
ElseIf fso.FolderExists(work) Then
    If TryWrite(work, why) Then
        Add "4. PDF", "作業フォルダ", "ok", work & "（書けます）"
    Else
        Add "4. PDF", "作業フォルダ", "ng", work & "（書けません: " & why & "）"
    End If
Else
    Add "4. PDF", "作業フォルダ", "info", work & "（まだありません。初めて PDF を作るときに作ります）"
End If

' ---------- 5. ファイル ----------
Dim want, staffOnly, f, missing, extra, bf, bfOk
want = Array("default.asp", "entry.asp", "tasks.asp", "error.asp", "setup_check.asp", "probe1.asp", "probe2.asp", "probe3.asp", _
             "probe4.asp", "probe5.asp", "probe6.asp", "probe7.asp", "web.config", "include\auth.asp", "include\config.asp", _
             "include\db.asp", "include\layout.asp", "include\sheet.asp", "include\grid.asp", "include\sql.asp", "include\pdf.asp", "css\style.css")
staffOnly = Array("staff.asp", "daily.asp", "report.asp", "detail.asp", "printpdf.asp", "summary.asp", "check.asp", "master.asp")
missing = ""
For Each f In want
    If Not fso.FileExists(Server.MapPath(Replace(f, "\", "/"))) Then missing = missing & f & " "
Next
extra = ""
For Each f In staffOnly
    If fso.FileExists(Server.MapPath(f)) Then
        If roleVal <> "staff" Then extra = extra & f & " "
    ElseIf roleVal = "staff" Then
        missing = missing & f & " "
    End If
Next
If missing = "" Then
    Add "5. ファイル", "そろっているか", "ok", "そろっています"
Else
    Add "5. ファイル", "そろっているか", "ng", "足りません: " & missing
End If
If extra <> "" Then Add "5. ファイル", "置いてはいけないファイル", "ng", "パート職員用のフォルダに職員用の画面があります: " & extra & "（消してください）"
For Each f In Array("include\auth.asp", "include\config.asp", "include\db.asp", "include\layout.asp", "include\sheet.asp", "include\grid.asp", "include\sql.asp", "include\pdf.asp")
    If fso.FileExists(Server.MapPath(Replace(f, "\", "/"))) Then
        bf = ReadBytes(Server.MapPath(Replace(f, "\", "/")), bfOk)
        If bfOk Then
            If AscB(MidB(bf, 1, 1)) = &HEF Then Add "5. ファイル", f, "ng", "先頭に BOM があります（メモ帳で「UTF-8（BOM 付き）」で保存されました）。「UTF-8」で保存し直してください"
        End If
    End If
Next

' ---------- 6. 画面を開いてみる ----------
Dim scheme, host, st, ctype, body, bin, pages, pg, num, label
If Request.ServerVariables("HTTPS") = "on" Then scheme = "https://" Else scheme = "http://"
host = Request.ServerVariables("HTTP_HOST")
baseUrl = scheme & host & Left(url, InStrRev(url, "/"))
For num = 1 To 7
    If TryHttp(baseUrl & "probe" & num & ".asp", 30000, st, ctype, body, bin, why) Then
        If st = 200 And Has(body, "PROBE" & num & "-OK") Then
            Add "6. 画面", "probe" & num & ".asp", "ok", "開けました"
        Else
            Add "6. 画面", "probe" & num & ".asp", "ng", "HTTP " & st & "：" & Gist(body)
        End If
    Else
        Add "6. 画面", "probe" & num & ".asp", "ng", "開けませんでした（" & why & "）"
    End If
Next
If roleVal = "staff" Then
    pages = Array("default.asp", "entry.asp", "tasks.asp", "staff.asp", "daily.asp", "report.asp", "summary.asp", "check.asp", "master.asp")
Else
    pages = Array("default.asp", "entry.asp", "tasks.asp")
End If
For Each pg In pages
    If TryHttp(baseUrl & pg, 30000, st, ctype, body, bin, why) Then
        If st = 200 And Not Has(body, "NIPPOU-ERROR") And Not Has(body, "NIPPOU-FATAL") And Has(body, "</html>") Then
            Add "6. 画面", pg, "ok", "開けました"
        Else
            Add "6. 画面", pg, "ng", "HTTP " & st & "：" & Gist(body)
        End If
    Else
        Add "6. 画面", pg, "ng", "開けませんでした（" & why & "）"
    End If
Next
If roleVal = "staff" And dbOk Then
    If TryHttp(baseUrl & "printpdf.asp?selftest=1", 120000, st, ctype, body, bin, why) Then
        If st = 200 And Has(ctype, "application/pdf") Then
            Add "6. 画面", "PDF の試し作成", "ok", "PDF ができました"
        Else
            Add "6. 画面", "PDF の試し作成", "ng", "HTTP " & st & "：" & Gist(body)
        End If
    Else
        Add "6. 画面", "PDF の試し作成", "ng", "開けませんでした（" & why & "）"
    End If
End If

okCount = 0
ngCount = 0
warnCount = 0
For i = 0 To nres
    If res(i)(2) = "ok" Then okCount = okCount + 1
    If res(i)(2) = "ng" Then ngCount = ngCount + 1
    If res(i)(2) = "warn" Then warnCount = warnCount + 1
Next
%><!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>設置チェック</title>
<style>
  body { font-family: "Meiryo UI", Meiryo, sans-serif; background: #eef2f7; margin: 0; }
  .hd { background: #1f4e8c; color: #fff; padding: 10px 16px; font-size: 20px; font-weight: bold; }
  .wrap { padding: 8px 16px 40px; }
  table.r { border-collapse: collapse; background: #fff; }
  table.r th, table.r td { border: 1px solid #9aa8ba; padding: 4px 8px; vertical-align: top; text-align: left; }
  table.r th { background: #e3eaf3; font-weight: normal; }
  td.st { font-weight: bold; white-space: nowrap; }
  tr.ok td.st { color: #1b7f3b; }
  tr.ng td { background: #fde8e8; }
  tr.ng td.st { color: #c62828; }
  tr.warn td.st { color: #b45309; }
  .pre { white-space: pre-wrap; word-break: break-all; }
  .sum { font-size: 18px; margin: 12px 0; }
</style></head><body>
<div class="hd">設置チェック（電話応対日報 集計システム）</div>
<div class="wrap">
<p class="sum">赤（だめ） <b><%= ngCount %></b> ／ 黄（注意） <b><%= warnCount %></b> ／ 緑（よい） <b><%= okCount %></b></p>
<% If ngCount > 0 Then %>
<p>上から見て、<b>最初の赤い行</b>が原因であることが多いです。手順書「03_困ったとき」で、その行の「項目」を探してください。</p>
<% Else %>
<p>赤い行はありません。</p>
<% End If %>
<table class="r">
<tr><th>区分</th><th>項目</th><th>結果</th><th>内容</th></tr>
<%
For i = 0 To nres
    Response.Write "<tr class=""" & res(i)(2) & """><td>" & HE(res(i)(0)) & "</td><td>" & HE(res(i)(1)) & "</td><td class=""st"">"
    Response.Write StLabel(res(i)(2)) & "</td><td class=""pre"">" & HE(res(i)(3)) & "</td></tr>"
Next
%>
</table>
<p>お問い合わせのときは、この画面を印刷するか、下の枠の中をコピーして送ってください。</p>
<textarea rows="12" cols="110" readonly><%
For i = 0 To nres
    Response.Write HE(res(i)(0) & " / " & res(i)(1) & " / " & StLabel(res(i)(2)) & " / " & res(i)(3)) & vbCrLf
Next
%></textarea>
</div></body></html>
