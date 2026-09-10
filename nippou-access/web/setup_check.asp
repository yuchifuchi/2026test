<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<% Response.CharSet = "utf-8" : Response.CodePage = 65001 %>
<%
' -----------------------------------------------------------------------------
'  設置チェック
'
'  「開いても真っ白」「500 エラー」で止まったときに、どこでつまずいているかを
'  1 画面で出す。ほかのファイルを読み込まない (#include しない) ので、
'  db.asp や sql.asp が壊れていてもこのページだけは開く。
'
'      http://<サーバー名>/nippou/setup_check.asp
'
'  設置が終わったら消してかまいません。
' -----------------------------------------------------------------------------
Dim gOK, gNG

Function E(v)
    Dim s
    s = "" & v
    s = Replace(s, "&", "&amp;") : s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;") : s = Replace(s, """", "&quot;")
    E = s
End Function

' テキストを UTF-8 として読む。FileSystemObject は UTF-8 を読めないため。
Function ReadUtf8(path)
    Dim st
    ReadUtf8 = ""
    On Error Resume Next
    Set st = Server.CreateObject("ADODB.Stream")
    st.Type = 2 : st.Charset = "utf-8" : st.Open
    st.LoadFromFile path
    If Err.Number = 0 Then ReadUtf8 = st.ReadText
    st.Close
    Err.Clear
End Function

' 1 行ぶんの結果を出す。ok = True/False/Null(参考情報)
Sub Row(name, ok, value, hint)
    Dim mark, cls
    If IsNull(ok) Then
        mark = "―" : cls = ""
    ElseIf ok Then
        mark = "OK" : cls = "ok" : gOK = gOK + 1
    Else
        mark = "NG" : cls = "ng" : gNG = gNG + 1
    End If
    Response.Write "<tr class=""" & cls & """><td class=""m"">" & mark & "</td><td>" & _
        E(name) & "</td><td>" & E(value) & "</td><td>" & hint & "</td></tr>" & vbCrLf
End Sub

Dim fso, sh, conn, rs, f, txt, i, p, hasFso
Dim appPath, dbPath, edgePath, tmpDir, probe, n, hasLogon
gOK = 0 : gNG = 0
appPath = Server.MapPath(".")
%>
<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>設置チェック | 電話応対日報 集計システム</title>
<style>
body{font-family:"Meiryo","Yu Gothic UI",system-ui,sans-serif; margin:0; background:#f4f6f8; color:#1b2733}
main{max-width:1000px; margin:0 auto; padding:24px 20px 60px}
h1{font-size:22px; margin:0 0 6px}
p.lead{color:#5b6b7c; margin:0 0 18px}
table{border-collapse:collapse; width:100%; background:#fff; font-size:14px}
th,td{border:1px solid #d3dbe3; padding:8px 10px; text-align:left; vertical-align:top}
th{background:#eef2f6; font-size:12px; letter-spacing:.06em}
td.m{width:3.4em; text-align:center; font-weight:700}
tr.ok td.m{color:#1c7c4a} tr.ng td.m{color:#a32020}
tr.ng td{background:#fbeaea}
code{background:#eef2f6; padding:1px 5px; border-radius:4px; font-size:13px}
.sum{margin:18px 0; padding:12px 14px; border-radius:8px; background:#fff; border:1px solid #d3dbe3}
.sum b{font-size:17px}
h2{font-size:16px; margin:26px 0 8px}
</style>
</head><body><main>
<h1>設置チェック</h1>
<p class="lead">
  この画面が出ている時点で、IIS と ASP は動いています。
  下の表で <b>NG</b> の行があれば、その右の「直しかた」を見てください。
</p>

<h2>1. サーバーと ASP</h2>
<table>
<tr><th>結果</th><th>見たもの</th><th>いまの値</th><th>直しかた</th></tr>
<%
Row "ASP が動いている", True, "" & Request.ServerVariables("SERVER_SOFTWARE"), _
    "このページが出ていれば OK です。"
Row "日本語が化けていない", Null, "あいうえお 漢字 ①②③", _
    "ここが化けていたら、.asp を Shift_JIS で保存し直してしまっています。"
Row "置き場所", Null, appPath, "web/ の中身をここに置いています。"

' ファイルを触るときのアカウント。フォルダの権限をどれに与えるかの手がかり。
Dim acct
acct = ""
On Error Resume Next
Set sh = Server.CreateObject("WScript.Shell")
If Err.Number = 0 Then acct = sh.ExpandEnvironmentStrings("%USERDOMAIN%\%USERNAME%")
Err.Clear
On Error GoTo 0
Row "サーバー側で動いているアカウント", Null, _
    IIf2(Len(acct) > 0, acct, "(取得できませんでした)"), _
    "アプリケーション プールの ID です。共有フォルダの権限は、" & _
    "このアカウント（匿名認証で別のアカウントを指定しているならそちら）に与えます。"
Row "URL", Null, "" & Request.ServerVariables("URL"), ""
%>
</table>

<h2>2. 認証（誰が使っているか）</h2>
<table>
<tr><th>結果</th><th>見たもの</th><th>いまの値</th><th>直しかた</th></tr>
<%
Dim who
who = Trim("" & Request.ServerVariables("LOGON_USER"))
If Len(who) = 0 Then who = Trim("" & Request.ServerVariables("AUTH_USER"))
Row "Windows 認証", (Len(who) > 0), _
    IIf2(Len(who) > 0, who, "(空。匿名で開いています)"), _
    "<b>空のままでも動きます。</b>ただし次の 2 つができません。<br>" & _
    "・誰が入力したかの記録（T_受電 の登録者が「(未認証)」になります）<br>" & _
    "・パート職員と職員で画面を分けること（<b>全員が職員として全画面を開けます</b>）<br>" & _
    "使い分けたいときは IIS マネージャー →「認証」で" & _
    " <b>Windows 認証 = 有効</b>／<b>匿名認証 = 無効</b> にしてください。" & _
    "あとから切り替えても、入力済みのデータには影響しません。"
Row "認証の種類", Null, "" & Request.ServerVariables("AUTH_TYPE"), ""
%>
</table>

<h2>3. データベース</h2>
<table>
<tr><th>結果</th><th>見たもの</th><th>いまの値</th><th>直しかた</th></tr>
<%
' db.asp は読み込まない (壊れていてもこのページを出したいため)。
' テキストとして開いて DB_PATH の行だけを取り出す。
dbPath = ""
hasFso = False
On Error Resume Next
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Err.Number <> 0 Then
    Row "ファイルを扱う部品 (FileSystemObject)", False, Err.Description, _
        "サーバーの設定でスクリプトが制限されています。情報システム担当にご相談ください。"
    Err.Clear
Else
    hasFso = True
    txt = ""
    If Not fso.FileExists(Server.MapPath("include/config.asp")) Then
        Row "include\config.asp がある", False, "ありません", _
            "include フォルダごと置けているか確認してください。" & _
            "（古い版をお使いの場合、設定は include\db.asp の中にあります）"
    Else
        txt = ReadUtf8(Server.MapPath("include/config.asp"))
        Row "include\config.asp がある", True, "見つかりました", ""
        i = InStr(txt, "Const DB_PATH")
        If i > 0 Then
            txt = Mid(txt, i)
            i = InStr(txt, """")
            If i > 0 Then
                txt = Mid(txt, i + 1)
                i = InStr(txt, """")
                If i > 0 Then dbPath = Left(txt, i - 1)
            End If
        End If
    End If
End If
On Error GoTo 0

Row "接続先の設定 (DB_PATH)", (Len(dbPath) > 0), _
    IIf2(Len(dbPath) > 0, dbPath, "(読み取れませんでした)"), _
    "<b>include\config.asp</b> の <code>Const DB_PATH</code> です。" & _
    "ここを、実際に .accdb を置いた場所に書き換えてください。"

If Len(dbPath) > 0 Then
    On Error Resume Next
    Row "その場所にファイルがある", fso.FileExists(dbPath), _
        IIf2(fso.FileExists(dbPath), "あります", "見つかりません"), _
        "共有フォルダに 日報集計_be.accdb を置き、その場所を DB_PATH に書いてください。" & _
        "ネットワーク越しの場合は <code>\\サーバー名\共有名\…</code> の形で書きます。"
    Err.Clear

    ' 置いたフォルダに書き込めるか (ACE はロックファイルを作るので読み取りだけでは動かない)
    p = fso.GetParentFolderName(dbPath)
    probe = ""
    If Len(p) > 0 Then
        probe = fso.BuildPath(p, "nippou_write_test.tmp")
        Err.Clear
        Set f = fso.CreateTextFile(probe, True)
        If Err.Number = 0 Then
            f.Close
            fso.DeleteFile probe, True
            Row "そのフォルダに書き込める", True, p, ""
        Else
            Row "そのフォルダに書き込める", False, Err.Description, _
                "<b>読み取りだけでは動きません。</b>そのフォルダに、" & _
                "アプリケーション プール ID (既定は <code>IIS AppPool\&lt;プール名&gt;</code>) の" & _
                "「変更」権限を与えてください。"
        End If
        Err.Clear
    End If
    On Error GoTo 0
End If

' ACE で実際につないでみる
On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
If Err.Number <> 0 Then
    Row "データベース機能 (ADODB)", False, Err.Description, "情報システム担当にご相談ください。"
    Err.Clear
Else
    Row "データベース機能 (ADODB)", True, "使えます", ""
    If Len(dbPath) > 0 Then
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
        If Err.Number <> 0 Then
            Row "データベースにつながる", False, Err.Description, _
                "「プロバイダーが見つかりません」→ <b>Microsoft Access Database Engine 2016 " & _
                "Redistributable</b> を入れてください。<br>" & _
                "入れてあるのに出る → <b>アプリケーション プールのビット数</b>が合っていません" & _
                "（64bit の ACE なら「32 ビット アプリケーションの有効化」は False）。<br>" & _
                "「排他的に開けません」→ 誰かが Access で開いています。"
            Err.Clear
        Else
            Row "データベースにつながる", True, "つながりました", ""
            n = -1
            Set rs = conn.Execute("SELECT Count(*) FROM [M_担当者]")
            If Err.Number = 0 Then n = rs.Fields(0).Value : rs.Close
            Err.Clear
            Row "担当者マスタが読める", (n >= 0), _
                IIf2(n >= 0, n & " 名", "読めません"), _
                "0 名や読めない場合は、置いた .accdb が古い可能性があります。"

            hasLogon = False
            Set rs = conn.Execute("SELECT Count(*) FROM [M_担当者] WHERE [ログオン名] Is Null")
            If Err.Number = 0 Then hasLogon = True : rs.Close
            Err.Clear
            Row "担当者マスタに「ログオン名」の欄がある", hasLogon, _
                IIf2(hasLogon, "あります", "ありません"), _
                "無い場合は、02_データベース_Access\任意_すでに使い始めている場合\ の" & _
                " <code>ログオン名の欄を足す.vbs</code> を実行してください。"
            conn.Close
        End If
    End If
End If
On Error GoTo 0
%>
</table>

<h2>4. 帳票の PDF 印刷</h2>
<table>
<tr><th>結果</th><th>見たもの</th><th>いまの値</th><th>直しかた</th></tr>
<%
edgePath = ""
On Error Resume Next
If Not hasFso Then
    Row "ファイルを扱う部品", False, "使えません", _
        "上の「3. データベース」の 1 行目をご覧ください。ここも確かめられません。"
End If
txt = ReadUtf8(Server.MapPath("include/config.asp"))
If Len(txt) > 0 Then
    i = InStr(txt, "Const PDF_EDGE_EXE")
    If i > 0 Then
        txt = Mid(txt, i)
        i = InStr(txt, """")
        If i > 0 Then
            txt = Mid(txt, i + 1)
            i = InStr(txt, """")
            If i > 0 Then edgePath = Left(txt, i - 1)
        End If
    End If
End If
Err.Clear

Row "変換プログラム (Edge) がある", fso.FileExists(edgePath), _
    IIf2(Len(edgePath) > 0, edgePath, "(読み取れませんでした)"), _
    "無い場合は include\config.asp の <code>PDF_EDGE_EXE</code> を実際の場所に直すか、" & _
    "wkhtmltopdf を入れて <code>PDF_ENGINE</code> を <code>""wkhtmltopdf""</code> にしてください。" & _
    "（PDF が作れなくても、ブラウザからの印刷はできます）"

Err.Clear
If Not IsObject(sh) Then Set sh = Server.CreateObject("WScript.Shell")
Row "外部プログラムを呼び出せる", (Err.Number = 0), _
    IIf2(Err.Number = 0, "呼び出せます", Err.Description), _
    "呼び出せない場合、PDF 印刷は使えません。ブラウザからの印刷をお使いください。"
Err.Clear

tmpDir = ""
Set f = Nothing
tmpDir = fso.GetSpecialFolder(2).Path
probe = fso.BuildPath(tmpDir, "nippou_write_test.tmp")
Err.Clear
Set f = fso.CreateTextFile(probe, True)
If Err.Number = 0 Then
    f.Close : fso.DeleteFile probe, True
    Row "作業用フォルダに書ける", True, tmpDir, ""
Else
    Row "作業用フォルダに書ける", False, Err.Description, _
        "そのフォルダに、アプリケーション プール ID の「変更」権限を与えてください。" & _
        "または include\config.asp の <code>PDF_WORK_DIR</code> に別のフォルダを指定してください。"
End If
Err.Clear
On Error GoTo 0
%>
</table>

<h2>5. ファイルがそろっているか</h2>
<table>
<tr><th>結果</th><th>見たもの</th><th>いまの値</th><th>直しかた</th></tr>
<%
Dim need, nm
need = Array("default.asp", "staff.asp", "entry.asp", "tasks.asp", "daily.asp", _
             "report.asp", "printpdf.asp", "summary.asp", "check.asp", "master.asp", _
             "error.asp", "css/style.css", "include/config.asp", "include/db.asp", _
             "include/auth.asp", _
             "include/layout.asp", "include/sheet.asp", "include/sql.asp", "include/pdf.asp")
On Error Resume Next
If Not hasFso Then
    Row "ファイルの確認", False, "できません", _
        "ファイルを扱う部品が使えないため、ここは確かめられません。"
End If
For i = 0 To UBound(need)
    nm = need(i)
    Row nm, fso.FileExists(Server.MapPath(nm)), _
        IIf2(fso.FileExists(Server.MapPath(nm)), "あります", "ありません"), _
        "wwwroot フォルダの中身を、フォルダの形のまま置いてください。"
Next
Err.Clear
On Error GoTo 0
%>
</table>

<h2>6. 各ページを開いてみる</h2>
<p class="lead" style="margin:0 0 10px">
  このサーバー自身から各ページを呼んで、返ってきた結果を見ます。
  <b>500</b> が出たページが、止まっている場所です。
  （<b>401</b> は Windows 認証が有効なときに出ます。異常ではありません）
</p>
<table>
<tr><th>結果</th><th>ページ</th><th>いまの値</th><th>直しかた</th></tr>
<%
Dim http, pages, base, u, st
base = "http://127.0.0.1"
If "" & Request.ServerVariables("SERVER_PORT") <> "80" Then
    base = base & ":" & Request.ServerVariables("SERVER_PORT")
End If
u = "" & Request.ServerVariables("URL")
If InStrRev(u, "/") > 0 Then u = Left(u, InStrRev(u, "/"))
base = base & u

pages = Array("default.asp", "entry.asp", "tasks.asp", "staff.asp", "daily.asp", _
              "report.asp", "summary.asp", "check.asp", "master.asp")
On Error Resume Next
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
If Err.Number <> 0 Then
    Row "ページの呼び出し", Null, "できません", _
        "この確認だけができません。ほかの結果をご覧ください。"
    Err.Clear
Else
    http.setTimeouts 3000, 3000, 8000, 8000
    Dim okPage, hint, code
    For i = 0 To UBound(pages)
        st = "" : hint = "" : okPage = False : code = 0
        Err.Clear
        http.open "GET", base & pages(i), False
        http.send
        If Err.Number <> 0 Then
            st = "呼び出せませんでした（" & Err.Description & "）"
            hint = "この確認だけができません。ほかの結果をご覧ください。"
            Err.Clear
        Else
            code = http.status
            st = code & " " & http.statusText
            okPage = (code < 500)
            If code = 500 Then
                hint = "<b>このページで止まっています。</b>IIS マネージャー →「ASP」→" & _
                       "「デバッグのプロパティ」→ <b>ブラウザーにエラーを送信する = True</b>" & _
                       " にしてから、このページを直接開くと理由が出ます。"
            ElseIf code = 401 Then
                hint = "Windows 認証が有効なので、この確認からは開けません（異常ではありません）。"
            End If
        End If
        Row pages(i), okPage, st, hint
    Next
End If
On Error GoTo 0
%>
</table>

<div class="sum">
  <b>OK <%= gOK %> 件 ／ NG <%= gNG %> 件</b><br>
  <% If gNG = 0 Then %>
    すべて通りました。<a href="default.asp">メニューを開く</a> と使い始められます。
  <% Else %>
    NG の行の「直しかた」を上から順に片づけてください。直したらこのページを再読み込みします。
  <% End If %>
  <p style="margin:10px 0 0; color:#5b6b7c; font-size:13px">
    このページ (setup_check.asp) は設置が終わったら消してかまいません。
  </p>
</div>
</main></body></html>
<%
Function IIf2(cond, a, b)
    If cond Then IIf2 = a Else IIf2 = b
End Function
%>
