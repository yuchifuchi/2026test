<%
' ============================================================
'  役割（パート職員用 / 職員用）の判定
'
'  認証はしない。置き場所（URL）を part と staff の 2 つに分け、
'  それぞれの config.asp の ROLE で役割を決める。
'  part フォルダには職員用の画面ファイルそのものを置かないので、そもそも開けない。
'  ここでの判定は「staff フォルダの config.asp を取り違えた」ときの二重の歯止め。
'
'  ROLE が空のときだけ、IIS が教えてくれるログオン名を M_担当者.[ログオン名] で引いて決める
'  （将来 IIS で Windows 認証を使うようにしたくなったとき用の道。今は使っていない）。
' ============================================================

Function IsStaff()
    Dim r, u
    r = LCase(Trim(ROLE))
    If r = "staff" Then
        IsStaff = True
    ElseIf r = "" Then
        u = AuthLogonName()
        If u = "" Then
            IsStaff = False
        Else
            IsStaff = (DbScalar(SqlStaffByLogon(), Array(u), 0) > 0)
        End If
    Else
        ' "part" と、それ以外の書き間違いはパート職員用として扱う（広く開けるより狭い方が安全）
        IsStaff = False
    End If
End Function

' Windows 認証のときのログオン名（「ドメイン\名前」の「名前」の部分）。匿名なら ""。
Function AuthLogonName()
    Dim u, i
    u = Request.ServerVariables("LOGON_USER")
    i = InStrRev(u, "\")
    If i > 0 Then u = Mid(u, i + 1)
    AuthLogonName = u
End Function

' 職員用の画面の先頭で呼ぶ。職員でなければ説明を出して止める。
' On Error Resume Next が生きている所で呼ばないこと（止まらなくなる。手順書 7-2）。
Sub RequireStaff()
    If IsStaff() Then Exit Sub
    Response.Clear
    Response.Status = "200 OK"
    Response.Write "<!DOCTYPE html>" & vbCrLf & "<html lang=""ja""><head><meta charset=""utf-8""><meta http-equiv=""X-UA-Compatible"" content=""IE=edge"">"
    Response.Write "<title>職員用の画面です</title><link rel=""stylesheet"" href=""css/style.css""></head><body>"
    Response.Write "<div class=""band""><div class=""band-in""><span class=""app"">" & DbH(APP_NAME) & "</span></div></div>"
    Response.Write "<div class=""wrap""><h1>この画面は職員用です</h1>"
    Response.Write "<p>パート職員の方は、ふだんお使いの入力画面から入力してください。</p>"
    Response.Write "<p><a class=""btn"" href=""default.asp"">入力画面のメニューへ</a></p>"
    Response.Write "<p class=""note"">（職員の方でこの画面が出たときは、このフォルダの config.asp の ROLE が ""staff"" になっているか確かめてください）</p>"
    Response.Write "</div></body></html>"
    Response.End
End Sub
%>
