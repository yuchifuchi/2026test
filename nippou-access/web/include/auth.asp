<%
' =============================================================================
'  利用者の特定と権限
'
'  認証は IIS の Windows 認証に任せる (このファイルはパスワードを扱わない)。
'  IIS マネージャーで対象サイトの
'      「Windows 認証」= 有効 ／ 「匿名認証」= 無効
'  にすると、LOGON_USER にドメイン\ユーザー名が入る。
'
'  権限は 2 段階だけ。増やすと運用が回らなくなるため。
'      一般 : 受付入力・その他業務・日報・集計表・入力もれ
'      管理 : 上記 + マスタ保守
' =============================================================================

' マスタ保守を触れる人のログオン名を「,」区切りで並べる。ドメイン名は書かない。
' 例) Const ADMIN_USERS = "t-okada,y-fujita"
' 空のままにすると全員が触れる (導入直後の既定)。
Const ADMIN_USERS = ""

' ログオン名。ドメイン部分は落とす。
Function CurrentUser()
    Dim u
    u = Trim(Request.ServerVariables("LOGON_USER") & "")
    If Len(u) = 0 Then u = Trim(Request.ServerVariables("AUTH_USER") & "")
    If Len(u) = 0 Then
        CurrentUser = "(未認証)"
    Else
        If InStr(u, "\") > 0 Then u = Mid(u, InStr(u, "\") + 1)
        CurrentUser = u
    End If
End Function

Function IsAdmin()
    Dim list, i, me_
    If Len(Trim(ADMIN_USERS)) = 0 Then
        IsAdmin = True                        ' 未設定なら全員可
        Exit Function
    End If
    me_ = LCase(CurrentUser())
    list = Split(ADMIN_USERS, ",")
    For i = 0 To UBound(list)
        If LCase(Trim(list(i))) = me_ Then
            IsAdmin = True
            Exit Function
        End If
    Next
    IsAdmin = False
End Function

' 管理者専用ページの先頭で呼ぶ。権限が無ければその場で止める。
Sub RequireAdmin()
    If IsAdmin() Then Exit Sub
    Response.Write "<!doctype html><html lang=""ja""><head><meta charset=""utf-8"">" & _
        "<title>権限がありません</title>" & _
        "<link rel=""stylesheet"" href=""css/style.css""></head><body><main>" & _
        "<h1>この画面は担当者だけが開けます</h1>" & _
        "<p class=""notice warn"">" & _
        "マスタ保守は、担当者として登録された方だけが操作できます。<br>" & _
        "変更が必要なときは担当者にご連絡ください。</p>" & _
        "<p><a class=""btn"" href=""default.asp"">メニューに戻る</a></p>" & _
        "</main></body></html>"
    Response.End
End Sub
%>
