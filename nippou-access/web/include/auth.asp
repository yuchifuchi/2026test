<%
' =============================================================================
'  利用者の特定と役割
'
'  認証は IIS の Windows 認証に任せる (このファイルはパスワードを扱わない)。
'  IIS マネージャーで対象サイトの
'      「Windows 認証」= 有効 ／ 「匿名認証」= 無効
'  にすると、LOGON_USER にドメイン\ユーザー名が入る。
'
'  役割は 2 つだけ。増やすと運用が回らなくなるため。
'      パート職員 : 受付入力・その他業務            (default.asp が入口)
'      職員       : 上記 + 日報・帳票印刷・集計表・入力もれ・マスタ保守
'                                                   (staff.asp が入口)
' =============================================================================

' 職員のログオン名を「,」区切りで並べる。ドメイン名は書かない。
' 例) Const STAFF_USERS = "t-okada,y-fujita"
' 空のままにすると全員が職員になる (導入直後の既定)。
' 名前を 1 つでも書いた時点で、書かれていない人はパート職員の画面だけになる。
Const STAFF_USERS = ""

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

' この人は職員か。
Function IsStaff()
    Dim list, i, who
    If Len(Trim(STAFF_USERS)) = 0 Then
        IsStaff = True                        ' 未設定なら全員が職員
        Exit Function
    End If
    who = LCase(CurrentUser())
    list = Split(STAFF_USERS, ",")
    For i = 0 To UBound(list)
        If LCase(Trim(list(i))) = who Then
            IsStaff = True
            Exit Function
        End If
    Next
    IsStaff = False
End Function

' その人にとっての「メニュー」。ロゴや「メニューに戻る」の行き先。
Function HomePage()
    If IsStaff() Then
        HomePage = "staff.asp"
    Else
        HomePage = "default.asp"
    End If
End Function

' 職員専用ページの先頭で呼ぶ。パート職員ならその場で止める。
Sub RequireStaff()
    If IsStaff() Then Exit Sub
    Response.Write "<!doctype html><html lang=""ja""><head><meta charset=""utf-8"">" & _
        "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & _
        "<title>この画面は職員用です</title>" & _
        "<link rel=""stylesheet"" href=""css/style.css""></head><body><main>" & _
        "<h1>この画面は職員のかた専用です</h1>" & _
        "<p class=""notice warn"">" & _
        "日報・帳票印刷・集計表・入力もれ・マスタ保守は、職員のかたが操作します。<br>" & _
        "パート職員のかたは「受付入力」と「その他業務」をお使いください。<br>" & _
        "この画面を使う必要があるときは、課の担当者にご連絡ください。</p>" & _
        "<p><a class=""btn"" href=""default.asp"">メニューに戻る</a></p>" & _
        "</main></body></html>"
    Response.End
End Sub
%>
