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
'
'  どちらになるかは、次の順に見て決める。
'      1) M_担当者 に、そのログオン名の行があれば その [職員区分]
'         → マスタ保守の画面から、職員のかたが変更できる
'      2) 無ければ config.asp の STAFF_USERS の一覧
'         → マスタに登録する前の予備
'      3) それも空なら 全員が職員 (導入直後の既定)
' =============================================================================

' 予備の一覧 STAFF_USERS は config.asp にあります。

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
' 1 ページのうちに何度も呼ばれる (メニュー・見出し・入口の判定) ので、
' 1 回調べたら覚えておく。
Dim gIsStaff, gIsStaffKnown

Function IsStaff()
    Dim list, i, who, kbn
    If gIsStaffKnown Then
        IsStaff = gIsStaff
        Exit Function
    End If
    gIsStaffKnown = True
    who = Trim(CurrentUser())

    ' --- 1) マスタで決める ---------------------------------------------------
    ' 退職日を過ぎた人・「有効」を外した人は、ここでは見つからない。
    '
    ' 先に接続だけ済ませておく。つながらない場合はここで画面を出して終わる。
    ' 下の On Error Resume Next の中で接続に失敗すると、その「終わる」が
    ' 効かなくなり、結局 ASP のエラー画面になってしまうため。
    OpenDb

    kbn = ""
    On Error Resume Next                     ' [ログオン名] が無い古い .accdb 対策
    kbn = DbScalar("SELECT [職員区分] FROM [M_担当者] " & _
                   "WHERE [ログオン名]=? AND [有効]=True " & _
                   "  AND ([在籍終了日] Is Null Or [在籍終了日]>=?)", _
                   Array(who, Date()), "")
    If Err.Number <> 0 Then
        kbn = ""                              ' [ログオン名] の無い古い .accdb でも動く
        Err.Clear
    End If
    On Error GoTo 0
    If Len(Trim("" & kbn)) > 0 Then
        gIsStaff = (Trim("" & kbn) = "職員")
        IsStaff = gIsStaff
        Exit Function
    End If

    ' --- 2) マスタに無ければ、このファイルの一覧で決める ---------------------
    If Len(Trim(STAFF_USERS)) = 0 Then
        gIsStaff = True                       ' 3) どちらも未設定なら全員が職員
        IsStaff = gIsStaff
        Exit Function
    End If
    gIsStaff = False
    list = Split(STAFF_USERS, ",")
    For i = 0 To UBound(list)
        If LCase(Trim(list(i))) = LCase(who) Then gIsStaff = True
    Next
    IsStaff = gIsStaff
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
        "この画面を使う必要があるときは、職員のかたにご連絡ください。<br>" & _
        "（マスタ保守の担当者一覧で、あなたのログオン名 " & H(CurrentUser()) & " を" & _
        "「職員」の行に入れてもらうと開けるようになります）</p>" & _
        "<p><a class=""btn"" href=""default.asp"">メニューに戻る</a></p>" & _
        "</main></body></html>"
    Response.End
End Sub
%>
