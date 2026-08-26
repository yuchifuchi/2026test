Attribute VB_Name = "modApp"
Option Compare Database
Option Explicit

'==============================================================================
' 実行時の処理
'
' 画面のボタン・イベントはすべて「=App_Click("名前")」「=App_Event("名前")」の
' 形でここへ集まる。フォームモジュールを持たないので、画面を作り直しても
' 処理は失われない。
'==============================================================================

'------------------------------------------------------------------------------
' ボタン
'------------------------------------------------------------------------------
Public Function App_Click(ByVal action As String) As Boolean
    On Error GoTo Fail
    App_Click = True

    Select Case action
        Case "日次入力":        DoCmd.OpenForm "F_日次入力"
        Case "日報":            Open日報 Date
        Case "日報印刷":        Print日報
        Case "集計表":          DoCmd.OpenForm "F_集計表"
        Case "未入力チェック":  DoCmd.OpenForm "F_未入力チェック"
        Case "受電明細":        DoCmd.OpenForm "F_受電明細"
        Case "業務実績":        Open業務実績
        Case "マスタ_担当者":   DoCmd.OpenForm "F_マスタ_担当者"
        Case "マスタ_製品":     DoCmd.OpenForm "F_マスタ_製品"
        Case "マスタ_区分":     DoCmd.OpenForm "F_マスタ_区分"
        Case "マスタ_業務項目": DoCmd.OpenForm "F_マスタ_業務項目"
        Case "取込":            Import_Excel_Dialog
        Case "終了":            DoCmd.Quit acQuitSaveAll
        Case "閉じる":          DoCmd.Close acForm, Screen.ActiveForm.Name, acSaveNo

        Case "日次入力_再表示": Refresh日次入力
        Case "日報確定":        Confirm日報
        Case "出勤一括登録":    Fill出勤
        Case "集計表_表示":     Show集計表
        Case "集計表_今週":     Set集計表期間 0
        Case "集計表_先週":     Set集計表期間 -7
        Case "集計表_印刷":     Print集計表
        Case "未入力_実行":     Show未入力

        Case Else
            MsgBox "未定義の操作です: " & action, vbExclamation, APP_NAME
    End Select
    Exit Function

Fail:
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "操作: " & action & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Function

'------------------------------------------------------------------------------
' フォームのイベント
'------------------------------------------------------------------------------
Public Function App_Event(ByVal action As String) As Boolean
    On Error GoTo Fail
    App_Event = True
    Dim f As Access.Form
    Set f = Screen.ActiveForm

    Select Case action
        Case "担当者_新規"
            SafeSet f, "担当者ID", NextId("M_担当者", "担当者ID")
            SafeSet f, "表示順", NextId("M_担当者", "表示順")
            SafeSet f, "職員区分", "パート"
            SafeSet f, "有効", True
            SafeSet f, "在籍開始日", Date
        Case "製品_新規"
            SafeSet f, "製品ID", NextId("M_製品", "製品ID")
            SafeSet f, "表示順", 99
            SafeSet f, "有効", True
            SafeSet f, "適用開始日", Date
        Case "区分_新規"
            SafeSet f, "区分ID", NextId("M_区分", "区分ID")
            SafeSet f, "表示順", 99
            SafeSet f, "有効", True
        Case "業務項目_新規"
            SafeSet f, "業務項目ID", NextId("M_業務項目", "業務項目ID")
            SafeSet f, "表示順", NextId("M_業務項目", "表示順")
            SafeSet f, "有効", True
    End Select
    Exit Function

Fail:
    Err.Clear
End Function

'------------------------------------------------------------------------------
' 日次入力
'------------------------------------------------------------------------------
Private Sub Refresh日次入力()
    Dim f As Access.Form
    Set f = Forms("F_日次入力")

    If IsNull(f!対象日) Then
        MsgBox "対象日を入力してください。", vbExclamation, APP_NAME
        Exit Sub
    End If
    If IsNull(f!担当者ID) Then
        MsgBox "担当者を選んでください。", vbExclamation, APP_NAME
        Exit Sub
    End If

    ' 日報のヘッダ行と出勤登録を先に用意しておく。
    ' 現行 Excel では「入力したのに集計に出てこない」ことがあったので、
    ' 入力を始めた時点で必ず出勤者として登録される作りにしてある。
    Ensure日報 CDate(f!対象日)
    Ensure出勤 CDate(f!対象日), CLng(f!担当者ID)

    f!sub明細.Form.Requery
    f.Requery
End Sub

Private Sub Open業務実績()
    Dim f As Access.Form, dt As Date, opId As Long
    Set f = Forms("F_日次入力")
    If IsNull(f!対象日) Or IsNull(f!担当者ID) Then
        MsgBox "対象日と担当者を先に選んでください。", vbExclamation, APP_NAME
        Exit Sub
    End If
    dt = CDate(f!対象日)
    opId = CLng(f!担当者ID)

    DoCmd.OpenForm "F_業務実績", , , _
        "[対象日]=" & D(dt) & " And [担当者ID]=" & opId
    With Forms("F_業務実績")
        .Controls("対象日").DefaultValue = D(dt)
        .Controls("担当者ID").DefaultValue = opId
        .caption = "電話応対以外の業務  " & Format(dt, "yyyy/mm/dd") & _
                   "  " & Nz(DLookup("[氏名]", "M_担当者", "[担当者ID]=" & opId), "")
    End With
End Sub

'------------------------------------------------------------------------------
' 日報
'------------------------------------------------------------------------------
Private Sub Open日報(ByVal dt As Date)
    Ensure日報 dt
    DoCmd.OpenForm "F_日報", , , "[対象日]=" & D(dt)
End Sub

Public Sub Ensure日報(ByVal dt As Date)
    If DCount("*", "T_日報", "[対象日]=" & D(dt)) > 0 Then Exit Sub
    ExecSQL "INSERT INTO [T_日報] ([対象日],[状態],[更新日時]) " & _
            "VALUES (" & D(dt) & ",'入力中',Now())"
End Sub

Public Sub Ensure出勤(ByVal dt As Date, ByVal opId As Long)
    If DCount("*", "T_出勤", "[対象日]=" & D(dt) & " And [担当者ID]=" & opId) > 0 Then Exit Sub
    ExecSQL "INSERT INTO [T_出勤] ([対象日],[担当者ID]) VALUES (" & D(dt) & "," & opId & ")"
End Sub

' 帳票・画面から参照する集計値。Q_日報_受電 に該当日が無ければ 0 を返す。
Public Function DailyValue(ByVal dt As Variant, ByVal fld As String) As Long
    On Error GoTo Zero
    If IsNull(dt) Then Exit Function
    DailyValue = Nz(DLookup("[" & fld & "]", "Q_日報_受電", "[対象日]=" & D(dt)), 0)
    Exit Function
Zero:
    DailyValue = 0
End Function

' 業務項目 n 件目の件数。帳票 R_日報 の ①～⑬ 欄が参照する。
Public Function TaskValue(ByVal dt As Variant, ByVal taskId As Long) As Variant
    On Error GoTo Blank
    If IsNull(dt) Then Exit Function
    TaskValue = DLookup("[件数]", "Q_日報_業務", _
                        "[対象日]=" & D(dt) & " And [業務項目ID]=" & taskId)
    Exit Function
Blank:
    TaskValue = Null
End Function

' 出勤者の氏名 (n 人目)。帳票の氏名欄が参照する。
Public Function AttendeeName(ByVal dt As Variant, ByVal idx As Long) As Variant
    Dim rs As DAO.Recordset, i As Long
    On Error GoTo Blank
    If IsNull(dt) Then Exit Function
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [氏名] FROM [Q_日報_出勤] WHERE [対象日]=" & D(dt) & " ORDER BY [表示順]")
    Do While Not rs.EOF
        i = i + 1
        If i = idx Then
            AttendeeName = rs!氏名
            rs.Close
            Exit Function
        End If
        rs.MoveNext
    Loop
    rs.Close
Blank:
End Function

' 令和表記の日付。現行 印刷用シート W2 と同じ体裁にする。
Public Function WarekiLong(ByVal dt As Variant) As String
    Dim d0 As Date, g As String, y As Long
    If IsNull(dt) Then Exit Function
    d0 = CDate(dt)
    If d0 >= #5/1/2019# Then
        g = "令和": y = Year(d0) - 2018
    ElseIf d0 >= #1/8/1989# Then
        g = "平成": y = Year(d0) - 1988
    Else
        g = "昭和": y = Year(d0) - 1925
    End If
    WarekiLong = " " & g & "　" & y & "年　　" & Month(d0) & "月　　" & Day(d0) & "日（" & _
                 Mid("日月火水木金土", Weekday(d0), 1) & "）"
End Function

Private Sub Fill出勤()
    Dim f As Access.Form, dt As Date, rs As DAO.Recordset, n As Long
    Set f = Forms("F_日報")
    If IsNull(f!対象日) Then Exit Sub
    dt = CDate(f!対象日)

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [担当者ID] FROM [M_担当者] WHERE [有効]=True " & _
        " AND ([在籍開始日] Is Null Or [在籍開始日] <= " & D(dt) & ")" & _
        " AND ([在籍終了日] Is Null Or [在籍終了日] >= " & D(dt) & ")")
    Do While Not rs.EOF
        If DCount("*", "T_出勤", "[対象日]=" & D(dt) & " And [担当者ID]=" & rs!担当者ID) = 0 Then
            ExecSQL "INSERT INTO [T_出勤] ([対象日],[担当者ID]) VALUES (" & _
                    D(dt) & "," & rs!担当者ID & ")"
            n = n + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    f!sub出勤.Form.Requery
    MsgBox n & " 名を出勤者に追加しました。" & vbCrLf & _
           "休みの方はサブフォームの行を削除してください。", vbInformation, APP_NAME
End Sub

Private Sub Confirm日報()
    Dim f As Access.Form, dt As Date, missing As Long
    Set f = Forms("F_日報")
    If IsNull(f!対象日) Then Exit Sub
    dt = CDate(f!対象日)

    If f.Dirty Then f.Dirty = False

    missing = DCount("*", "Q_未入力チェック", "[対象日]=" & D(dt))
    If missing > 0 Then
        If MsgBox("出勤登録があるのに実績が 1 件も無い担当者が " & missing & " 名います。" & vbCrLf & _
                  "このまま確定しますか？" & vbCrLf & vbCrLf & _
                  "(「いいえ」を選ぶと入力もれチェック画面を開きます)", _
                  vbQuestion + vbYesNo, APP_NAME) = vbNo Then
            DoCmd.OpenForm "F_未入力チェック"
            Forms("F_未入力チェック")!対象日 = dt
            Show未入力
            Exit Sub
        End If
    End If

    ExecSQL "UPDATE [T_日報] SET [状態]='確定',[確定日時]=Now(),[更新日時]=Now() " & _
            "WHERE [対象日]=" & D(dt)
    f.Requery
    MsgBox Format(dt, "yyyy年m月d日") & " の日報を確定しました。", vbInformation, APP_NAME
End Sub

Private Sub Print日報()
    Dim dt As Variant
    dt = ActiveDate("F_日報")
    If IsNull(dt) Then dt = InputDate("印刷する日報の対象日を入力してください。")
    If IsNull(dt) Then Exit Sub
    Ensure日報 CDate(dt)
    DoCmd.OpenReport "R_日報", acViewPreview, , "[対象日]=" & D(dt)
End Sub

'------------------------------------------------------------------------------
' 集計表
'------------------------------------------------------------------------------
Private Sub Show集計表()
    Dim f As Access.Form, d1 As Date, d2 As Date
    Set f = Forms("F_集計表")
    If IsNull(f!開始日) Or IsNull(f!終了日) Then Exit Sub
    d1 = CDate(f!開始日): d2 = CDate(f!終了日)
    If d2 < d1 Then
        MsgBox "終了日が開始日より前になっています。", vbExclamation, APP_NAME
        Exit Sub
    End If

    f!sub集計.Form.RecordSource = _
        "SELECT * FROM [Q_週次集計指定] WHERE [対象日] Between " & D(d1) & " And " & D(d2)
    f!sub明細.Form.RecordSource = _
        "SELECT * FROM [Q_受電明細] WHERE [対象日] Between " & D(d1) & " And " & D(d2) & _
        " ORDER BY [対象日],[氏名],[区分ID]"
    f!sub集計.Form.Requery
    f!sub明細.Form.Requery
End Sub

Private Sub Set集計表期間(ByVal offsetDays As Long)
    Dim f As Access.Form, monday As Date
    Set f = Forms("F_集計表")
    monday = Date - Weekday(Date, vbMonday) + 1 + offsetDays
    f!開始日 = monday
    f!終了日 = monday + 4
    Show集計表
End Sub

Private Sub Print集計表()
    Dim f As Access.Form
    Set f = Forms("F_集計表")
    If IsNull(f!開始日) Or IsNull(f!終了日) Then Exit Sub
    DoCmd.OpenReport "R_集計表", acViewPreview, , _
        "[対象日] Between " & D(f!開始日) & " And " & D(f!終了日)
End Sub

Private Sub Show未入力()
    Dim f As Access.Form
    Set f = Forms("F_未入力チェック")
    If IsNull(f!対象日) Then Exit Sub
    f!sub結果.Form.RecordSource = _
        "SELECT * FROM [Q_未入力チェック] WHERE [対象日]=" & D(f!対象日)
    f!sub結果.Form.Requery
    If f!sub結果.Form.Recordset.RecordCount = 0 Then
        MsgBox "入力もれはありません。", vbInformation, APP_NAME
    End If
End Sub

'------------------------------------------------------------------------------
' 小物
'------------------------------------------------------------------------------
Public Function NextId(ByVal tbl As String, ByVal fld As String) As Long
    NextId = Nz(DMax("[" & fld & "]", tbl), 0) + 1
End Function

Public Function CurrentUserName() As String
    On Error Resume Next
    CurrentUserName = Environ$("USERNAME")
    If Len(CurrentUserName) = 0 Then CurrentUserName = "unknown"
End Function

Private Sub SafeSet(ByRef f As Access.Form, ByVal ctlName As String, ByVal v As Variant)
    On Error Resume Next
    f.Controls(ctlName).Value = v
    Err.Clear
End Sub

Private Function ActiveDate(ByVal formName As String) As Variant
    On Error GoTo Nope
    ActiveDate = Forms(formName)!対象日
    Exit Function
Nope:
    ActiveDate = Null
End Function

Private Function InputDate(ByVal prompt As String) As Variant
    Dim s As String
    s = InputBox(prompt, APP_NAME, Format(Date, "yyyy/mm/dd"))
    If Len(Trim$(s)) = 0 Then
        InputDate = Null
    ElseIf IsDate(s) Then
        InputDate = CDate(s)
    Else
        MsgBox "日付として読み取れませんでした: " & s, vbExclamation, APP_NAME
        InputDate = Null
    End If
End Function
