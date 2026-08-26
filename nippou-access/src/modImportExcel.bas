Attribute VB_Name = "modImportExcel"
Option Compare Database
Option Explicit

'==============================================================================
' 現行 Excel からの取込
'
' 取込元は「集計表（…）.xlsm」の『データ』シート。列の意味は現行の転記マクロと同じ:
'   A=日付  B=製品名  C〜H=集計列 3〜8 の件数  J=備考(区分名)  K=担当者(姓)
'   L=備考2  M=備考3
'
' 区分は「旧転記名」で引き当てる。現行 Sheet2 の 3 行目の文言をマスタに
' 保存してあるのはこのため。引き当てられなかった行は T_取込ログ に残す。
'==============================================================================

Private Const LINK_NAME As String = "_取込元データ"

Public Sub Import_Excel_Dialog()
    Dim path As String
    path = PickExcelFile()
    If Len(path) = 0 Then Exit Sub

    If MsgBox("次のファイルから取り込みます。" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _
              "同じ 日付・担当者・区分・製品 の行が既にある場合は、件数を上書きします。" & vbCrLf & _
              "続けますか？", vbQuestion + vbOKCancel, APP_NAME) = vbCancel Then Exit Sub

    Import_集計表 path
End Sub

Public Sub Import_集計表(ByVal path As String)
    Dim rs As DAO.Recordset
    Dim rowNo As Long, okCnt As Long, ngCnt As Long, skipCnt As Long
    Dim dt As Variant, prodName As String, kubunName As String, opName As String
    Dim colId As Long, cnt As Long, opId As Long, kbId As Long, prId As Long
    Dim biko2 As Variant, biko3 As Variant
    Dim i As Long

    On Error GoTo Fail

    EnsureImportLog
    ExecSQL "DELETE FROM [T_取込ログ] WHERE [取込元]=" & Q(path)

    DropLink
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, LINK_NAME, _
                              path, False, "データ!A9:R5000"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & LINK_NAME & "]", dbOpenSnapshot)

    Do While Not rs.EOF
        rowNo = rowNo + 1
        dt = rs.Fields(0).Value                       ' A: 日付
        prodName = NzS(rs.Fields(1).Value)            ' B: 製品名
        kubunName = NzS(rs.Fields(9).Value)           ' J: 備考 (旧転記名)
        opName = NzS(rs.Fields(10).Value)             ' K: 担当者 (姓)
        biko2 = rs.Fields(11).Value                   ' L: 備考2
        biko3 = rs.Fields(12).Value                   ' M: 備考3

        ' C〜H (集計列 3〜8) のうち 0 でない列を探す
        colId = 0: cnt = 0
        For i = 2 To 7
            If IsNumeric(rs.Fields(i).Value) Then
                If Nz(rs.Fields(i).Value, 0) <> 0 Then
                    colId = i + 1
                    cnt = CLng(rs.Fields(i).Value)
                    Exit For
                End If
            End If
        Next i

        If Not IsDate(dt) Or colId = 0 Then
            skipCnt = skipCnt + 1                     ' 見出し行・SUMIF 用の空行など
        Else
            opId = FindOperator(opName)
            kbId = FindKubun(kubunName, colId)
            prId = FindProduct(prodName, kbId)

            If opId = 0 Or kbId = 0 Then
                LogImport path, rowNo, CStr(dt), opName, prodName, kubunName, cnt, _
                    IIf(opId = 0, "担当者『" & opName & "』がマスタにありません。", "") & _
                    IIf(kbId = 0, "区分『" & kubunName & "』(集計列" & colId & _
                        ") がマスタにありません。", "")
                ngCnt = ngCnt + 1
            Else
                UpsertJuden CDate(dt), opId, kbId, prId, cnt, biko2, biko3
                okCnt = okCnt + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    DropLink

    ' 取り込んだ日付には日報のヘッダ行と出勤登録を用意しておく
    NormalizeAfterImport

    MsgBox "取込が終わりました。" & vbCrLf & vbCrLf & _
           "取込        : " & okCnt & " 行" & vbCrLf & _
           "引き当て失敗: " & ngCnt & " 行" & vbCrLf & _
           "対象外      : " & skipCnt & " 行 (見出し・空行など)" & vbCrLf & vbCrLf & _
           IIf(ngCnt > 0, "失敗した行は T_取込ログ を開いて確認してください。", ""), _
           IIf(ngCnt > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub

Fail:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    DropLink
    MsgBox "取込中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

'------------------------------------------------------------------------------
' 引き当て
'------------------------------------------------------------------------------
' 担当者は 姓 で引く。現行の集計表には姓しか入っていないため。
' 姓で見つからない場合は「旧B1値」(現行 Sheet1!B1 の値) でも探す。
Private Function FindOperator(ByVal nm As String) As Long
    Dim v As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function

    v = DLookup("[担当者ID]", "M_担当者", "[姓]=" & Q(nm))
    If IsNull(v) Then v = DLookup("[担当者ID]", "M_担当者", "[氏名]=" & Q(nm))
    If IsNull(v) Then
        ' 髙/高 のような異体字ゆれを吸収する
        v = DLookup("[担当者ID]", "M_担当者", "[姓] Like " & Q("*" & Mid$(nm, 2) & "*"))
    End If
    FindOperator = Nz(v, 0)
End Function

Private Function FindKubun(ByVal nm As String, ByVal colId As Long) As Long
    Dim v As Variant
    nm = Trim$(nm)

    If Len(nm) > 0 Then
        v = DLookup("[区分ID]", "M_区分", "[旧転記名]=" & Q(nm) & " And [集計列ID]=" & colId)
        If IsNull(v) Then v = DLookup("[区分ID]", "M_区分", "[区分名]=" & Q(nm) & _
                                      " And [集計列ID]=" & colId)
        If IsNull(v) Then v = DLookup("[区分ID]", "M_区分", "[旧転記名]=" & Q(nm))
    Else
        ' 備考が空の行は、集計列だけで一意に決まる区分 (申込関係・抽選結果) を使う
        v = DLookup("[区分ID]", "M_区分", _
                    "[集計列ID]=" & colId & " And (([旧転記名] Is Null) Or ([旧転記名]=''))")
    End If
    FindKubun = Nz(v, 0)
End Function

' 製品名は上段ブロックと【その他】ブロックで重複するので、
' まず区分と同じブロックで探し、無ければ名前だけで探す。
Private Function FindProduct(ByVal nm As String, ByVal kbId As Long) As Long
    Dim v As Variant, blockId As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then
        FindProduct = 0
        Exit Function
    End If

    If kbId > 0 Then
        blockId = DLookup("[ブロックID]", "M_区分", "[区分ID]=" & kbId)
        If Not IsNull(blockId) Then
            v = DLookup("[製品ID]", "M_製品", _
                        "[製品名]=" & Q(nm) & " And [ブロックID]=" & blockId)
        End If
    End If
    If IsNull(v) Then v = DLookup("[製品ID]", "M_製品", "[製品名]=" & Q(nm))
    FindProduct = Nz(v, 0)
End Function

'------------------------------------------------------------------------------
' 書込
'------------------------------------------------------------------------------
Private Sub UpsertJuden(ByVal dt As Date, ByVal opId As Long, ByVal kbId As Long, _
                        ByVal prId As Long, ByVal cnt As Long, _
                        ByVal biko2 As Variant, ByVal biko3 As Variant)
    Dim where As String
    where = "[対象日]=" & D(dt) & " And [担当者ID]=" & opId & _
            " And [区分ID]=" & kbId & " And [製品ID]=" & prId

    If DCount("*", "T_受電", where) > 0 Then
        ExecSQL "UPDATE [T_受電] SET [件数]=" & cnt & _
                ",[備考2]=" & Q(NzS(biko2)) & ",[備考3]=" & Q(NzS(biko3)) & _
                ",[更新日時]=Now() WHERE " & where
    Else
        ExecSQL "INSERT INTO [T_受電] " & _
                "([対象日],[担当者ID],[区分ID],[製品ID],[件数],[備考2],[備考3]," & _
                " [登録日時],[更新日時],[登録者]) VALUES (" & _
                D(dt) & "," & opId & "," & kbId & "," & prId & "," & cnt & "," & _
                Q(NzS(biko2)) & "," & Q(NzS(biko3)) & ",Now(),Now()," & _
                Q("取込:" & CurrentUserName) & ")"
    End If
End Sub

' 取り込んだ日付ごとに、日報ヘッダと出勤登録を補う。
Private Sub NormalizeAfterImport()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT DISTINCT [対象日] FROM [T_受電] " & _
        "WHERE [対象日] NOT IN (SELECT [対象日] FROM [T_日報])")
    Do While Not rs.EOF
        Ensure日報 CDate(rs![対象日])
        rs.MoveNext
    Loop
    rs.Close

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT DISTINCT J.[対象日], J.[担当者ID] FROM [T_受電] AS J " & _
        "WHERE NOT EXISTS (SELECT 1 FROM [T_出勤] AS A " & _
        " WHERE A.[対象日]=J.[対象日] AND A.[担当者ID]=J.[担当者ID])")
    Do While Not rs.EOF
        Ensure出勤 CDate(rs![対象日]), CLng(rs![担当者ID])
        rs.MoveNext
    Loop
    rs.Close
End Sub

'------------------------------------------------------------------------------
' 取込ログ
'------------------------------------------------------------------------------
Public Sub EnsureImportLog()
    If TableExists("T_取込ログ") Then Exit Sub
    ExecDDL "CREATE TABLE [T_取込ログ] (" & _
        "[ログID] COUNTER NOT NULL CONSTRAINT [PK_取込ログ] PRIMARY KEY," & _
        "[取込日時] DATETIME," & _
        "[取込元] TEXT(255)," & _
        "[行番号] LONG," & _
        "[日付] TEXT(50)," & _
        "[担当者] TEXT(100)," & _
        "[製品名] TEXT(100)," & _
        "[区分名] TEXT(100)," & _
        "[件数] LONG," & _
        "[理由] MEMO)"
End Sub

Private Sub LogImport(ByVal src As String, ByVal rowNo As Long, ByVal dt As String, _
                      ByVal opName As String, ByVal prodName As String, _
                      ByVal kubunName As String, ByVal cnt As Long, ByVal reason As String)
    ExecSQL "INSERT INTO [T_取込ログ] " & _
            "([取込日時],[取込元],[行番号],[日付],[担当者],[製品名],[区分名],[件数],[理由]) " & _
            "VALUES (Now()," & Q(src) & "," & rowNo & "," & Q(dt) & "," & Q(opName) & "," & _
            Q(prodName) & "," & Q(kubunName) & "," & cnt & "," & Q(reason) & ")"
End Sub

'------------------------------------------------------------------------------
' 小物
'------------------------------------------------------------------------------
Private Sub DropLink()
    On Error Resume Next
    DoCmd.DeleteObject acTable, LINK_NAME
    Err.Clear
End Sub

Private Function NzS(ByVal v As Variant) As String
    NzS = Trim$(Nz(v, ""))
End Function

Private Function PickExcelFile() As String
    Dim fd As Object
    On Error GoTo Manual
    Set fd = Application.FileDialog(3)          ' msoFileDialogFilePicker
    With fd
        .Title = "取り込む集計表を選んでください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ブック", "*.xlsm; *.xlsx; *.xls"
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
    Exit Function
Manual:
    PickExcelFile = InputBox("取り込む集計表のフルパスを入力してください。", APP_NAME)
End Function
