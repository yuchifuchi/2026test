Attribute VB_Name = "modImportExcel"
Option Compare Database
Option Explicit

'==============================================================================
' Œ»s Excel ‚©‚ç‚Ìæ
'
' æŒ³‚ÍuWŒv•\icj.xlsmv‚Ìwƒf[ƒ^xƒV[ƒgB—ñ‚ÌˆÓ–¡‚ÍŒ»s‚Ì“]‹Lƒ}ƒNƒ‚Æ“¯‚¶:
'   A=“ú•t  B=»•i–¼  C`H=WŒv—ñ 3`8 ‚ÌŒ”  J=”õl(‹æ•ª–¼)  K=’S“–Ò(©)
'   L=”õl2  M=”õl3
'
' ‹æ•ª‚Íu‹Œ“]‹L–¼v‚Åˆø‚«“–‚Ä‚éBŒ»s Sheet2 ‚Ì 3 s–Ú‚Ì•¶Œ¾‚ğƒ}ƒXƒ^‚É
' •Û‘¶‚µ‚Ä‚ ‚é‚Ì‚Í‚±‚Ì‚½‚ßBˆø‚«“–‚Ä‚ç‚ê‚È‚©‚Á‚½s‚Í T_æƒƒO ‚Éc‚·B
'==============================================================================

Private Const LINK_NAME As String = "_æŒ³ƒf[ƒ^"

Public Sub Import_Excel_Dialog()
    Dim path As String
    path = PickExcelFile()
    If Len(path) = 0 Then Exit Sub

    If MsgBox("Ÿ‚Ìƒtƒ@ƒCƒ‹‚©‚çæ‚è‚İ‚Ü‚·B" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _
              "“¯‚¶ “ú•tE’S“–ÒE‹æ•ªE»•i ‚Ìs‚ªŠù‚É‚ ‚éê‡‚ÍAŒ”‚ğã‘‚«‚µ‚Ü‚·B" & vbCrLf & _
              "‘±‚¯‚Ü‚·‚©H", vbQuestion + vbOKCancel, APP_NAME) = vbCancel Then Exit Sub

    Import_WŒv•\ path
End Sub

Public Sub Import_WŒv•\(ByVal path As String)
    Dim rs As DAO.Recordset
    Dim rowNo As Long, okCnt As Long, ngCnt As Long, skipCnt As Long
    Dim dt As Variant, prodName As String, kubunName As String, opName As String
    Dim colId As Long, cnt As Long, opId As Long, kbId As Long, prId As Long
    Dim biko2 As Variant, biko3 As Variant
    Dim i As Long

    On Error GoTo Fail

    EnsureImportLog
    ExecSQL "DELETE FROM [T_æƒƒO] WHERE [æŒ³]=" & Q(path)

    DropLink
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, LINK_NAME, _
                              path, False, "ƒf[ƒ^!A9:R5000"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM [" & LINK_NAME & "]", dbOpenSnapshot)

    Do While Not rs.EOF
        rowNo = rowNo + 1
        dt = rs.Fields(0).Value                       ' A: “ú•t
        prodName = NzS(rs.Fields(1).Value)            ' B: »•i–¼
        kubunName = NzS(rs.Fields(9).Value)           ' J: ”õl (‹Œ“]‹L–¼)
        opName = NzS(rs.Fields(10).Value)             ' K: ’S“–Ò (©)
        biko2 = rs.Fields(11).Value                   ' L: ”õl2
        biko3 = rs.Fields(12).Value                   ' M: ”õl3

        ' C`H (WŒv—ñ 3`8) ‚Ì‚¤‚¿ 0 ‚Å‚È‚¢—ñ‚ğ’T‚·
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
            skipCnt = skipCnt + 1                     ' Œ©o‚µsESUMIF —p‚Ì‹ós‚È‚Ç
        Else
            opId = FindOperator(opName)
            kbId = FindKubun(kubunName, colId)
            prId = FindProduct(prodName, kbId)

            If opId = 0 Or kbId = 0 Then
                LogImport path, rowNo, CStr(dt), opName, prodName, kubunName, cnt, _
                    IIf(opId = 0, "’S“–Òw" & opName & "x‚ªƒ}ƒXƒ^‚É‚ ‚è‚Ü‚¹‚ñB", "") & _
                    IIf(kbId = 0, "‹æ•ªw" & kubunName & "x(WŒv—ñ" & colId & _
                        ") ‚ªƒ}ƒXƒ^‚É‚ ‚è‚Ü‚¹‚ñB", "")
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

    ' æ‚è‚ñ‚¾“ú•t‚É‚Í“ú•ñ‚Ìƒwƒbƒ_s‚Æo‹Î“o˜^‚ğ—pˆÓ‚µ‚Ä‚¨‚­
    NormalizeAfterImport

    MsgBox "æ‚ªI‚í‚è‚Ü‚µ‚½B" & vbCrLf & vbCrLf & _
           "æ        : " & okCnt & " s" & vbCrLf & _
           "ˆø‚«“–‚Ä¸”s: " & ngCnt & " s" & vbCrLf & _
           "‘ÎÛŠO      : " & skipCnt & " s (Œ©o‚µE‹ós‚È‚Ç)" & vbCrLf & vbCrLf & _
           IIf(ngCnt > 0, "¸”s‚µ‚½s‚Í T_æƒƒO ‚ğŠJ‚¢‚ÄŠm”F‚µ‚Ä‚­‚¾‚³‚¢B", ""), _
           IIf(ngCnt > 0, vbExclamation, vbInformation), APP_NAME
    Exit Sub

Fail:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    DropLink
    MsgBox "æ’†‚ÉƒGƒ‰[‚ª”­¶‚µ‚Ü‚µ‚½B" & vbCrLf & vbCrLf & _
           "ƒGƒ‰[ " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

'------------------------------------------------------------------------------
' ˆø‚«“–‚Ä
'------------------------------------------------------------------------------
' ’S“–Ò‚Í © ‚Åˆø‚­BŒ»s‚ÌWŒv•\‚É‚Í©‚µ‚©“ü‚Á‚Ä‚¢‚È‚¢‚½‚ßB
' ©‚ÅŒ©‚Â‚©‚ç‚È‚¢ê‡‚Íu‹ŒB1’lv(Œ»s Sheet1!B1 ‚Ì’l) ‚Å‚à’T‚·B
Private Function FindOperator(ByVal nm As String) As Long
    Dim v As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function

    v = DLookup("[’S“–ÒID]", "M_’S“–Ò", "[©]=" & Q(nm))
    If IsNull(v) Then v = DLookup("[’S“–ÒID]", "M_’S“–Ò", "[–¼]=" & Q(nm))
    If IsNull(v) Then
        ' îà/‚ ‚Ì‚æ‚¤‚ÈˆÙ‘Ìš‚ä‚ê‚ğ‹zû‚·‚é
        v = DLookup("[’S“–ÒID]", "M_’S“–Ò", "[©] Like " & Q("*" & Mid$(nm, 2) & "*"))
    End If
    FindOperator = Nz(v, 0)
End Function

Private Function FindKubun(ByVal nm As String, ByVal colId As Long) As Long
    Dim v As Variant
    nm = Trim$(nm)

    If Len(nm) > 0 Then
        v = DLookup("[‹æ•ªID]", "M_‹æ•ª", "[‹Œ“]‹L–¼]=" & Q(nm) & " And [WŒv—ñID]=" & colId)
        If IsNull(v) Then v = DLookup("[‹æ•ªID]", "M_‹æ•ª", "[‹æ•ª–¼]=" & Q(nm) & _
                                      " And [WŒv—ñID]=" & colId)
        If IsNull(v) Then v = DLookup("[‹æ•ªID]", "M_‹æ•ª", "[‹Œ“]‹L–¼]=" & Q(nm))
    Else
        ' ”õl‚ª‹ó‚Ìs‚ÍAWŒv—ñ‚¾‚¯‚ÅˆêˆÓ‚ÉŒˆ‚Ü‚é‹æ•ª (\ŠÖŒWE’Š‘IŒ‹‰Ê) ‚ğg‚¤
        v = DLookup("[‹æ•ªID]", "M_‹æ•ª", _
                    "[WŒv—ñID]=" & colId & " And (([‹Œ“]‹L–¼] Is Null) Or ([‹Œ“]‹L–¼]=''))")
    End If
    FindKubun = Nz(v, 0)
End Function

' »•i–¼‚Íã’iƒuƒƒbƒN‚Æy‚»‚Ì‘¼zƒuƒƒbƒN‚Åd•¡‚·‚é‚Ì‚ÅA
' ‚Ü‚¸‹æ•ª‚Æ“¯‚¶ƒuƒƒbƒN‚Å’T‚µA–³‚¯‚ê‚Î–¼‘O‚¾‚¯‚Å’T‚·B
Private Function FindProduct(ByVal nm As String, ByVal kbId As Long) As Long
    Dim v As Variant, blockId As Variant
    nm = Trim$(nm)
    If Len(nm) = 0 Then
        FindProduct = 0
        Exit Function
    End If

    If kbId > 0 Then
        blockId = DLookup("[ƒuƒƒbƒNID]", "M_‹æ•ª", "[‹æ•ªID]=" & kbId)
        If Not IsNull(blockId) Then
            v = DLookup("[»•iID]", "M_»•i", _
                        "[»•i–¼]=" & Q(nm) & " And [ƒuƒƒbƒNID]=" & blockId)
        End If
    End If
    If IsNull(v) Then v = DLookup("[»•iID]", "M_»•i", "[»•i–¼]=" & Q(nm))
    FindProduct = Nz(v, 0)
End Function

'------------------------------------------------------------------------------
' ‘
'------------------------------------------------------------------------------
Private Sub UpsertJuden(ByVal dt As Date, ByVal opId As Long, ByVal kbId As Long, _
                        ByVal prId As Long, ByVal cnt As Long, _
                        ByVal biko2 As Variant, ByVal biko3 As Variant)
    Dim where As String
    where = "[‘ÎÛ“ú]=" & D(dt) & " And [’S“–ÒID]=" & opId & _
            " And [‹æ•ªID]=" & kbId & " And [»•iID]=" & prId

    If DCount("*", "T_ó“d", where) > 0 Then
        ExecSQL "UPDATE [T_ó“d] SET [Œ”]=" & cnt & _
                ",[”õl2]=" & Q(NzS(biko2)) & ",[”õl3]=" & Q(NzS(biko3)) & _
                ",[XV“ú]=Now() WHERE " & where
    Else
        ExecSQL "INSERT INTO [T_ó“d] " & _
                "([‘ÎÛ“ú],[’S“–ÒID],[‹æ•ªID],[»•iID],[Œ”],[”õl2],[”õl3]," & _
                " [“o˜^“ú],[XV“ú],[“o˜^Ò]) VALUES (" & _
                D(dt) & "," & opId & "," & kbId & "," & prId & "," & cnt & "," & _
                Q(NzS(biko2)) & "," & Q(NzS(biko3)) & ",Now(),Now()," & _
                Q("æ:" & CurrentUserName) & ")"
    End If
End Sub

' æ‚è‚ñ‚¾“ú•t‚²‚Æ‚ÉA“ú•ñƒwƒbƒ_‚Æo‹Î“o˜^‚ğ•â‚¤B
Private Sub NormalizeAfterImport()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT DISTINCT [‘ÎÛ“ú] FROM [T_ó“d] " & _
        "WHERE [‘ÎÛ“ú] NOT IN (SELECT [‘ÎÛ“ú] FROM [T_“ú•ñ])")
    Do While Not rs.EOF
        Ensure“ú•ñ CDate(rs![‘ÎÛ“ú])
        rs.MoveNext
    Loop
    rs.Close

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT DISTINCT J.[‘ÎÛ“ú], J.[’S“–ÒID] FROM [T_ó“d] AS J " & _
        "WHERE NOT EXISTS (SELECT 1 FROM [T_o‹Î] AS A " & _
        " WHERE A.[‘ÎÛ“ú]=J.[‘ÎÛ“ú] AND A.[’S“–ÒID]=J.[’S“–ÒID])")
    Do While Not rs.EOF
        Ensureo‹Î CDate(rs![‘ÎÛ“ú]), CLng(rs![’S“–ÒID])
        rs.MoveNext
    Loop
    rs.Close
End Sub

'------------------------------------------------------------------------------
' æƒƒO
'------------------------------------------------------------------------------
Public Sub EnsureImportLog()
    If TableExists("T_æƒƒO") Then Exit Sub
    ExecDDL "CREATE TABLE [T_æƒƒO] (" & _
        "[ƒƒOID] COUNTER NOT NULL CONSTRAINT [PK_æƒƒO] PRIMARY KEY," & _
        "[æ“ú] DATETIME," & _
        "[æŒ³] TEXT(255)," & _
        "[s”Ô†] LONG," & _
        "[“ú•t] TEXT(50)," & _
        "[’S“–Ò] TEXT(100)," & _
        "[»•i–¼] TEXT(100)," & _
        "[‹æ•ª–¼] TEXT(100)," & _
        "[Œ”] LONG," & _
        "[——R] MEMO)"
End Sub

Private Sub LogImport(ByVal src As String, ByVal rowNo As Long, ByVal dt As String, _
                      ByVal opName As String, ByVal prodName As String, _
                      ByVal kubunName As String, ByVal cnt As Long, ByVal reason As String)
    ExecSQL "INSERT INTO [T_æƒƒO] " & _
            "([æ“ú],[æŒ³],[s”Ô†],[“ú•t],[’S“–Ò],[»•i–¼],[‹æ•ª–¼],[Œ”],[——R]) " & _
            "VALUES (Now()," & Q(src) & "," & rowNo & "," & Q(dt) & "," & Q(opName) & "," & _
            Q(prodName) & "," & Q(kubunName) & "," & cnt & "," & Q(reason) & ")"
End Sub

'------------------------------------------------------------------------------
' ¬•¨
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
        .Title = "æ‚è‚ŞWŒv•\‚ğ‘I‚ñ‚Å‚­‚¾‚³‚¢"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel ƒuƒbƒN", "*.xlsm; *.xlsx; *.xls"
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
    Exit Function
Manual:
    PickExcelFile = InputBox("æ‚è‚ŞWŒv•\‚Ìƒtƒ‹ƒpƒX‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B", APP_NAME)
End Function
