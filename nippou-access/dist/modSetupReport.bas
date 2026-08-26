Attribute VB_Name = "modSetupReport"
Option Compare Database
Option Explicit

'==============================================================================
' 帳票 (レポート) の自動生成
'
' R_日報 は現行の「日報集計印刷用フォーム.xlsm ＞ 印刷用シート」の体裁を
' そのまま再現する。ただし値の出どころは T_受電 ただ 1 つに統一してある。
'
' 現行 印刷用シート との対応:
'   W2         → 令和表記の日付          (WarekiLong)
'   C10        → 出勤者数                (Q_日報_ヘッダ.出勤者数)
'   F10:O16    → 出勤者の氏名            (AttendeeName)
'   C16        → 回線数                  (T_日報.回線数)
'   B18/E18    → 合計 / うち職員         (Q_日報_受電)
'   F18～S18   → 申込・抽選・払込用紙・商品発送・その他
'   U19/U20    → 内 交換 / 内 返金
'   H39～T43   → 電話応対以外の業務 ①～⑬ (TaskValue)
'==============================================================================

Private Const W As Long = 10500          ' 印字領域の幅 (A4 縦・左右余白 1.2cm)
Private Const LN As Long = 300           ' 標準の行高

Public Sub Setup_Reports()
    Echo_ "帳票を作成しています..."
    Build_R_日報
    Build_R_集計表
    Build_R_受電明細
    Echo_ "  帳票作成 完了"
End Sub

'==============================================================================
' 日報 (現行レイアウト再現)
'==============================================================================
Private Sub Build_R_日報()
    Dim rpt As Access.Report, c As Access.Control
    Dim y As Long, i As Long, x As Long
    Dim caps As Variant, flds As Variant, colW As Long

    Set rpt = CreateReport()
    rpt.RecordSource = "Q_日報_ヘッダ"

    '--- ヘッダ行 ---
    y = 0
    RLabel rpt, "【課内限り】", 0, y, 2400, 280, 9, False, 1
    RExpr rpt, "txt日付", "=WarekiLong([対象日])", 5800, y, W - 5800, 300, 11, False, 3

    '--- 表題 ---
    y = 420
    RLabel rpt, "電話応対報告書日報集計表", 0, y, W, 560, 16, True, 2

    '--- 出勤者 ---
    y = 1080
    RLabel rpt, "出勤者", 0, y, 1300, LN, 10, True, 1
    RExpr rpt, "txt出勤者数", "=[出勤者数]", 1300, y, 700, LN, 11, True, 2, True
    RLabel rpt, "名", 2050, y, 500, LN, 10, False, 1

    RLabel rpt, "氏名", 3000, y, 3200, LN, 10, True, 2, True
    RLabel rpt, "氏名", 6200, y, 3200, LN, 10, True, 2, True
    RLabel rpt, "備考", 9400, y, W - 9400, LN, 10, True, 2, True

    ' 氏名は 2 列 × 7 行 = 14 名まで (現行 F10:O16 と同じ枠)
    For i = 1 To 7
        RExpr rpt, "txt氏名" & i, "=AttendeeName([対象日]," & i & ")", _
              3000, y + i * LN, 3200, LN, 10, False, 1, True
        RExpr rpt, "txt氏名" & (i + 7), "=AttendeeName([対象日]," & (i + 7) & ")", _
              6200, y + i * LN, 3200, LN, 10, False, 1, True
        RLabel rpt, "", 9400, y + i * LN, W - 9400, LN, 10, False, 1, True
    Next i
    y = y + 8 * LN + 80

    '--- 回線数 ---
    RLabel rpt, "（", 0, y, 300, LN, 10, False, 1
    RExpr rpt, "txt回線数", "=[回線数]", 300, y, 700, LN, 11, True, 2, True
    RLabel rpt, "回線 ）", 1030, y, 1600, LN, 10, False, 1
    y = y + LN + 120

    '--- 問合せ件数 ---
    RLabel rpt, "問合せ件数", 0, y, 1500, LN, 10, True, 1
    RLabel rpt, "（　）内は職員受電数", 0, y + LN, 1500, LN, 8, False, 1

    RExpr rpt, "txt合計", "=[合計]", 1500, y, 900, LN + 40, 13, True, 2, True
    RLabel rpt, "件", 2400, y, 400, LN, 10, False, 1
    RExpr rpt, "txt合計職員", "='(' & [合計_職員] & ')'", 1500, y + LN + 60, 900, LN, 10, False, 2

    caps = Array("申込", "抽選", "払込用紙", "商品発送", "その他")
    flds = Array("申込", "抽選", "払込用紙", "商品発送", "その他")
    colW = (W - 2900) \ 5
    For i = 0 To 4
        x = 2900 + i * colW
        RLabel rpt, CStr(caps(i)), x, y, colW, LN, 10, True, 2, True
        RExpr rpt, "txt" & flds(i), "=[" & flds(i) & "]", _
              x, y + LN, colW, LN + 40, 12, False, 2, True
        RExpr rpt, "txt" & flds(i) & "_職", "='(' & [" & flds(i) & "_職員] & ')'", _
              x, y + 2 * LN + 40, colW, LN, 9, False, 2, True
    Next i
    y = y + 3 * LN + 100

    '--- 内訳 ---
    RLabel rpt, "内　　交換", 6400, y, 1600, LN, 10, True, 3
    RExpr rpt, "txt内交換", "=[内交換]", 8100, y, 800, LN, 10, False, 3, True
    RLabel rpt, "件", 8950, y, 400, LN, 10, False, 1
    y = y + LN
    RLabel rpt, "内　　返金", 6400, y, 1600, LN, 10, True, 3
    RExpr rpt, "txt内返金", "=[内返金]", 8100, y, 800, LN, 10, False, 3, True
    RLabel rpt, "件", 8950, y, 400, LN, 10, False, 1
    y = y + LN + 160

    '--- 記述欄 ---
    y = AddMemoBlock(rpt, y, "特記事項（報告書の電話件数だけでは伝わり難い事項など）", "特記事項", 1500)
    y = AddMemoBlock(rpt, y, "職員に代わった案件（概要）", "職員代替案件", 1500)
    y = AddMemoBlock(rpt, y, "要望（お客様からの問い合わせを減らすための改善提案等）", "要望", 1500)

    '--- 電話対応以外の業務 ---
    RLabel rpt, "電話対応以外の業務（データ入力作業等）", 0, y, W, LN, 10, True, 1
    y = y + LN + 40

    Dim rs As DAO.Recordset, n As Long, col As Long, row As Long
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT [業務項目ID],[帳票表示名] FROM [M_業務項目] WHERE [有効]=True ORDER BY [表示順]")
    Do While Not rs.EOF
        col = IIf(n < 8, 0, 1)                       ' 左 8 件 / 右 5 件 (現行と同じ割り)
        row = IIf(n < 8, n, n - 8)
        x = col * 5300
        RLabel rpt, CStr(rs![帳票表示名]), x, y + row * LN, 3500, LN, 9, False, 1
        RExpr rpt, "txt業務" & rs![業務項目ID], _
              "=TaskValue([対象日]," & rs![業務項目ID] & ")", _
              x + 3500, y + row * LN, 800, LN, 9, False, 3, True
        RLabel rpt, "件", x + 4350, y + row * LN, 400, LN, 9, False, 1
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    y = y + 8 * LN + 200

    RExpr rpt, "txt印字", _
          "='印刷: ' & Format(Now(),'yyyy/mm/dd hh:nn') & '　　状態: ' & Nz([状態],'')", _
          0, y, W, 260, 8, False, 1
    y = y + 300

    rpt.caption = "電話応対報告書日報集計表"
    rpt.Section(acDetail).Height = y
    rpt.Width = W
    SetupPage rpt, 1
    EndReport rpt, "R_日報"
End Sub

' ラベル + 枠付きメモ欄をひとかたまりで置く
Private Function AddMemoBlock(ByRef rpt As Access.Report, ByVal y As Long, _
                              ByVal caption As String, ByVal fld As String, _
                              ByVal h As Long) As Long
    RLabel rpt, caption, 0, y, W, 280, 10, True, 1
    Dim c As Access.Control
    Set c = CreateReportControl(rpt.Name, acTextBox, acDetail, "", fld, 0, y + 280, W, h)
    c.Name = "txt" & fld
    c.fontSize = 9
    c.BorderStyle = 1
    c.CanGrow = False
    c.CanShrink = False
    AddMemoBlock = y + 280 + h + 120
End Function

'==============================================================================
' 集計表
'==============================================================================
Private Sub Build_R_集計表()
    Dim rpt As Access.Report, c As Access.Control
    Dim cols As Variant, i As Long, x As Long, cw As Long
    Set rpt = CreateReport()
    rpt.RecordSource = "SELECT * FROM [Q_週次集計指定] ORDER BY [対象日];"

    cols = Array("申込方法", "抽選結果", "納付書発送", "商品発送", "その他", "商品交換", "計")
    cw = 1150

    On Error Resume Next
    rpt.Section(acPageHeader).Height = 1200
    rpt.Section(acPageFooter).Height = 400
    rpt.Section(acDetail).Height = LN + 20
    On Error GoTo 0

    RLabelS rpt, acPageHeader, "集計表", 0, 0, 4000, 480, 15, True, 1
    RExprS rpt, acPageHeader, "txt期間", _
           "='期間: ' & Format([対象日],'yyyy/mm/dd')", 4200, 100, 4000, 300, 9, False, 1

    RLabelS rpt, acPageHeader, "日付", 0, 800, 1600, LN, 10, True, 2, True
    x = 1600
    For i = 0 To UBound(cols)
        RLabelS rpt, acPageHeader, CStr(cols(i)), x, 800, cw, LN, 10, True, 2, True
        x = x + cw
    Next i

    Set c = CreateReportControl(rpt.Name, acTextBox, acDetail, "", "対象日", 0, 10, 1600, LN)
    c.Name = "対象日"
    c.Format = "yyyy/mm/dd(aaa)"
    c.BorderStyle = 1
    x = 1600
    For i = 0 To UBound(cols)
        Set c = CreateReportControl(rpt.Name, acTextBox, acDetail, "", CStr(cols(i)), x, 10, cw, LN)
        c.Name = CStr(cols(i))
        c.TextAlign = 3
        c.BorderStyle = 1
        c.FontBold = (cols(i) = "計")
        x = x + cw
    Next i

    ' 合計行
    On Error Resume Next
    rpt.Section(acFooter).Height = LN + 100
    On Error GoTo 0
    RLabelS rpt, acFooter, "計", 0, 20, 1600, LN, 10, True, 2, True
    x = 1600
    For i = 0 To UBound(cols)
        RExprS rpt, acFooter, "sum" & i, "=Sum([" & cols(i) & "])", x, 20, cw, LN, 10, True, 3, True
        x = x + cw
    Next i

    rpt.caption = "集計表"
    rpt.Width = 1600 + cw * (UBound(cols) + 1)
    SetupPage rpt, 1
    EndReport rpt, "R_集計表"
End Sub

'==============================================================================
' 受付明細
'==============================================================================
Private Sub Build_R_受電明細()
    Dim rpt As Access.Report, c As Access.Control
    Dim flds As Variant, ws As Variant, i As Long, x As Long
    Set rpt = CreateReport()
    rpt.RecordSource = "SELECT * FROM [Q_受電明細] ORDER BY [対象日],[氏名],[区分ID];"

    flds = Array("対象日", "氏名", "製品名", "区分名", "集計列名", "計", "備考2")
    ws = Array(1300, 1500, 2800, 3200, 1300, 600, 3300)

    On Error Resume Next
    rpt.Section(acPageHeader).Height = 900
    rpt.Section(acDetail).Height = 280
    On Error GoTo 0

    RLabelS rpt, acPageHeader, "受付明細", 0, 0, 4000, 420, 14, True, 1
    x = 0
    For i = 0 To UBound(flds)
        RLabelS rpt, acPageHeader, CStr(flds(i)), x, 540, CLng(ws(i)), 280, 9, True, 2, True
        Set c = CreateReportControl(rpt.Name, acTextBox, acDetail, "", CStr(flds(i)), x, 0, CLng(ws(i)), 260)
        c.Name = CStr(flds(i))
        c.fontSize = 8
        If flds(i) = "対象日" Then c.Format = "yyyy/mm/dd"
        If flds(i) = "計" Then c.TextAlign = 3
        x = x + CLng(ws(i))
    Next i

    rpt.caption = "受付明細"
    rpt.Width = x
    SetupPage rpt, 2                       ' 横向き
    EndReport rpt, "R_受電明細"
End Sub

'==============================================================================
' 生成ヘルパ
'==============================================================================
Private Sub EndReport(ByRef rpt As Access.Report, ByVal finalName As String)
    Dim tmp As String
    tmp = rpt.Name
    DoCmd.Close acReport, tmp, acSaveYes
    DeleteObjectIfExists acReport, finalName
    DoCmd.Rename finalName, acReport, tmp
    Echo_ "  帳票 " & finalName
End Sub

' orientation: 1=縦 2=横
Private Sub SetupPage(ByRef rpt As Access.Report, ByVal orientation As Integer)
    On Error Resume Next
    With rpt.Printer
        .PaperSize = 9                    ' A4
        .orientation = orientation
        .TopMargin = 1000
        .BottomMargin = 800
        .LeftMargin = 680
        .RightMargin = 680
    End With
    Err.Clear
End Sub

Private Sub RLabel(ByRef rpt As Access.Report, ByVal caption As String, _
                   ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long, _
                   Optional ByVal fs As Integer = 9, Optional ByVal bold As Boolean = False, _
                   Optional ByVal align As Integer = 1, Optional ByVal boxed As Boolean = False)
    RLabelS rpt, acDetail, caption, l, t, w, h, fs, bold, align, boxed
End Sub

Private Sub RLabelS(ByRef rpt As Access.Report, ByVal sect As Integer, ByVal caption As String, _
                    ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long, _
                    Optional ByVal fs As Integer = 9, Optional ByVal bold As Boolean = False, _
                    Optional ByVal align As Integer = 1, Optional ByVal boxed As Boolean = False)
    Dim c As Access.Control
    ' Caption が空だと Access がラベルを作れないので、枠だけ欲しい時は空白 1 文字を入れる
    Set c = CreateReportControl(rpt.Name, acLabel, sect, "", "", l, t, w, h)
    c.caption = IIf(Len(caption) = 0, " ", caption)
    c.fontSize = fs
    c.FontBold = bold
    c.TextAlign = align
    If boxed Then c.BorderStyle = 1
End Sub

Private Sub RExpr(ByRef rpt As Access.Report, ByVal name As String, ByVal expr As String, _
                  ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long, _
                  Optional ByVal fs As Integer = 9, Optional ByVal bold As Boolean = False, _
                  Optional ByVal align As Integer = 1, Optional ByVal boxed As Boolean = False)
    RExprS rpt, acDetail, name, expr, l, t, w, h, fs, bold, align, boxed
End Sub

Private Sub RExprS(ByRef rpt As Access.Report, ByVal sect As Integer, ByVal name As String, _
                   ByVal expr As String, ByVal l As Long, ByVal t As Long, _
                   ByVal w As Long, ByVal h As Long, _
                   Optional ByVal fs As Integer = 9, Optional ByVal bold As Boolean = False, _
                   Optional ByVal align As Integer = 1, Optional ByVal boxed As Boolean = False)
    Dim c As Access.Control
    Set c = CreateReportControl(rpt.Name, acTextBox, sect, "", "", l, t, w, h)
    c.Name = name
    c.ControlSource = expr
    c.fontSize = fs
    c.FontBold = bold
    c.TextAlign = align
    If boxed Then c.BorderStyle = 1
End Sub
