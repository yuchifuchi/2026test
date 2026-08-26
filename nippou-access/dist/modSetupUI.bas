Attribute VB_Name = "modSetupUI"
Option Compare Database
Option Explicit

'==============================================================================
' 画面 (フォーム) の自動生成
'
' フォームはコードから組み立てる。理由:
'   ・.accdb をバイナリで配布せずに済み、Git で差分が追える
'   ・作り直しても毎回同じ画面になる (属人化しない)
'
' イベントはすべて「=関数名()」形式の式で書き、フォームモジュールを持たせない。
' こうすると「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」の
' 設定を有効にしなくても構築できる。処理の実体は modApp にある。
'==============================================================================

' レイアウト定数 (単位: twip / 1cm = 567twip)
Private Const ROW_H As Long = 320
Private Const GAP As Long = 80
Private Const LBL_W As Long = 1800
Private Const FORM_W As Long = 13000

Public Sub Setup_Forms()
    Echo_ "画面を作成しています..."

    ' 親より先に子 (サブフォーム) を作る
    Build_F_日次入力_明細
    Build_F_日報_出勤
    Build_F_業務実績
    Build_F_受電明細

    Build_F_日次入力
    Build_F_日報
    Build_F_集計表
    Build_F_未入力チェック

    Build_F_マスタ_担当者
    Build_F_マスタ_製品
    Build_F_マスタ_区分
    Build_F_マスタ_業務項目

    Build_F_メニュー

    Echo_ "  画面作成 完了"
End Sub

'==============================================================================
' メニュー
'==============================================================================
Private Sub Build_F_メニュー()
    Dim frm As Access.Form, t As Long
    Set frm = CreateForm()

    AddLabel frm, acDetail, APP_NAME, 240, 240, 7000, 500, 16, True
    AddLabel frm, acDetail, "毎日の受付記録を入力し、日報と集計表を同じデータから出力します。", _
             240, 780, 9000, 300, 9, False

    t = 1300
    AddButton frm, 240, t, 4200, 640, "① 本日の受付を入力する", "日次入力", 12
    AddButton frm, 4560, t, 4200, 640, "② 日報を作成・確認する", "日報", 12

    t = t + 760
    AddButton frm, 240, t, 4200, 560, "③ 日報を印刷する", "日報印刷", 11
    AddButton frm, 4560, t, 4200, 560, "④ 集計表を見る／印刷する", "集計表", 11

    t = t + 700
    AddLabel frm, acDetail, "― 確認・点検 ―", 240, t, 4000, 300, 10, True
    t = t + 340
    AddButton frm, 240, t, 4200, 500, "入力もれチェック", "未入力チェック", 10
    AddButton frm, 4560, t, 4200, 500, "受付明細を一覧で見る", "受電明細", 10

    t = t + 640
    AddLabel frm, acDetail, "― 保守 (担当者の入退職・製品の追加はここ) ―", 240, t, 6000, 300, 10, True
    t = t + 340
    AddButton frm, 240, t, 2700, 500, "担当者マスタ", "マスタ_担当者", 10
    AddButton frm, 3000, t, 2700, 500, "製品マスタ", "マスタ_製品", 10
    AddButton frm, 5760, t, 3000, 500, "区分マスタ", "マスタ_区分", 10
    t = t + 560
    AddButton frm, 240, t, 2700, 500, "業務項目マスタ", "マスタ_業務項目", 10
    AddButton frm, 3000, t, 2700, 500, "Excel から取込", "取込", 10
    AddButton frm, 5760, t, 3000, 500, "終了", "終了", 10

    t = t + 700
    AddLabel frm, acDetail, "Version " & APP_VERSION, 240, t, 3000, 260, 8, False

    frm.Caption = APP_NAME
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.AutoCenter = True
    frm.Section(acDetail).Height = t + 400
    frm.Width = 9200

    EndForm frm, "F_メニュー"
End Sub

'==============================================================================
' 日次入力  (対象日 + 担当者 を選び、明細をデータシートで入力する)
'==============================================================================
Private Sub Build_F_日次入力()
    Dim frm As Access.Form, c As Access.Control, t As Long
    Set frm = CreateForm()

    AddLabel frm, acDetail, "本日の受付記録を入力", 240, 200, 5000, 420, 14, True

    t = 700
    AddLabel frm, acDetail, "対象日", 240, t + 40, LBL_W, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 240 + LBL_W, t, 2000, ROW_H)
    c.Name = "対象日"
    c.FontSize = 11
    c.Format = "yyyy/mm/dd"
    c.DefaultValue = "=Date()"
    c.AfterUpdate = "=App_Click(""日次入力_再表示"")"
    c.StatusBarText = "入力する日付。既定は本日。"

    AddLabel frm, acDetail, "担当者", 4400, t + 40, 1200, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "", 5600, t, 3200, ROW_H)
    c.Name = "担当者ID"
    c.FontSize = 11
    SetRowSource c, "Q_選択_担当者", 5, "0;0;3000;900;0", 3900
    c.AfterUpdate = "=App_Click(""日次入力_再表示"")"
    c.StatusBarText = "在籍中の担当者だけが出ます。退職者は担当者マスタで在籍終了日を入れてください。"

    AddButton frm, 9100, t - 20, 1600, ROW_H + 40, "表示", "日次入力_再表示", 10

    t = t + ROW_H + GAP + 60
    AddLabel frm, acDetail, _
        "区分を選び、件数を入れてください。製品を伴わない区分は製品欄を空のままで構いません。", _
        240, t, 11000, 280, 9, False

    t = t + 320
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", 240, t, 12400, 5200)
    c.Name = "sub明細"
    c.SourceObject = "F_日次入力_明細"
    c.LinkMasterFields = "対象日;担当者ID"
    c.LinkChildFields = "対象日;担当者ID"

    t = t + 5200 + GAP
    AddLabel frm, acDetail, "この担当者・この日の合計", 240, t + 40, 3400, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 3700, t, 1400, ROW_H)
    c.Name = "txt合計"
    c.FontSize = 12
    c.FontBold = True
    c.TextAlign = 3
    c.Locked = True
    c.ControlSource = "=Nz([sub明細].[Form]![txt件数計],0)"

    AddLabel frm, acDetail, "件", 5150, t + 40, 500, ROW_H, 11, False

    AddButton frm, 8000, t - 20, 2200, ROW_H + 60, "電話応対以外の業務", "業務実績", 10
    AddButton frm, 10400, t - 20, 2240, ROW_H + 60, "閉じる", "閉じる", 10

    frm.Caption = "日次入力"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AutoCenter = True
    frm.Section(acDetail).Height = t + ROW_H + 300
    frm.Width = FORM_W

    EndForm frm, "F_日次入力"
End Sub

Private Sub Build_F_日次入力_明細()
    Dim frm As Access.Form, c As Access.Control
    Set frm = CreateForm()
    frm.RecordSource = "T_受電"

    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "区分ID", 0, 0, 4600, ROW_H)
    c.Name = "区分ID"
    SetRowSource c, "Q_選択_区分", 5, "0;0;0;0;4400", 4600
    c.StatusBarText = "問合せの区分。ブロック / 区分名 の形で表示されます。"

    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "製品ID", 4600, 0, 4600, ROW_H)
    c.Name = "製品ID"
    SetRowSource c, "Q_選択_製品", 5, "0;0;0;0;4400", 4600
    c.DefaultValue = "0"
    c.StatusBarText = "製品を伴わない区分は「（製品指定なし）」のままで構いません。"

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "件数", 9200, 0, 900, ROW_H)
    c.Name = "件数"
    c.TextAlign = 3
    c.DefaultValue = "1"

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "備考2", 10100, 0, 2600, ROW_H)
    c.Name = "備考2"

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "備考3", 12700, 0, 2600, ROW_H)
    c.Name = "備考3"

    ' 登録日時・登録者は既定値で入れる。
    ' サブフォームのイベントから Screen.ActiveForm を見ると親フォームが返るので、
    ' イベント経由では設定できない。
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "登録日時", 15300, 0, 100, ROW_H)
    c.Name = "登録日時"
    c.DefaultValue = "=Now()"
    c.Visible = False
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "登録者", 15400, 0, 100, ROW_H)
    c.Name = "登録者"
    c.DefaultValue = "=CurrentUserName()"
    c.Visible = False

    ' 日付・担当者はリンク用に置くだけ (データシートでは幅 0 で隠す)
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "対象日", 15500, 0, 100, ROW_H)
    c.Name = "対象日"
    c.Visible = False
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "担当者ID", 15600, 0, 100, ROW_H)
    c.Name = "担当者ID"
    c.Visible = False

    ' フッタに合計を出す。親フォームがここを参照する。
    On Error Resume Next
    frm.Section(acFooter).Height = ROW_H + 100
    Set c = CreateControl(frm.Name, acTextBox, acFooter, "", "", 9200, 40, 900, ROW_H)
    c.Name = "txt件数計"
    c.ControlSource = "=Sum([件数])"
    c.TextAlign = 3
    On Error GoTo 0

    frm.DefaultView = 2          ' データシート
    frm.ViewsAllowed = 2
    frm.Caption = "受付明細"
    frm.OrderBy = "区分ID"
    frm.OrderByOn = True

    EndForm frm, "F_日次入力_明細"
End Sub

'==============================================================================
' 電話応対以外の業務
'==============================================================================
Private Sub Build_F_業務実績()
    Dim frm As Access.Form, c As Access.Control
    Set frm = CreateForm()
    frm.RecordSource = "T_業務実績"

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "対象日", 0, 0, 1600, ROW_H)
    c.Name = "対象日"
    c.Format = "yyyy/mm/dd"

    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "担当者ID", 1600, 0, 2600, ROW_H)
    c.Name = "担当者ID"
    SetRowSource c, "Q_選択_担当者", 5, "0;800;2400;0;0", 3400

    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "業務項目ID", 4200, 0, 4200, ROW_H)
    c.Name = "業務項目ID"
    SetRowSource c, "M_業務項目", 4, "0;500;3600;0", 4200

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "件数", 8400, 0, 900, ROW_H)
    c.Name = "件数"
    c.TextAlign = 3

    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "電話応対以外の業務"
    EndForm frm, "F_業務実績"
End Sub

'==============================================================================
' 日報 (ヘッダ + 出勤者 + 集計値の確認)
'==============================================================================
Private Sub Build_F_日報()
    Dim frm As Access.Form, c As Access.Control, t As Long, i As Long
    Dim cols As Variant, caps As Variant
    Set frm = CreateForm()
    frm.RecordSource = "T_日報"

    AddLabel frm, acDetail, "電話応対報告書 日報集計表", 240, 200, 6000, 440, 14, True

    t = 700
    AddLabel frm, acDetail, "対象日", 240, t + 40, LBL_W, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "対象日", 240 + LBL_W, t, 2000, ROW_H)
    c.Name = "対象日"
    c.FontSize = 11
    c.Format = "yyyy/mm/dd"

    AddLabel frm, acDetail, "回線数", 4400, t + 40, 1200, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "回線数", 5600, t, 900, ROW_H)
    c.Name = "回線数"
    c.TextAlign = 3

    AddLabel frm, acDetail, "状態", 7000, t + 40, 900, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "状態", 7900, t, 1400, ROW_H)
    c.Name = "状態"
    c.Locked = True

    AddButton frm, 9600, t - 20, 1500, ROW_H + 40, "確定", "日報確定", 10
    AddButton frm, 11200, t - 20, 1500, ROW_H + 40, "印刷", "日報印刷", 10

    '--- 集計値 (T_受電 から自動計算。手入力欄ではない) ---
    t = t + ROW_H + GAP + 60
    AddLabel frm, acDetail, "問合せ件数 (受付明細から自動集計。手入力はできません)", _
             240, t, 8000, 300, 10, True
    t = t + 340

    cols = Array("申込", "抽選", "払込用紙", "商品発送", "その他", "合計")
    For i = 0 To 5
        AddLabel frm, acDetail, cols(i), 240 + i * 1700, t, 1600, 280, 9, True, 2
        Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", _
                              240 + i * 1700, t + 300, 1600, 380)
        c.Name = "txt" & cols(i)
        c.ControlSource = "=DailyValue([対象日]," & Q(CStr(cols(i))) & ")"
        c.TextAlign = 2
        c.FontSize = 12
        c.FontBold = (i = 5)
        c.Locked = True
        Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", _
                              240 + i * 1700, t + 690, 1600, 300)
        c.Name = "txt" & cols(i) & "_職員"
        c.ControlSource = "='(' & DailyValue([対象日]," & Q(CStr(cols(i)) & "_職員") & ") & ')'"
        c.TextAlign = 2
        c.Locked = True
    Next i
    AddLabel frm, acDetail, "（　）内は職員受電数", 240 + 6 * 1700, t + 690, 2400, 300, 9, False

    t = t + 1050
    AddLabel frm, acDetail, "内 交換", 240, t + 40, 1200, ROW_H, 10, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1440, t, 900, ROW_H)
    c.Name = "txt内交換"
    c.ControlSource = "=DailyValue([対象日],""内交換"")"
    c.TextAlign = 3
    c.Locked = True
    AddLabel frm, acDetail, "内 返金", 2600, t + 40, 1200, ROW_H, 10, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 3800, t, 900, ROW_H)
    c.Name = "txt内返金"
    c.ControlSource = "=DailyValue([対象日],""内返金"")"
    c.TextAlign = 3
    c.Locked = True

    '--- 出勤者 ---
    t = t + ROW_H + GAP + 40
    AddLabel frm, acDetail, "出勤者", 240, t, 2000, 300, 10, True
    AddButton frm, 2400, t - 40, 2600, 360, "在籍者を一括登録", "出勤一括登録", 9
    t = t + 340
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", 240, t, 6000, 2400)
    c.Name = "sub出勤"
    c.SourceObject = "F_日報_出勤"
    c.LinkMasterFields = "対象日"
    c.LinkChildFields = "対象日"

    '--- 記述欄 ---
    caps = Array("特記事項", "職員代替案件", "要望")
    Dim tops As Variant
    tops = Array(t, t + 820, t + 1640)
    For i = 0 To 2
        AddLabel frm, acDetail, caps(i), 6500, tops(i), 2400, 280, 10, True
        Set c = CreateControl(frm.Name, acTextBox, acDetail, "", CStr(caps(i)), _
                              6500, tops(i) + 300, 6200, 480)
        c.Name = "txt" & caps(i)
        c.ScrollBars = 2
    Next i

    t = t + 2400 + GAP
    AddButton frm, 11200, t, 1500, ROW_H + 40, "閉じる", "閉じる", 10

    frm.Caption = "日報"
    frm.RecordSelectors = False
    frm.NavigationButtons = True
    frm.AutoCenter = True
    frm.Section(acDetail).Height = t + ROW_H + 300
    frm.Width = FORM_W

    EndForm frm, "F_日報"
End Sub

Private Sub Build_F_日報_出勤()
    Dim frm As Access.Form, c As Access.Control
    Set frm = CreateForm()
    frm.RecordSource = "T_出勤"

    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "担当者ID", 0, 0, 2800, ROW_H)
    c.Name = "担当者ID"
    SetRowSource c, "Q_選択_担当者", 5, "0;800;2400;0;0", 3400

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "勤務時間", 2800, 0, 2200, ROW_H)
    c.Name = "勤務時間"
    c.DefaultValue = """9：00 ～ 17：00"""

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "備考", 5000, 0, 2600, ROW_H)
    c.Name = "備考"

    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "対象日", 7600, 0, 100, ROW_H)
    c.Name = "対象日"
    c.Visible = False

    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "出勤者"
    EndForm frm, "F_日報_出勤"
End Sub

'==============================================================================
' 集計表 (週次)
'==============================================================================
Private Sub Build_F_集計表()
    Dim frm As Access.Form, c As Access.Control, t As Long
    Set frm = CreateForm()

    AddLabel frm, acDetail, "集計表", 240, 200, 4000, 440, 14, True
    AddLabel frm, acDetail, _
        "期間を指定して表示します。日報と同じ受付明細から集計するので、数値は必ず一致します。", _
        240, 660, 10000, 280, 9, False

    t = 1000
    AddLabel frm, acDetail, "開始日", 240, t + 40, 1000, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1240, t, 1800, ROW_H)
    c.Name = "開始日"
    c.Format = "yyyy/mm/dd"
    c.DefaultValue = "=Date()-Weekday(Date(),2)+1"

    AddLabel frm, acDetail, "終了日", 3200, t + 40, 1000, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 4200, t, 1800, ROW_H)
    c.Name = "終了日"
    c.Format = "yyyy/mm/dd"
    c.DefaultValue = "=Date()-Weekday(Date(),2)+5"

    AddButton frm, 6200, t - 20, 1400, ROW_H + 40, "表示", "集計表_表示", 10
    AddButton frm, 7700, t - 20, 1400, ROW_H + 40, "今週", "集計表_今週", 10
    AddButton frm, 9200, t - 20, 1400, ROW_H + 40, "先週", "集計表_先週", 10
    AddButton frm, 10700, t - 20, 1600, ROW_H + 40, "印刷", "集計表_印刷", 10

    t = t + ROW_H + GAP + 60
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", 240, t, 12400, 2200)
    c.Name = "sub集計"
    c.SourceObject = "Table.Q_週次集計指定"

    t = t + 2200 + GAP
    AddLabel frm, acDetail, "明細", 240, t, 2000, 300, 10, True
    t = t + 320
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", 240, t, 12400, 3600)
    c.Name = "sub明細"
    c.SourceObject = "F_受電明細"

    t = t + 3600 + GAP
    AddButton frm, 11200, t, 1440, ROW_H + 40, "閉じる", "閉じる", 10

    frm.Caption = "集計表"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AutoCenter = True
    frm.OnLoad = "=App_Click(""集計表_表示"")"
    frm.Section(acDetail).Height = t + ROW_H + 300
    frm.Width = FORM_W

    EndForm frm, "F_集計表"
End Sub

Private Sub Build_F_受電明細()
    Dim frm As Access.Form, c As Access.Control
    Dim flds As Variant, w As Variant, i As Long
    Set frm = CreateForm()
    frm.RecordSource = "Q_受電明細"

    flds = Array("対象日", "氏名", "ブロック名", "製品名", "区分名", "集計列名", _
                 "計", "備考2", "備考3")
    w = Array(1500, 1800, 2000, 3000, 3000, 1600, 700, 2200, 2200)
    Dim x As Long
    x = 0
    For i = 0 To UBound(flds)
        Set c = CreateControl(frm.Name, acTextBox, acDetail, "", CStr(flds(i)), x, 0, CLng(w(i)), ROW_H)
        c.Name = CStr(flds(i))
        c.Locked = True
        x = x + CLng(w(i))
    Next i

    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = False
    frm.Caption = "受付明細"
    EndForm frm, "F_受電明細"
End Sub

'==============================================================================
' 入力もれチェック
'==============================================================================
Private Sub Build_F_未入力チェック()
    Dim frm As Access.Form, c As Access.Control, t As Long
    Set frm = CreateForm()

    AddLabel frm, acDetail, "入力もれチェック", 240, 200, 5000, 440, 14, True
    AddLabel frm, acDetail, _
        "出勤登録があるのに受付・業務の実績が 1 件も無い担当者を出します。" & _
        "現行 Excel で誰にも気付かれずデータが欠けていた問題への対策です。", _
        240, 660, 11000, 280, 9, False

    t = 1000
    AddLabel frm, acDetail, "対象日", 240, t + 40, 1000, ROW_H, 11, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 1240, t, 1800, ROW_H)
    c.Name = "対象日"
    c.Format = "yyyy/mm/dd"
    c.DefaultValue = "=Date()"
    AddButton frm, 3200, t - 20, 1400, ROW_H + 40, "チェック", "未入力_実行", 10

    t = t + ROW_H + GAP + 60
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", 240, t, 8000, 3600)
    c.Name = "sub結果"
    c.SourceObject = "Table.Q_未入力チェック"

    t = t + 3600 + GAP
    AddButton frm, 6800, t, 1440, ROW_H + 40, "閉じる", "閉じる", 10

    frm.Caption = "入力もれチェック"
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.AutoCenter = True
    frm.Section(acDetail).Height = t + ROW_H + 300
    frm.Width = 8600
    EndForm frm, "F_未入力チェック"
End Sub

'==============================================================================
' マスタ保守
'==============================================================================
Private Sub Build_F_マスタ_担当者()
    Dim frm As Access.Form, c As Access.Control
    Dim flds As Variant, w As Variant, i As Long, x As Long
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [M_担当者] ORDER BY [表示順];"

    flds = Array("担当者コード", "姓", "名", "氏名", "カナ", "職員区分", _
                 "在籍開始日", "在籍終了日", "表示順", "有効", "備考")
    w = Array(1200, 1400, 1400, 2400, 2000, 1200, 1600, 1600, 900, 800, 4000)
    x = 0
    For i = 0 To UBound(flds)
        Set c = CreateControl(frm.Name, IIf(flds(i) = "有効", acCheckBox, acTextBox), _
                              acDetail, "", CStr(flds(i)), x, 0, CLng(w(i)), ROW_H)
        c.Name = CStr(flds(i))
        x = x + CLng(w(i))
    Next i
    frm.Controls("在籍開始日").Format = "yyyy/mm/dd"
    frm.Controls("在籍終了日").Format = "yyyy/mm/dd"
    frm.Controls("在籍終了日").StatusBarText = _
        "退職日を入れると、その日以降は入力画面の候補に出なくなります。過去データはそのまま残ります。"

    AddUsageColumn frm, "担当者", "担当者ID", x
    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "担当者マスタ"
    frm.BeforeInsert = "=App_Event(""担当者_新規"")"
    frm.OnDelete = "=App_Delete(""担当者"")"
    EndForm frm, "F_マスタ_担当者"
End Sub

Private Sub Build_F_マスタ_製品()
    Dim frm As Access.Form, c As Access.Control
    Dim flds As Variant, w As Variant, i As Long, x As Long
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [M_製品] ORDER BY [ブロックID],[表示順];"

    flds = Array("製品ID", "ブロックID", "製品名", "適用開始日", "適用終了日", "表示順", "有効")
    w = Array(900, 1100, 4200, 1600, 1600, 900, 800)
    x = 0
    For i = 0 To UBound(flds)
        Set c = CreateControl(frm.Name, IIf(flds(i) = "有効", acCheckBox, acTextBox), _
                              acDetail, "", CStr(flds(i)), x, 0, CLng(w(i)), ROW_H)
        c.Name = CStr(flds(i))
        x = x + CLng(w(i))
    Next i
    frm.Controls("適用開始日").Format = "yyyy/mm/dd"
    frm.Controls("適用終了日").Format = "yyyy/mm/dd"

    AddUsageColumn frm, "製品", "製品ID", x
    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "製品マスタ"
    frm.BeforeInsert = "=App_Event(""製品_新規"")"
    frm.OnDelete = "=App_Delete(""製品"")"
    EndForm frm, "F_マスタ_製品"
End Sub

Private Sub Build_F_マスタ_区分()
    Dim frm As Access.Form, c As Access.Control
    Dim flds As Variant, w As Variant, i As Long, x As Long
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [M_区分] ORDER BY [ブロックID],[表示順];"

    flds = Array("区分ID", "ブロックID", "区分名", "集計列ID", "内訳区分", "表示順", "有効", "旧転記名")
    w = Array(800, 1100, 4000, 1100, 1200, 900, 800, 3200)
    x = 0
    For i = 0 To UBound(flds)
        Set c = CreateControl(frm.Name, IIf(flds(i) = "有効", acCheckBox, acTextBox), _
                              acDetail, "", CStr(flds(i)), x, 0, CLng(w(i)), ROW_H)
        c.Name = CStr(flds(i))
        x = x + CLng(w(i))
    Next i
    frm.Controls("集計列ID").StatusBarText = _
        "3=申込方法 4=抽選結果 5=納付書発送 6=商品発送 7=その他 8=商品交換。集計表のどの列に積むかを決めます。"
    frm.Controls("旧転記名").Locked = True
    frm.Controls("旧転記名").StatusBarText = "現行 Excel の Sheet2 3 行目にあった名前。移行時の照合用です。"

    AddUsageColumn frm, "区分", "区分ID", x
    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "区分マスタ"
    frm.BeforeInsert = "=App_Event(""区分_新規"")"
    frm.OnDelete = "=App_Delete(""区分"")"
    EndForm frm, "F_マスタ_区分"
End Sub

Private Sub Build_F_マスタ_業務項目()
    Dim frm As Access.Form, c As Access.Control
    Dim flds As Variant, w As Variant, i As Long, x As Long
    Set frm = CreateForm()
    frm.RecordSource = "SELECT * FROM [M_業務項目] ORDER BY [表示順];"

    flds = Array("業務項目ID", "番号", "項目名", "帳票表示名", "表示順", "有効")
    w = Array(1100, 700, 3400, 4200, 900, 800)
    x = 0
    For i = 0 To UBound(flds)
        Set c = CreateControl(frm.Name, IIf(flds(i) = "有効", acCheckBox, acTextBox), _
                              acDetail, "", CStr(flds(i)), x, 0, CLng(w(i)), ROW_H)
        c.Name = CStr(flds(i))
        x = x + CLng(w(i))
    Next i

    AddUsageColumn frm, "業務項目", "業務項目ID", x
    frm.DefaultView = 2
    frm.ViewsAllowed = 2
    frm.Caption = "業務項目マスタ"
    frm.BeforeInsert = "=App_Event(""業務項目_新規"")"
    frm.OnDelete = "=App_Delete(""業務項目"")"
    EndForm frm, "F_マスタ_業務項目"
End Sub

'==============================================================================
' 生成ヘルパ
'==============================================================================
Private Sub EndForm(ByRef frm As Access.Form, ByVal finalName As String)
    Dim tmp As String
    tmp = frm.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DeleteObjectIfExists acForm, finalName
    DoCmd.Rename finalName, acForm, tmp
    Echo_ "  画面 " & finalName
End Sub

Private Function AddLabel(ByRef frm As Access.Form, ByVal sect As Integer, _
                          ByVal caption As String, ByVal l As Long, ByVal t As Long, _
                          ByVal w As Long, ByVal h As Long, _
                          Optional ByVal fontSize As Integer = 9, _
                          Optional ByVal bold As Boolean = False, _
                          Optional ByVal align As Integer = 1) As Access.Control
    Dim c As Access.Control
    Set c = CreateControl(frm.Name, acLabel, sect, "", "", l, t, w, h)
    c.caption = caption
    c.fontSize = fontSize
    c.FontBold = bold
    c.TextAlign = align
    Set AddLabel = c
End Function

Private Function AddButton(ByRef frm As Access.Form, ByVal l As Long, ByVal t As Long, _
                           ByVal w As Long, ByVal h As Long, ByVal caption As String, _
                           ByVal action As String, _
                           Optional ByVal fontSize As Integer = 10) As Access.Control
    Dim c As Access.Control
    Set c = CreateControl(frm.Name, acCommandButton, acDetail, "", "", l, t, w, h)
    c.caption = caption
    c.fontSize = fontSize
    c.OnClick = "=App_Click(""" & action & """)"
    Set AddButton = c
End Function

 ' 「使用件数」列。削除して良いかどうかが、消す前に一目で分かるようにする。
Private Sub AddUsageColumn(ByRef frm As Access.Form, ByVal kind As String, _
                           ByVal idField As String, ByVal x As Long)
    Dim c As Access.Control
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", x, 0, 1100, ROW_H)
    c.Name = "使用件数"
    c.ControlSource = "=MasterUsage(""" & kind & """,[" & idField & "])"
    c.TextAlign = 3
    c.Locked = True
    c.StatusBarText = "この" & kind & "を参照している実績の件数。" & _
                      "0 でない行は削除できません。「有効」を外して無効にしてください。"
End Sub

' RowSourceType は日本語版 Access では値が日本語なので、英語・日本語の両方を試す。
Private Sub SetRowSource(ByRef c As Access.Control, ByVal rowSource As String, _
                         ByVal colCount As Integer, ByVal colWidths As String, _
                         ByVal listWidth As Long)
    On Error Resume Next
    c.RowSourceType = "Table/Query"
    If Err.Number <> 0 Then
        Err.Clear
        c.RowSourceType = "テーブル/クエリ"
        Err.Clear
    End If
    On Error GoTo 0
    c.RowSource = rowSource
    c.ColumnCount = colCount
    c.ColumnWidths = colWidths
    c.BoundColumn = 1
    c.ListWidth = listWidth
    c.LimitToList = True
End Sub
