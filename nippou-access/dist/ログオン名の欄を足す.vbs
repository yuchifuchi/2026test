' ============================================================================
'  電話応対日報 集計システム
'  すでに使っている 日報集計_be.accdb に「ログオン名」の欄を足します
'
'  いつ使うのか
'      すでに受付入力を始めていて、データを消したくないときだけ。
'      まだ使い始めていないなら、新しい 日報集計_be.accdb に
'      入れ替えるほうが簡単です。
'
'  使いかた
'      このファイルの上に 日報集計_be.accdb をドラッグして落とすか、
'      ダブルクリックして、出てきた窓にファイルの場所を貼り付けてください。
'
'  Access は要りません (Web サーバーに入っている ACE でも動きます)。
'  何度実行しても大丈夫です (すでに欄があれば、そのまま終わります)。
' ============================================================================
Option Explicit

Dim fso, path, conn, msg
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count > 0 Then
  path = WScript.Arguments(0)
Else
  path = InputBox("日報集計_be.accdb の場所を貼り付けてください。" & vbCrLf & vbCrLf & _
                  "(エクスプローラーのアドレス欄からコピーできます)", _
                  "電話応対日報", "\\サーバー名\共有フォルダ\日報\日報集計_be.accdb")
End If

If Len(Trim(path & "")) = 0 Then WScript.Quit
If Not fso.FileExists(path) Then
  MsgBox "ファイルが見つかりません。" & vbCrLf & vbCrLf & path, vbCritical, "電話応対日報"
  WScript.Quit
End If

On Error Resume Next
Set conn = CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & path & ";"
If Err.Number <> 0 Then
  MsgBox "ファイルを開けませんでした。" & vbCrLf & vbCrLf & _
         Err.Description & vbCrLf & vbCrLf & _
         "・誰かが Access で開いていないか" & vbCrLf & _
         "・このパソコンに Access か Access Database Engine が入っているか" & vbCrLf & _
         "を確認してください。", vbCritical, "電話応対日報"
  WScript.Quit
End If

msg = ""
Err.Clear
conn.Execute "ALTER TABLE [M_担当者] ADD COLUMN [ログオン名] TEXT(50)"
If Err.Number <> 0 Then
  If InStr(LCase(Err.Description), "already exists") > 0 Or _
     InStr(Err.Description, "既に") > 0 Or InStr(Err.Description, "すでに") > 0 Then
    msg = "「ログオン名」の欄は、すでにありました。何もしていません。"
  Else
    MsgBox "欄を追加できませんでした。" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "電話応対日報"
    conn.Close
    WScript.Quit
  End If
Else
  msg = "「ログオン名」の欄を追加しました。"
End If

Err.Clear
conn.Execute "CREATE INDEX [IX_担当者_ログオン名] ON [M_担当者] ([ログオン名])"
Err.Clear                        ' 索引が既にあってもかまわない
conn.Close

MsgBox msg & vbCrLf & vbCrLf & _
       "このあと、Web の「マスタ保守」→「担当者」タブで、" & vbCrLf & _
       "職員のかたのログオン名を入れて「保存」してください。", _
       vbInformation, "電話応対日報"
