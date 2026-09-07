<%
' =============================================================================
'  帳票 (日報) の本体
'
'  画面表示 (report.asp) と PDF 出力 (printpdf.asp) の両方がここを呼ぶ。
'  片方だけ直して食い違う、を避けるために 1 か所にまとめてある。
'
'  体裁は現行「日報集計印刷用フォーム.xlsm ＞ 印刷用」シートの再現。
'  値の出どころは T_受電 ただ 1 つなので、この紙と集計表は必ず一致する。
' =============================================================================

' 指定日の帳票 HTML を組み立てて返す。
Function SheetHtml(dt)
    Dim hd, rs, names(14), i, h
    ' 左右とも 8 行ぶん取る。右は 5 件しか使わないが、
    ' 8 行の表を回すので配列の大きさは左右そろえておく (添字エラーを避ける)。
    Dim leftN(7), leftV(7), rightN(7), rightV(7), n, tasks

    Ensure日報 dt
    Set hd = DbQuery("SELECT * FROM [Q_日報_ヘッダ] WHERE [対象日]=?", Array(dt))

    ' 出勤者を 2 列 × 7 行に流し込む (現行 F10:O16 と同じ枠)
    i = 0
    Set rs = DbQuery("SELECT [氏名] FROM [Q_日報_出勤] WHERE [対象日]=? " & _
                     "ORDER BY [表示順]", Array(dt))
    Do While Not rs.EOF And i < 14
        names(i) = rs("氏名") : i = i + 1
        rs.MoveNext
    Loop
    rs.Close

    ' 電話応対以外の業務を 左 8 件 / 右 5 件 に割る (現行の並びと同じ)
    n = 0
    Set tasks = DbQuery( _
      "SELECT TM.[業務項目ID],TM.[帳票表示名], " & _
      " (SELECT Nz(Sum(W.[件数]),0) FROM [T_業務実績] AS W " & _
      "   WHERE W.[対象日]=? AND W.[業務項目ID]=TM.[業務項目ID]) AS [件数] " & _
      "FROM [M_業務項目] AS TM WHERE TM.[有効]=True ORDER BY TM.[表示順]", Array(dt))
    Do While Not tasks.EOF
        If n < 8 Then
            leftN(n) = tasks("帳票表示名") : leftV(n) = tasks("件数")
        ElseIf n < 13 Then
            rightN(n - 8) = tasks("帳票表示名") : rightV(n - 8) = tasks("件数")
        End If
        n = n + 1
        tasks.MoveNext
    Loop
    tasks.Close

    h = "<div class=""sheet"">" & vbCrLf
    h = h & "<div class=""kanai"">【課内限り】</div>" & vbCrLf
    h = h & "<div class=""topright"">" & H(Wareki(dt)) & "</div>" & vbCrLf
    h = h & "<h1 class=""title"">電話応対報告書日報集計表</h1>" & vbCrLf

    ' 左 3 列は Excel の結合セルと同じく rowspan でまとめる。
    ' セルを空けて罫線を消すと表の枠が切れてしまうため。
    h = h & "<table><tr>" & vbCrLf
    h = h & "<th rowspan=""8"" class=""mid"" style=""width:22mm"">出勤者</th>" & vbCrLf
    h = h & "<td rowspan=""8"" class=""mid num"" style=""width:14mm"">" & _
            H(hd("出勤者数")) & "</td>" & vbCrLf
    h = h & "<td rowspan=""8"" class=""mid"" style=""width:10mm"">名</td>" & vbCrLf
    h = h & "<th style=""width:52mm"">氏名</th>" & vbCrLf
    h = h & "<th style=""width:52mm"">氏名</th><th>備考</th></tr>" & vbCrLf
    For i = 0 To 6
        h = h & "<tr><td>" & H(names(i)) & "&nbsp;</td>" & _
                "<td>" & H(names(i + 7)) & "&nbsp;</td><td>&nbsp;</td></tr>" & vbCrLf
    Next
    h = h & "</table>" & vbCrLf

    h = h & "<p style=""margin:6px 0; font-size:9.5pt"">（　<b>" & _
            H(hd("回線数")) & "</b>　回線 ）</p>" & vbCrLf

    h = h & "<table><tr>" & vbCrLf
    h = h & "<th rowspan=""2"" style=""width:26mm"">問合せ件数<br>" & _
            "<span style=""font-size:8pt"">（　）内は職員受電数</span></th>" & vbCrLf
    h = h & "<td rowspan=""2"" class=""big"" style=""width:24mm"">" & _
            H(hd("合計")) & " 件</td>" & vbCrLf
    h = h & "<th style=""text-align:center"">申込</th>" & _
            "<th style=""text-align:center"">抽選</th>" & _
            "<th style=""text-align:center"">払込用紙</th>" & _
            "<th style=""text-align:center"">商品発送</th>" & _
            "<th style=""text-align:center"">その他</th></tr><tr>" & vbCrLf
    h = h & Cell(hd("申込"), hd("申込_職員"))
    h = h & Cell(hd("抽選"), hd("抽選_職員"))
    h = h & Cell(hd("払込用紙"), hd("払込用紙_職員"))
    h = h & Cell(hd("商品発送"), hd("商品発送_職員"))
    h = h & Cell(hd("その他"), hd("その他_職員"))
    h = h & "</tr></table>" & vbCrLf

    h = h & "<p style=""text-align:right; margin:6px 0"">内　　交換 <b>" & _
            H(hd("内交換")) & "</b> 件　／　内　　返金 <b>" & _
            H(hd("内返金")) & "</b> 件</p>" & vbCrLf

    h = h & Memo("特記事項（報告書の電話件数だけでは伝わり難い事項など）", hd("特記事項"))
    h = h & Memo("職員に代わった案件（概要）", hd("職員代替案件"))
    h = h & Memo("要望（お客様からの問い合わせを減らすための改善提案等）", hd("要望"))

    h = h & "<div class=""cap"">電話対応以外の業務（データ入力作業等）</div>" & vbCrLf
    h = h & "<table>" & vbCrLf
    For i = 0 To 7
        h = h & "<tr>" & _
            "<td style=""width:34%"">" & H(leftN(i)) & "&nbsp;</td>" & _
            "<td style=""width:9%; text-align:right"">" & H(leftV(i)) & "</td>" & _
            "<td style=""width:5%"">" & IIfS(Len("" & leftN(i)) > 0, "件", "") & "</td>" & _
            "<td style=""width:34%"">" & H(rightN(i)) & "&nbsp;</td>" & _
            "<td style=""width:9%; text-align:right"">" & H(rightV(i)) & "</td>" & _
            "<td style=""width:9%"">" & _
                IIfS(Len("" & rightN(i)) > 0, "件", "") & "</td>" & _
            "</tr>" & vbCrLf
    Next
    h = h & "</table>" & vbCrLf
    h = h & "</div>" & vbCrLf

    hd.Close
    SheetHtml = h
End Function

' 件数セル (大きな数字と、その下の職員内訳)
Function Cell(v, staff)
    Cell = "<td class=""big"">" & H(v) & _
           "<div class=""sub"">(" & H(staff) & ")</div></td>" & vbCrLf
End Function

' 見出しつきの記述欄
Function Memo(caption, body)
    Memo = "<div class=""cap"">" & caption & "</div>" & vbCrLf & _
           "<div class=""memo"">" & H(body) & "</div>" & vbCrLf
End Function
%>
