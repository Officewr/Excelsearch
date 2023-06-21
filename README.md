Sub CopyData()

    Dim layoutSheet As Worksheet
    Dim commandSheet As Worksheet
    Dim summarySheet As Worksheet
    Dim lastRow As Long
    Dim summaryLastRow As Long
    Dim copyRange As Range
    Dim i As Long
    
    ' レイアウトシート、コマンドシート、まとめシートを設定する（シート名に応じて変更してください）
    Set layoutSheet = ThisWorkbook.Worksheets("レイアウトシート")
    Set commandSheet = ThisWorkbook.Worksheets("コマンドシート")
    Set summarySheet = ThisWorkbook.Worksheets("まとめシート")
    
    ' コピー範囲を初期化する
    Set copyRange = Nothing

    ' レイアウトシートのセルA1とセルA2の値を確認し、条件に合致する場合、コピー範囲を設定する
    If layoutSheet.Range("A1").Value = "条件1" Then
        ' セルA1に条件1が記載されている場合、コマンドシートのA列のセルA2からセルA50までをコピーする
        Set copyRange = commandSheet.Range("A2:A50")
    ElseIf layoutSheet.Range("A1").Value = "条件2" And layoutSheet.Range("A2").Value = "タイプ1" Then
        ' セルA1に条件2、セルA2にタイプ1が記載されている場合、コマンドシートのB列のセルA2からセルA50までをコピーする
        Set copyRange = commandSheet.Range("B2:B50")
    End If
    
    ' コピー範囲が存在する場合、まとめシートにデータを追加する
    If Not copyRange Is Nothing Then
        ' レイアウトシートのセルD5の値を1行目にコピーする
        layoutSheet.Range("D5").Copy Destination:=summarySheet.Cells(summaryLastRow + 1, 1)
        
        ' まとめシートの1行目を確認し、既にデータが入力されている列への追記を行う
        Dim firstRow As Range
        Set firstRow = summarySheet.Range("1:1")
        Dim col As Range
        
        For Each col In firstRow
            If WorksheetFunction.CountA(col) = 0 Then
                ' 列が空白の場合、コピー範囲の値を追記する
                col.Value = copyRange.Value
                Exit For
            End If
        Next col
    End If
    
End Sub
