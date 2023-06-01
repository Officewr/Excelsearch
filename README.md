Const SEARCH_WORD = "\*.doc*"
Const SHEET_OUTPUT = "search"
Const CELL_PRINT_COL = 1
Const CELL_PRINT_ROW = 6
Const CELL_SEARCH_WORD = "B3"

Dim nowRow As Long

' メイン処理
Sub searchMacro()
    Dim buf As String
    Dim Path As String
    Dim myBook As Workbook
    nowRow = CELL_PRINT_ROW
    Set myBook = ThisWorkbook
    
    If Range(CELL_SEARCH_WORD) <> "" Then
        Path = getFolderName()
        
        Call reset
        
        Call searchFile(Path, myBook)
        
        If nowRow = CELL_PRINT_ROW Then
            MsgBox "検索結果：「" & Range(CELL_SEARCH_WORD) & "」が含まれるファイルはありませんでした。"
        Else
            MsgBox "検索結果：「" & Range(CELL_SEARCH_WORD) & "」が含まれるファイルが" & nowRow - CELL_PRINT_ROW & "件ヒットしました！"
        End If
    Else
        MsgBox "検索ワードを入力してください"
    End If
End Sub

' シートの初期化
Private Sub reset()
    Application.ScreenUpdating = False
    Sheets(SHEET_OUTPUT).UsedRange.ClearContents
    Application.ScreenUpdating = True
End Sub

' 再帰的にファイルを検索
Private Sub searchFile(ByVal Path As String, ByRef myBook As Workbook)
    On Error Resume Next

    Dim buf As String, f As Object
    buf = Dir(Path & SEARCH_WORD)
    
    searchWord = Range(CELL_SEARCH_WORD)
        
    Do While buf <> ""
        Call grepWord(searchWord, myBook, Path, buf)
        buf = Dir()
    Loop
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).Files
            If UCase(Right(f.Name, 4)) = ".DOC" Then
                Call grepWord(searchWord, myBook, Path, f.Name)
            End If
        Next f
        For Each f In .GetFolder(Path).SubFolders
            Call searchFile(f.Path, myBook)
        Next f
    End With
End Sub

' ダイアログでフォルダ名取得
Private Function getFolderName()
    Dim folderPath As Variant
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = 0 Then
            End
        End If
        folderPath = .SelectedItems(1)
    End With
    
    getFolderName = folderPath
End Function

' ワードファイル内の文字検索
Private Sub grepWord(ByVal searchWord, ByRef myBook As Workbook, ByVal Path As String, ByVal buf As String)
    Dim filePath
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim wordRange As Object
    Dim findResult As Object
    
    fullPath = Path & "\" & buf

    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
    End If
    wordApp.Visible = False
    On Error GoTo 0
    
    Set wordDoc = wordApp.Documents.Open(fullPath, ReadOnly:=True)
    If Err.Number = 1004 Then
        Err.Clear
    Else
        For Each wordRange In wordDoc.StoryRanges
            Set findResult = wordRange.Find.Execute(FindText:=searchWord, MatchCase:=False, MatchWholeWord:=False)
            If findResult Then
                Do
                    Call writeSheet(myBook, Path, buf, wordDoc.Name, wordRange)
                    Set findResult = wordRange.Find.Execute
                Loop While findResult
            End If
        Next wordRange
    End If
    
    wordDoc.Close SaveChanges:=False
    wordApp.Quit
    
    Set wordRange = Nothing
    Set wordDoc = Nothing
    Set wordApp = Nothing
End Sub

' 検索結果を出力
Private Sub writeSheet(ByRef myBook As Workbook, _
                        ByVal Path As String, _
                        ByVal buf As String, _
                        ByVal docName As String, _
                        ByRef wordRange As Object)
    Dim outputSheet As Worksheet
    Dim outputCell As Range
    
    Set outputSheet = myBook.Worksheets(SHEET_OUTPUT)
    Set outputCell = outputSheet.Cells(nowRow, CELL_PRINT_COL)
    
    outputCell.Value = buf
    outputCell.Offset(0, 1).Value = Path
    outputCell.Offset(0, 2).Value = docName
    outputCell.Offset(0, 3).Value = wordRange.Text
    
    nowRow = nowRow + 1
End Sub
