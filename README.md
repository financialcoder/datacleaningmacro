# datacleaningmacro
Sub CountBlanksInColumn()
    Dim colName As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim blankCount As Long
    Dim cell As Range

    ' Set the worksheet (change "Sheet1" to your sheet name)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Prompt the user for the column name
    colName = InputBox("Enter the column name (e.g., A, B, C, ...):")

    ' Find the last filled row in the specified column
    lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row
    
    ' Count the number of blank cells in the specified column
    blankCount = 0
    For Each cell In ws.Range(colName & "1:" & colName & lastRow)
        If IsEmpty(cell.Value) Then
            blankCount = blankCount + 1
        End If
    Next cell

    ' Display the result
    MsgBox "Number of blank cells in column " & colName & ": " & blankCount
End Sub
