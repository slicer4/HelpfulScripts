Sub UpdateHistoricalScores()
    Dim wsHistorical As Worksheet
    Dim wsNewScores As Worksheet
    Dim tblName As String
    Dim isValid As Boolean
    Dim dataRowCount As Long
    Dim i As Long
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim lastDataRow As Long

    ' Set worksheets
    Set wsHistorical = ThisWorkbook.Sheets("Historical Scores")
    Set wsNewScores = ThisWorkbook.Sheets("New Scores")

    ' Find last used row in column A of "New Scores", starting from row 3
    lastDataRow = wsNewScores.Cells(wsNewScores.Rows.Count, "A").End(xlUp).Row
    If lastDataRow < 3 Then
        MsgBox "No new scores found to copy.", vbExclamation
        Exit Sub
    End If

    dataRowCount = lastDataRow - 2 ' Because data starts at row 3

    ' Insert 4 new columns to the left of column A in Historical Scores
    wsHistorical.Columns("A:D").Insert Shift:=xlToRight

    ' Add headers
    wsHistorical.Range("A1").Value = "Number"
    wsHistorical.Range("B1").Value = "AVI Score"
    wsHistorical.Range("C1").Value = "HVA"

    ' Copy data from New Scores into Historical Scores
    For i = 1 To dataRowCount
        wsHistorical.Cells(i + 1, 1).Value = wsNewScores.Cells(i + 2, 1).Value ' A
        wsHistorical.Cells(i + 1, 2).Value = wsNewScores.Cells(i + 2, 6).Value ' F
        wsHistorical.Cells(i + 1, 3).Value = wsNewScores.Cells(i + 2, 7).Value ' G
    Next i

    ' Set the range for the new table (headers in Row 1)
    Set tblRange = wsHistorical.Range("A1:C" & dataRowCount + 1)

    ' Prompt user for table name
    isValid = False
    Do While Not isValid
        tblName = Application.InputBox("Enter a name for the new table:", "Table Name", Type:=2)
        If tblName = "False" Then
            ' User cancelled -- clean up inserted columns
            wsHistorical.Columns("A:D").Delete
            Exit Sub
        End If

        On Error Resume Next
        Set tbl = wsHistorical.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        tbl.Name = tblName
        isValid = (Err.Number = 0)
        On Error GoTo 0

        If Not isValid Then
            MsgBox "Invalid table name. Please try again.", vbExclamation
            If Not tbl Is Nothing Then tbl.Delete
        End If
    Loop

    MsgBox "Table '" & tblName & "' created successfully!", vbInformation
End Sub
