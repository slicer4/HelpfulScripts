Sub UpdateSource_ValidationTables()
    ' Wrapper for both tasks
    Call UpdateSourceTable
    Call UpdateValidationTable
End Sub

Sub UpdateSourceTable()
    Const SOURCE_SHEET As String = "CMDB Export"
    Const DEST_SHEET As String = "Source Data"
    Const TABLE_NAME As String = "SourceData"

    Dim wsSource As Worksheet, wsTable As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim dataRange As Range

    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(SOURCE_SHEET)
    Set wsTable = ThisWorkbook.Sheets(DEST_SHEET)
    Set tbl = wsTable.ListObjects(TABLE_NAME)
    On Error GoTo 0

    If wsSource Is Nothing Or wsTable Is Nothing Or tbl Is Nothing Then
        MsgBox "One or more required worksheets or tables are missing.", vbCritical
        Exit Sub
    End If

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    If lastRow <= 1 Then
        MsgBox "No data to copy from 'CMDB Export' sheet."
        Exit Sub
    End If

    Set dataRange = wsSource.Range("A2").Resize(lastRow - 1, wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column)

    ' Clear existing table data
    If tbl.DataBodyRange Is Nothing Then
        ' no rows to clear
    Else
        tbl.DataBodyRange.Delete
    End If

    ' Resize table to fit new data
    tbl.Resize tbl.HeaderRowRange.Resize(dataRange.Rows.Count + 1, dataRange.Columns.Count)
    tbl.DataBodyRange.Value = dataRange.Value

    ' Sort by Number and URL Type
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns("Number").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=tbl.ListColumns("URL Type").Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    MsgBox "SourceData table updated successfully!"
End Sub

Sub UpdateValidationTable()
    Const VALIDATION_SHEET As String = "Validation"
    Const VALIDATION_TABLE_NAME As String = "Validation"

    Dim wsValidation As Worksheet
    Dim tblValidation As ListObject
    Dim rngSource As Range
    Dim dict As Object
    Dim arrSource() As Variant
    Dim arrValidation() As Variant
    Dim i As Long, j As Long
    Dim Key As Variant
    Dim tblSource As ListObject

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set tblSource = ThisWorkbook.Sheets("Source Data").ListObjects("SourceData")
    Set wsValidation = ThisWorkbook.Sheets(VALIDATION_SHEET)
    Set tblValidation = wsValidation.ListObjects(VALIDATION_TABLE_NAME)

    Set rngSource = tblSource.ListColumns("Number").DataBodyRange
    If rngSource Is Nothing Then
        MsgBox "No source data available to update Validation table.", vbExclamation
        GoTo CleanExit
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    arrSource = rngSource.Value
    For i = 1 To UBound(arrSource, 1)
        If Not dict.exists(arrSource(i, 1)) Then
            dict.Add arrSource(i, 1), arrSource(i, 1)
        End If
    Next i

    ' Clear all existing rows
    Do While tblValidation.ListRows.Count > 0
        tblValidation.ListRows(tblValidation.ListRows.Count).Delete
    Loop

    ' Write unique values to validation table
    ReDim arrValidation(1 To dict.Count, 1 To 1)
    j = 1
    For Each Key In dict.Keys
        arrValidation(j, 1) = Key
        j = j + 1
    Next Key

    For i = 1 To UBound(arrValidation, 1)
        tblValidation.ListRows.Add
        tblValidation.DataBodyRange(i, 1).Value = arrValidation(i, 1)
    Next i

    ' Copy formulas from first row to all others, excluding first column
    Dim colCount As Long
    colCount = tblValidation.ListColumns.Count

    If tblValidation.ListRows.Count > 1 Then
        Dim rFirstRow As Range
        Set rFirstRow = tblValidation.ListRows(1).Range
    
        Dim col As ListColumn
        Dim formulaText As String
        Dim iRow As Long, iCol As Long
    
        For iCol = 2 To colCount ' Start at 2 to skip first column
            formulaText = rFirstRow.Cells(1, iCol).Formula
            If formulaText <> "" Then
                For iRow = 2 To tblValidation.ListRows.Count
                    tblValidation.DataBodyRange.Cells(iRow, iCol).Formula = formulaText
                Next iRow
            End If
        Next iCol
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Validation table updated successfully!"
End Sub
