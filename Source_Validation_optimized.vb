Sub UpdateSource_ValidationTables()
    Const SHEET_SOURCE As String = "CMDB Export"
    Const SHEET_TABLE As String = "Source Data"
    Const SHEET_VALIDATION As String = "Validation"
    Const TABLE_SOURCE As String = "SourceData"
    Const TABLE_VALIDATION As String = "Validation"
    Const COL_UNIQUE_KEY As String = "Number"
    Const COL_SORT_SECOND As String = "URL Type"

    Dim wsSource As Worksheet, wsTable As Worksheet, wsValidation As Worksheet
    Dim tblSource As ListObject, tblValidation As ListObject
    Dim lastRow As Long, colCount As Long
    Dim dataRange As Range

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Set worksheets and tables
    Set wsSource = ThisWorkbook.Sheets(SHEET_SOURCE)
    Set wsTable = ThisWorkbook.Sheets(SHEET_TABLE)
    Set wsValidation = ThisWorkbook.Sheets(SHEET_VALIDATION)

    Set tblSource = wsTable.ListObjects(TABLE_SOURCE)
    Set tblValidation = wsValidation.ListObjects(TABLE_VALIDATION)

    ' Determine data range from source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "No data to copy from '" & SHEET_SOURCE & "' sheet.", vbInformation
        GoTo Cleanup
    End If

    colCount = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Set dataRange = wsSource.Range("A2").Resize(lastRow - 1, colCount)

    ' Clear old data from SourceData table
    If tblSource.ListRows.Count > 0 Then tblSource.DataBodyRange.Delete

    ' Resize SourceData table to match new data
    tblSource.Resize tblSource.HeaderRowRange.Resize(dataRange.Rows.Count + 1, colCount)

    ' Copy data to table
    tblSource.DataBodyRange.Value = dataRange.Value

    ' Sort SourceData table
    With tblSource.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tblSource.ListColumns(COL_UNIQUE_KEY).Range, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=tblSource.ListColumns(COL_SORT_SECOND).Range, _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    MsgBox "SourceData table updated successfully!", vbInformation

    ' === Update Validation Table ===

    Dim dict As Object, arrSource As Variant
    Dim i As Long, key As Variant
    Set dict = CreateObject("Scripting.Dictionary")

    arrSource = tblSource.ListColumns(COL_UNIQUE_KEY).DataBodyRange.Value
    For i = 1 To UBound(arrSource, 1)
        If Not dict.Exists(arrSource(i, 1)) Then
            dict.Add arrSource(i, 1), arrSource(i, 1)
        End If
    Next i

    ' Delete all rows from Validation table
    If tblValidation.ListRows.Count > 0 Then tblValidation.DataBodyRange.Delete
    If dict.Count = 0 Then
        MsgBox "Validation table cleared. No unique values found.", vbInformation
        GoTo Cleanup
    End If

    ' Resize Validation table to fit new number of rows (retain 23 columns)
    tblValidation.Resize tblValidation.HeaderRowRange.Resize(dict.Count + 1, tblValidation.ListColumns.Count)

    ' Write unique keys to first column
    tblValidation.DataBodyRange.Columns(1).Value = Application.WorksheetFunction.Transpose(dict.Keys)

    ' Copy formulas from first data row to rest (excluding "Number" column)
    Dim topRow As Range
    Set topRow = tblValidation.ListRows(1).Range
    If tblValidation.ListRows.Count > 1 Then
        topRow.Offset(1, 1).Resize(tblValidation.ListRows.Count - 1, tblValidation.ListColumns.Count - 1).Formula = _
            topRow.Offset(0, 1).Resize(1, tblValidation.ListColumns.Count - 1).Formula
    End If

    MsgBox "Validation table updated successfully!", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
