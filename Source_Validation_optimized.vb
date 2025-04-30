Sub UpdateSource_ValidationTables()
    ' === Constants ===
    Const SHEET_CMDB As String = "CMDB Export"
    Const SHEET_SOURCE As String = "Source Data"
    Const SHEET_VALIDATION As String = "Validation"
    Const TABLE_SOURCE As String = "SourceData"
    Const TABLE_VALIDATION As String = "Validation"
    Const COLUMN_NUMBER As String = "Number"
    Const COLUMN_URL_TYPE As String = "URL Type"

    ' === Variable declarations ===
    Dim wsSource As Worksheet, wsTable As Worksheet, wsValidation As Worksheet
    Dim tblSource As ListObject, tblValidation As ListObject
    Dim lastRow As Long, colCount As Long
    Dim dataRange As Range
    Dim rngSource As Range
    Dim dict As Object
    Dim arrSource() As Variant, arrValidation() As Variant
    Dim i As Long, j As Long
    Dim key As Variant
    Dim topRow As Range

    ' === Preserve Excel state ===
    Dim calcState As XlCalculation
    Dim screenUpdateState As Boolean, eventsState As Boolean
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo Cleanup

    ' === Set worksheets and tables ===
    Set wsSource = ThisWorkbook.Sheets(SHEET_CMDB)
    Set wsTable = ThisWorkbook.Sheets(SHEET_SOURCE)
    Set tblSource = wsTable.ListObjects(TABLE_SOURCE)
    Set wsValidation = ThisWorkbook.Sheets(SHEET_VALIDATION)
    Set tblValidation = wsValidation.ListObjects(TABLE_VALIDATION)

    ' === Determine source data range ===
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    colCount = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

    If lastRow <= 1 Then
        MsgBox "No data to copy from 'CMDB Export' sheet.", vbExclamation
        GoTo Cleanup
    End If

    Set dataRange = wsSource.Range("A2").Resize(lastRow - 1, colCount)

    ' === Prepare SourceData table ===
    If tblSource.ListRows.Count > 0 Then
        tblSource.DataBodyRange.Delete
    Else
        tblSource.ListRows.Add
    End If

    tblSource.Resize tblSource.HeaderRowRange.Resize(dataRange.Rows.Count + 1, dataRange.Columns.Count)
    tblSource.DataBodyRange.Cells(1, 1).Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

    ' === Sort the SourceData table ===
    With tblSource.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tblSource.ListColumns(COLUMN_NUMBER).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=tblSource.ListColumns(COLUMN_URL_TYPE).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    MsgBox "SourceData table updated successfully!", vbInformation

    ' === Build dictionary of unique 'Number' values ===
    Set rngSource = tblSource.ListColumns(COLUMN_NUMBER).DataBodyRange
    arrSource = rngSource.Value
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(arrSource, 1)
        If Not dict.exists(arrSource(i, 1)) Then
            dict.Add arrSource(i, 1), arrSource(i, 1)
        End If
    Next i

    ' === Update Validation table ===
    If tblValidation.ListRows.Count > 0 Then
        tblValidation.DataBodyRange.Delete
    Else
        tblValidation.ListRows.Add
    End If

    ' Prepare array of unique values
    ReDim arrValidation(1 To dict.Count, 1 To 1)
    j = 1
    For Each key In dict.Keys
        arrValidation(j, 1) = key
        j = j + 1
    Next key

    ' Resize validation table
    tblValidation.Resize tblValidation.HeaderRowRange.Resize(UBound(arrValidation, 1) + 1, 1)
    tblValidation.DataBodyRange.Value = arrValidation

    ' Copy formulas from first data row (if applicable)
    If tblValidation.ListRows.Count > 1 Then
        Set topRow = tblValidation.ListRows(1).Range
        If tblValidation.ListColumns.Count > 1 Then
            topRow.Offset(0, 1).Resize(1, tblValidation.ListColumns.Count - 1).Copy
            tblValidation.DataBodyRange.Offset(1, 1).Resize(tblValidation.ListRows.Count - 1, tblValidation.ListColumns.Count - 1).PasteSpecial xlPasteFormulas
        End If
    End If

    MsgBox "Validation table updated successfully!", vbInformation

Cleanup:
    ' === Restore Excel state ===
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState

    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbCritical
    End If
End Sub
