Sub UpdateSource_ValidationTables()
    ' === Constants ===
    Const SRC_WS As String = "CMDB Export"
    Const DATA_WS As String = "Source Data"
    Const VALID_WS As String = "Validation"
    Const SRC_TBL As String = "SourceData"
    Const VALID_TBL As String = "Validation"

    ' === Variables ===
    Dim wsSource As Worksheet, wsTable As Worksheet, wsValidation As Worksheet
    Dim tblSource As ListObject, tblValidation As ListObject
    Dim lastRow As Long, colCount As Long
    Dim dataRange As Range
    Dim rngSource As Range
    Dim dict As Object
    Dim arrSource() As Variant, arrValidation() As Variant
    Dim i As Long, j As Long
    Dim topRow As Range
    Dim Key As Variant

    ' === Performance Settings ===
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ' === Setup Worksheets and Tables ===
    Set wsSource = ThisWorkbook.Sheets(SRC_WS)
    Set wsTable = ThisWorkbook.Sheets(DATA_WS)
    Set wsValidation = ThisWorkbook.Sheets(VALID_WS)

    Set tblSource = wsTable.ListObjects(SRC_TBL)
    Set tblValidation = wsValidation.ListObjects(VALID_TBL)

    ' === Step 1: Copy Data to SourceData Table ===
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    If lastRow > 1 Then
        colCount = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        Set dataRange = wsSource.Range("A2").Resize(lastRow - 1, colCount)

        ' Clear old data
        If tblSource.ListRows.Count > 0 Then
            tblSource.DataBodyRange.ClearContents
        Else
            ' Add one row so DataBodyRange is not Nothing
            tblSource.ListRows.Add
        End If

        ' Resize table to match new data
        tblSource.Resize tblSource.HeaderRowRange.Resize(, colCount)

        ' Write new data
        Set rngSource = tblSource.DataBodyRange.Cells(1, 1)
        rngSource.Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

        ' Sort by Number, then URL Type
        With tblSource.Sort
            .SortFields.Clear
            .SortFields.Add Key:=tblSource.ListColumns("Number").Range, Order:=xlAscending
            .SortFields.Add Key:=tblSource.ListColumns("URL Type").Range, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With

    Else
        MsgBox "No data to copy from 'CMDB Export' sheet.", vbInformation
        GoTo Cleanup
    End If

    ' === Step 2: Update Validation Table ===

    ' Collect unique 'Number' values from SourceData table
    Set rngSource = tblSource.ListColumns("Number").DataBodyRange
    arrSource = rngSource.Value
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(arrSource, 1)
        If Len(Trim(arrSource(i, 1))) > 0 Then
            If Not dict.Exists(arrSource(i, 1)) Then
                dict.Add arrSource(i, 1), arrSource(i, 1)
            End If
        End If
    Next i

    ' Prepare array for Validation table
    ReDim arrValidation(1 To dict.Count, 1 To 1)
    j = 1
    For Each Key In dict.Keys
        arrValidation(j, 1) = Key
        j = j + 1
    Next Key

    ' Clear existing data (except header) from Validation table
    If tblValidation.ListRows.Count > 0 Then
        tblValidation.DataBodyRange.ClearContents
    End If

    ' Resize table if needed
    If tblValidation.ListRows.Count < dict.Count Then
        For i = tblValidation.ListRows.Count + 1 To dict.Count
            tblValidation.ListRows.Add
        Next i
    ElseIf tblValidation.ListRows.Count > dict.Count Then
        For i = tblValidation.ListRows.Count To dict.Count + 1 Step -1
            tblValidation.ListRows(i).Delete
        Next i
    End If

    ' Write unique numbers to Validation table
    tblValidation.DataBodyRange.Columns(1).Value = arrValidation

    ' Copy formulas from first data row to remaining rows (columns 2 and onward)
    Set topRow = tblValidation.ListRows(1).Range
    If tblValidation.ListRows.Count > 1 Then
        topRow.Offset(1, 1).Resize(tblValidation.ListRows.Count - 1, tblValidation.ListColumns.Count - 1).Formula = _
            topRow.Offset(0, 1).Resize(1, tblValidation.ListColumns.Count - 1).Formula
    End If

    MsgBox "SourceData and Validation tables updated successfully!", vbInformation

Cleanup:
    ' === Restore Settings ===
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
End Sub
