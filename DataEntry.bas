Attribute VB_Name = "Module1"
Option Explicit

'========================
' CONFIG (edit if needed)
'========================
Private Const DB_SHEET As String = "Rationalized_DB"
Private Const DB_TABLE As String = "tblDB"

Private Const ENTRY_SHEET As String = "DB_Entry"
Private Const ENTRY_TABLE As String = "tblEntry"

'========================
' STEP 3A: Create entry sheet + table
'========================
Public Sub CreateEntryInterface()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDB As Worksheet, wsE As Worksheet
    Dim loDB As ListObject, loE As ListObject
    Dim hdr As Variant
    Dim lastCol As Long, c As Long
    
    Set wsDB = wb.Worksheets(DB_SHEET)
    Set loDB = wsDB.ListObjects(DB_TABLE)
    
    'Create / reset entry sheet
    Set wsE = GetOrCreateSheet(wb, ENTRY_SHEET)
    wsE.Cells.Clear
    
    'Title + instructions
    wsE.Range("A1").Value = "Data Entry (writes into " & DB_TABLE & ")"
    wsE.Range("A1").Font.Bold = True
    wsE.Range("A2").Value = "Fill rows in the table below, then run CommitEntries."
    
    'Copy headers from DB table
    hdr = loDB.HeaderRowRange.Value '1 x n array
    lastCol = loDB.ListColumns.Count
    
    wsE.Range("A4").Resize(1, lastCol).Value = hdr
    wsE.Range("A4").Resize(1, lastCol).Font.Bold = True
    
    'Create entry table with 20 blank rows
    Dim rng As Range
    Set rng = wsE.Range("A4").Resize(21, lastCol) 'header + 20 rows
    Set loE = wsE.ListObjects.Add(xlSrcRange, rng, , xlYes)
    loE.name = ENTRY_TABLE
    loE.TableStyle = "TableStyleLight9"
    
    wsE.Columns.AutoFit
    wsE.Activate
    
    'Apply validations on Item/Year/Month columns if present
    ApplyEntryValidations loDB, loE
    
    MsgBox "Entry sheet created: " & ENTRY_SHEET & " (" & ENTRY_TABLE & ")", vbInformation
End Sub

'========================
' STEP 3B: Push entry rows into DB (append or update)
'========================
Public Sub CommitEntries()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsDB As Worksheet, wsE As Worksheet
    Dim loDB As ListObject, loE As ListObject
    
    Set wsDB = wb.Worksheets(DB_SHEET)
    Set loDB = wsDB.ListObjects(DB_TABLE)
    
    Set wsE = wb.Worksheets(ENTRY_SHEET)
    Set loE = wsE.ListObjects(ENTRY_TABLE)
    
    If loE.DataBodyRange Is Nothing Then
        MsgBox "No rows to commit in " & ENTRY_TABLE, vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Build fast lookup: key = Item|Year|Month -> DB row index in DataBodyRange
    Dim keyToRow As Object: Set keyToRow = CreateObject("Scripting.Dictionary")
    BuildDBKeyIndex loDB, keyToRow
    
    'Map column names -> column index (both tables share same headers)
    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To loDB.ListColumns.Count
        colMap(loDB.ListColumns(i).name) = i
    Next i
    
    'Identify required columns
    If Not colMap.Exists("Item") Or Not colMap.Exists("Year") Or Not colMap.Exists("Month") Then
        MsgBox "DB table must contain columns: Item, Year, Month", vbCritical
        GoTo CleanExit
    End If
    
    Dim r As Long, wrote As Long, updated As Long, skipped As Long
    Dim itemV As String, yearV As String, monthV As String, k As String
    
    For r = 1 To loE.DataBodyRange.rows.Count
        itemV = Trim$(CStr(loE.DataBodyRange.Cells(r, colMap("Item")).Value))
        yearV = Trim$(CStr(loE.DataBodyRange.Cells(r, colMap("Year")).Value))
        monthV = Trim$(CStr(loE.DataBodyRange.Cells(r, colMap("Month")).Value))
        
        If Len(itemV) = 0 And Len(yearV) = 0 And Len(monthV) = 0 Then
            'completely blank row
            GoTo NextRow
        End If
        
        If Len(itemV) = 0 Or Len(yearV) = 0 Or Len(monthV) = 0 Then
            skipped = skipped + 1
            GoTo NextRow
        End If
        
        'Skip unwanted values (safety)
        If UCase$(yearV) = "202X" Then GoTo NextRow
        If UCase$(monthV) = "FY" Then GoTo NextRow
        
        k = MakeKey(itemV, yearV, monthV)
        
        If keyToRow.Exists(k) Then
            'Update existing row
            WriteRowValues loE, r, loDB, CLng(keyToRow(k)), colMap
            updated = updated + 1
        Else
            'Append new row
            Dim newRow As ListRow
            Set newRow = loDB.ListRows.Add
            WriteRowValuesToListRow loE, r, newRow, colMap
            keyToRow(k) = newRow.Index 'index in DataBodyRange
            wrote = wrote + 1
        End If
        
        'Clear committed entry row for speed
        ClearEntryRowKeepFormulas loE, r
        
NextRow:
    Next r
    
    'Optional: refresh PQ / pivots etc
    ThisWorkbook.RefreshAll
    
    EnsureDBDiffColumnsAreFormulas loDB
    loDB.Range.Calculate
    MsgBox "Committed. Appended: " & wrote & " | Updated: " & updated & " | Skipped (incomplete): " & skipped, vbInformation
    
    Call EnsureItemsListName(ThisWorkbook.Worksheets(DB_SHEET).ListObjects(DB_TABLE))

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'========================
' HELPERS
'========================
Private Sub BuildDBKeyIndex(ByVal loDB As ListObject, ByVal dict As Object)
    dict.RemoveAll
    
    If loDB.DataBodyRange Is Nothing Then Exit Sub
    
    Dim colItem As Long, colYear As Long, colMonth As Long
    colItem = GetColIndex(loDB, "Item")
    colYear = GetColIndex(loDB, "Year")
    colMonth = GetColIndex(loDB, "Month")
    
    If colItem = 0 Or colYear = 0 Or colMonth = 0 Then Exit Sub
    
    Dim r As Long
    Dim itemV As String, yearV As String, monthV As String, k As String
    
    For r = 1 To loDB.DataBodyRange.rows.Count
        itemV = Trim$(CStr(loDB.DataBodyRange.Cells(r, colItem).Value))
        yearV = Trim$(CStr(loDB.DataBodyRange.Cells(r, colYear).Value))
        monthV = Trim$(CStr(loDB.DataBodyRange.Cells(r, colMonth).Value))
        If Len(itemV) > 0 And Len(yearV) > 0 And Len(monthV) > 0 Then
            k = MakeKey(itemV, yearV, monthV)
            If Not dict.Exists(k) Then dict.Add k, r
        End If
    Next r
End Sub

Private Function MakeKey(ByVal itemV As String, ByVal yearV As String, ByVal monthV As String) As String
    MakeKey = UCase$(Trim$(itemV)) & "|" & UCase$(Trim$(yearV)) & "|" & UCase$(Trim$(monthV))
End Function

Private Sub WriteRowValues(ByVal loE As ListObject, ByVal entryRow As Long, _
                           ByVal loDB As ListObject, ByVal dbRowIndex As Long, _
                           ByVal colMap As Object)
    Dim c As Long, colName As String
    For c = 1 To loDB.ListColumns.Count
        colName = loDB.ListColumns(c).name
        If Not IsCalcColumn(colName) Then
            loDB.DataBodyRange.Cells(dbRowIndex, c).Value = CleanDiv0ToNA(loE.DataBodyRange.Cells(entryRow, c).Value)
        End If
    Next c
End Sub
Private Sub WriteRowValuesToListRow(ByVal loE As ListObject, ByVal entryRow As Long, _
                                    ByVal newRow As ListRow, ByVal colMap As Object)
    Dim c As Long, colName As String
    For c = 1 To newRow.Parent.ListColumns.Count
        colName = newRow.Parent.ListColumns(c).name
        If Not IsCalcColumn(colName) Then
            newRow.Range.Cells(1, c).Value = CleanDiv0ToNA(loE.DataBodyRange.Cells(entryRow, c).Value)
        End If
    Next c
End Sub

Private Function CleanDiv0ToNA(ByVal v As Variant) As Variant
    If IsError(v) Then
        If v = CVErr(xlErrDiv0) Then
            CleanDiv0ToNA = CVErr(xlErrNA)
        Else
            CleanDiv0ToNA = v
        End If
    Else
        CleanDiv0ToNA = v
    End If
End Function

Private Function GetColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    On Error GoTo Fail
    GetColIndex = lo.ListColumns(colName).Index
    Exit Function
Fail:
    GetColIndex = 0
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = name
    End If
End Function

Private Sub ApplyEntryValidations(ByVal loDB As ListObject, ByVal loE As ListObject)
    'Build lists on the entry sheet (hidden columns) and apply data validation
    Dim wsE As Worksheet: Set wsE = loE.Parent
    Dim colItem As Long, colYear As Long, colMonth As Long
    colItem = GetColIndex(loE, "Item")
    colYear = GetColIndex(loE, "Year")
    colMonth = GetColIndex(loE, "Month")
    
    Dim startCol As Long: startCol = loE.Range.Column + loE.Range.Columns.Count + 2 'to the right of table
    
    'Unique lists from DB
    Dim items As Variant, years As Variant, months As Variant
    items = UniqueListFromColumn(loDB, "Item")
    years = UniqueListFromColumn(loDB, "Year")
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec") 'clean months
    
    'Write lists
    Dim rngItems As Range, rngYears As Range, rngMonths As Range
    Set rngItems = WriteList(wsE, startCol, 2, items, "ItemsList")
    Set rngYears = WriteList(wsE, startCol + 1, 2, years, "YearsList")
    Set rngMonths = WriteList(wsE, startCol + 2, 2, months, "MonthsList")
    
    'Apply validation to entire data body columns
    If colItem > 0 Then ApplyDV loE.ListColumns(colItem).DataBodyRange, rngItems
    'If colYear > 0 Then ApplyDV loE.ListColumns(colYear).DataBodyRange, rngYears
    If colMonth > 0 Then ApplyDV loE.ListColumns(colMonth).DataBodyRange, rngMonths
    
    'Hide helper columns
    wsE.Columns(startCol).Hidden = True
    wsE.Columns(startCol + 1).Hidden = True
    wsE.Columns(startCol + 2).Hidden = True
End Sub

Private Sub ApplyDV(ByVal target As Range, ByVal listRange As Range)
    If target Is Nothing Then Exit Sub
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:="=" & listRange.Worksheet.name & "!" & listRange.Address
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

Private Function UniqueListFromColumn(ByVal lo As ListObject, ByVal colName As String) As Variant
    Dim colIdx As Long: colIdx = GetColIndex(lo, colName)
    If colIdx = 0 Or lo.DataBodyRange Is Nothing Then
        UniqueListFromColumn = Array()
        Exit Function
    End If
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, v As String
    
    For r = 1 To lo.DataBodyRange.rows.Count
        v = Trim$(CStr(lo.DataBodyRange.Cells(r, colIdx).Value))
        If Len(v) > 0 Then
            If UCase$(v) <> "202X" And UCase$(v) <> "FY" Then
                If Not dict.Exists(v) Then dict.Add v, True
            End If
        End If
    Next r
    
    UniqueListFromColumn = dict.keys
End Function

Private Function WriteList(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long, _
                           ByVal arr As Variant, ByVal nameTag As String) As Range
    Dim n As Long, i As Long
    
    'Clear column
    ws.Columns(col).ClearContents
    
    If IsArray(arr) Then
        On Error Resume Next
        n = UBound(arr) - LBound(arr) + 1
        On Error GoTo 0
    Else
        n = 0
    End If
    
    If n <= 0 Then
        ws.Cells(startRow, col).Value = "" 'empty list
        Set WriteList = ws.Cells(startRow, col)
        Exit Function
    End If
    
    For i = 0 To n - 1
        ws.Cells(startRow + i, col).Value = arr(LBound(arr) + i)
    Next i
    
    Set WriteList = ws.Range(ws.Cells(startRow, col), ws.Cells(startRow + n - 1, col))
End Function
Private Function NormalizeHeader(ByVal s As String) As String
    s = Replace(s, Chr(160), " ") ' non-breaking spaces
    NormalizeHeader = LCase$(Trim$(s))
End Function

Private Function IsCalcColumn(ByVal colName As String) As Boolean
    Dim n As String
    n = NormalizeHeader(colName)
    IsCalcColumn = (n = NormalizeHeader("Difference forecast/sales") Or _
                    n = NormalizeHeader("Difference forecast/sales%"))
End Function

Private Sub ClearEntryRowKeepFormulas(ByVal loE As ListObject, ByVal entryRowIndex As Long)
    Dim c As Long, colName As String
    For c = 1 To loE.ListColumns.Count
        colName = loE.ListColumns(c).name
        If Not IsCalcColumn(colName) Then
            loE.DataBodyRange.Cells(entryRowIndex, c).ClearContents
        End If
    Next c
End Sub

Private Sub EnsureDBDiffColumnsAreFormulas(ByVal loDB As ListObject)
    On Error GoTo SafeExit

    'Difference forecast/sales
    loDB.ListColumns("Difference forecast/sales").DataBodyRange.Formula = _
        "=IF(OR([@[Forecast CY]]="""",[@Invoiced]=""""),"""",[@[Forecast CY]]-[@Invoiced])"

    'Difference forecast/sales%
    loDB.ListColumns("Difference forecast/sales%").DataBodyRange.Formula = _
        "=IF(OR([@[Forecast CY]]="""",[@Invoiced]=""""),"""",IFERROR([@[Difference forecast/sales]]/[@Invoiced],NA()))"

SafeExit:
End Sub
