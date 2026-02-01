Attribute VB_Name = "Module2"
Option Explicit

Private Const DB_SHEET As String = "CleanedUpData"
Private Const DB_TABLE As String = "tblDB_1"

Private Const READ_SHEET As String = "DB_Read"
Private Const MAX_PRODUCTS As Long = 3

'========================
' 1) Create the reading UI
'========================
Public Sub CreateReadingSheet()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet, wsDB As Worksheet
    Dim loDB As ListObject

    Set wsDB = wb.Worksheets(DB_SHEET)
    Set loDB = wsDB.ListObjects(DB_TABLE)

    Set ws = GetOrCreateSheet(wb, READ_SHEET)
    ws.Cells.Clear

    'Header
    ws.Range("A1").Value = "Product Reader"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Pick product(s), choose view & layout, then click Refresh."

    'Inputs
    ws.Range("A4").Value = "Product 1"
    ws.Range("A5").Value = "Product 2 (optional)"
    ws.Range("A6").Value = "Product 3 (optional)"
    ws.Range("A8").Value = "View"
    ws.Range("A9").Value = "Layout"

    ws.Range("B4:B6").ClearContents
    ws.Range("B8").Value = "Monthly"
    ws.Range("B9").Value = "Chronological"

    'Dropdowns: View + Layout
    With ws.Range("B8").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Monthly,Yearly"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With ws.Range("B9").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Chronological,Original"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    'Prepare item list (hidden)
    Dim listName As String
    listName = EnsureItemsListName(loDB)
    
    ApplyDVByName ws.Range("B4"), listName
    ApplyDVByName ws.Range("B5"), listName
    ApplyDVByName ws.Range("B6"), listName

    'Button-like cell (simple)
    ws.Range("D4").Value = "Refresh"
    ws.Range("D4").Font.Bold = True
    ws.Range("D4").Interior.ColorIndex = 15
    ws.Range("D4").HorizontalAlignment = xlCenter
    ws.Range("D4").Borders.LineStyle = xlContinuous

    ws.Range("A11").Value = "Summary"
    ws.Range("A11").Font.Bold = True

    ws.Range("A15").Value = "Output"
    ws.Range("A15").Font.Bold = True

    ws.Columns("A:D").AutoFit

    MsgBox "Reading sheet created: " & READ_SHEET & vbCrLf & _
           "To refresh: run RefreshReadingView (or assign it to a button).", vbInformation
End Sub

'========================
' 2) Build the report
'========================
Public Sub RefreshReadingView()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet, wsDB As Worksheet
    Dim loDB As ListObject

    Set ws = wb.Worksheets(READ_SHEET)
    Set wsDB = wb.Worksheets(DB_SHEET)
    Set loDB = wsDB.ListObjects(DB_TABLE)

    'Read user inputs
    Dim products(1 To MAX_PRODUCTS) As String
    products(1) = Trim$(CStr(ws.Range("B4").Value))
    products(2) = Trim$(CStr(ws.Range("B5").Value))
    products(3) = Trim$(CStr(ws.Range("B6").Value))

    Dim viewMode As String, layoutMode As String
    viewMode = LCase$(Trim$(CStr(ws.Range("B8").Value)))          'monthly/yearly
    layoutMode = LCase$(Trim$(CStr(ws.Range("B9").Value)))        'chronological/original

    Dim hasAny As Boolean
    hasAny = (Len(products(1)) > 0 Or Len(products(2)) > 0 Or Len(products(3)) > 0)
    If Not hasAny Then
        MsgBox "Select at least Product 1.", vbExclamation
        Exit Sub
    End If

    'Clear previous output
    ws.Range("A12:Z9999").ClearContents
    ws.Range("A12:Z9999").Borders.LineStyle = xlNone

    'Pull DB into array
    If loDB.DataBodyRange Is Nothing Then
        MsgBox "tblDB is empty.", vbExclamation
        Exit Sub
    End If

    Dim dataArr As Variant, hdrArr As Variant
    dataArr = loDB.DataBodyRange.Value
    hdrArr = loDB.HeaderRowRange.Value

    'Build column index map
    Dim col As Object: Set col = CreateObject("Scripting.Dictionary")
    Dim j As Long
    For j = 1 To loDB.ListColumns.Count
        col(loDB.ListColumns(j).name) = j
    Next j

    'Required
    If Not col.Exists("Item") Or Not col.Exists("Year") Or Not col.Exists("Month") Then
        MsgBox "tblDB must include Item, Year, Month.", vbCritical
        Exit Sub
    End If

    'Metrics to display (feel free to add/remove)
    Dim metrics As Variant
    metrics = Array("Total Month", "Forecast CY", "Invoiced", "Difference forecast/sales", "Difference forecast/sales%")

    'Output cursor
    Dim outRow As Long: outRow = 12

    Dim p As Long
    For p = 1 To MAX_PRODUCTS
        If Len(products(p)) = 0 Then GoTo NextProduct

        'Filter rows for this product
        Dim rows As Collection: Set rows = New Collection
        Dim r As Long
        For r = 1 To UBound(dataArr, 1)
            Dim itemV As String, yearV As String, monthV As String
            itemV = Trim$(CStr(dataArr(r, col("Item"))))
            yearV = Trim$(CStr(dataArr(r, col("Year"))))
            monthV = Trim$(CStr(dataArr(r, col("Month"))))

            If UCase$(itemV) = UCase$(products(p)) Then
                If UCase$(yearV) <> "202X" And UCase$(monthV) <> "FY" Then
                    rows.Add r
                End If
            End If
        Next r

        If rows.Count = 0 Then
            ws.Cells(outRow, 1).Value = products(p) & " — no data found."
            outRow = outRow + 2
            GoTo NextProduct
        End If

        'Build periods list
        Dim periods As Variant
        If viewMode = "yearly" Then
            periods = BuildYearPeriods(dataArr, rows, col)
        Else
            periods = BuildMonthPeriods(dataArr, rows, col, layoutMode)
        End If

        'Title
        ws.Cells(outRow, 1).Value = products(p)
        ws.Cells(outRow, 1).Font.Bold = True
        outRow = outRow + 1

        'Write header row: Metric + periods
        ws.Cells(outRow, 1).Value = "Metric"
        ws.Cells(outRow, 1).Font.Bold = True

        For j = 0 To UBound(periods)
            ws.Cells(outRow, 2 + j).Value = periods(j)
            ws.Cells(outRow, 2 + j).Font.Bold = True
        Next j

        'Write metric rows
        Dim i As Long
        For i = 0 To UBound(metrics)
            ws.Cells(outRow + 1 + i, 1).Value = metrics(i)

            For j = 0 To UBound(periods)
                ws.Cells(outRow + 1 + i, 2 + j).Value = _
                    GetValueForPeriod(dataArr, rows, col, viewMode, layoutMode, metrics(i), periods(j))
            Next j
        Next i

        'Format block
        Dim blockTop As Long, blockLeft As Long, blockBottom As Long, blockRight As Long
        blockTop = outRow
        blockLeft = 1
        blockBottom = outRow + 1 + UBound(metrics)
        blockRight = 1 + 1 + UBound(periods)

        With ws.Range(ws.Cells(blockTop, blockLeft), ws.Cells(blockBottom, blockRight))
            .Borders.LineStyle = xlContinuous
            .Columns.AutoFit
        End With

        'Summary for this product above output (simple)
        WriteSummary ws, products(p), dataArr, rows, col, viewMode, outRow - 2

        outRow = blockBottom + 2

NextProduct:
    Next p

    ws.Columns.AutoFit
End Sub

'========================
' Value retrieval
'========================
Private Function GetValueForPeriod(ByVal dataArr As Variant, ByVal rows As Collection, ByVal col As Object, _
                                  ByVal viewMode As String, ByVal layoutMode As String, _
                                  ByVal metricName As String, ByVal periodLabel As String) As Variant
    If Not col.Exists(metricName) Then
        GetValueForPeriod = ""
        Exit Function
    End If

    Dim yearWanted As String, monthWanted As String
    If viewMode = "yearly" Then
        yearWanted = periodLabel
        GetValueForPeriod = AggregateYear(dataArr, rows, col, metricName, yearWanted)
        Exit Function
    End If

    'Monthly: period label like "2024-Jan"
    yearWanted = Split(periodLabel, "-")(0)
    monthWanted = Split(periodLabel, "-")(1)

    Dim r As Long, idx As Long
    For idx = 1 To rows.Count
        r = rows(idx)
        If Trim$(CStr(dataArr(r, col("Year")))) = yearWanted And _
           UCase$(Left$(Trim$(CStr(dataArr(r, col("Month")))), 3)) = UCase$(Left$(monthWanted, 3)) Then

            GetValueForPeriod = dataArr(r, col(metricName))
            Exit Function
        End If
    Next idx

    GetValueForPeriod = ""
End Function

Private Function AggregateYear(ByVal dataArr As Variant, ByVal rows As Collection, ByVal col As Object, _
                               ByVal metricName As String, ByVal yearWanted As String) As Variant
    'Sum numeric metrics for that year.
    'For % metric: compute (Sum(Diff)/Sum(Invoiced)) with NA on div0.
    Dim sumV As Double, sumDiff As Double, sumInv As Double
    Dim hasNum As Boolean: hasNum = False

    Dim idx As Long, r As Long, y As String
    For idx = 1 To rows.Count
        r = rows(idx)
        y = Trim$(CStr(dataArr(r, col("Year"))))
        If y = yearWanted Then
            If LCase$(metricName) = LCase$("Difference forecast/sales%") Then
                If col.Exists("Difference forecast/sales") And col.Exists("Invoiced") Then
                    If IsNumeric(dataArr(r, col("Difference forecast/sales"))) Then sumDiff = sumDiff + CDbl(dataArr(r, col("Difference forecast/sales")))
                    If IsNumeric(dataArr(r, col("Invoiced"))) Then sumInv = sumInv + CDbl(dataArr(r, col("Invoiced")))
                End If
            Else
                If IsNumeric(dataArr(r, col(metricName))) Then
                    sumV = sumV + CDbl(dataArr(r, col(metricName)))
                    hasNum = True
                End If
            End If
        End If
    Next idx

    If LCase$(metricName) = LCase$("Difference forecast/sales%") Then
        If sumInv = 0 Then
            AggregateYear = CVErr(xlErrNA)
        Else
            AggregateYear = sumDiff / sumInv
        End If
    Else
        If hasNum Then AggregateYear = sumV Else AggregateYear = ""
    End If
End Function

'========================
' Period builders
'========================
Private Function BuildYearPeriods(ByVal dataArr As Variant, ByVal rows As Collection, ByVal col As Object) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim idx As Long, r As Long, y As String
    For idx = 1 To rows.Count
        r = rows(idx)
        y = Trim$(CStr(dataArr(r, col("Year"))))
        If Len(y) > 0 And UCase$(y) <> "202X" Then dict(y) = True
    Next idx

    Dim arr As Variant: arr = dict.keys
    SortStringsNumeric arr
    BuildYearPeriods = arr
End Function

Private Function BuildMonthPeriods(ByVal dataArr As Variant, ByVal rows As Collection, ByVal col As Object, _
                                   ByVal layoutMode As String) As Variant
    'Return array like "2024-Jan"
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim idx As Long, r As Long, y As String, m As String, key As String

    For idx = 1 To rows.Count
        r = rows(idx)
        y = Trim$(CStr(dataArr(r, col("Year"))))
        m = Trim$(CStr(dataArr(r, col("Month"))))
        If Len(y) = 0 Or Len(m) = 0 Then GoTo NextR
        If UCase$(y) = "202X" Or UCase$(m) = "FY" Then GoTo NextR

        key = y & "-" & Month3(m)
        dict(key) = True
NextR:
    Next idx

    Dim arr As Variant: arr = dict.keys
    If LCase$(layoutMode) = "original" Then
        SortPeriodsOriginal arr
    Else
        SortPeriodsChrono arr
    End If

    BuildMonthPeriods = arr
End Function

'Chronological sort: Year asc, Month asc
Private Sub SortPeriodsChrono(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If PeriodKeyChrono(arr(j)) < PeriodKeyChrono(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

'Original sort: Month group (Jan across years), then Feb across years...
Private Sub SortPeriodsOriginal(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If PeriodKeyOriginal(arr(j)) < PeriodKeyOriginal(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function PeriodKeyChrono(ByVal periodLabel As String) As Long
    Dim y As Long, m As Long
    y = CLng(Split(periodLabel, "-")(0))
    m = MonthNumFrom3(Split(periodLabel, "-")(1))
    PeriodKeyChrono = y * 100 + m
End Function

Private Function PeriodKeyOriginal(ByVal periodLabel As String) As Long
    Dim y As Long, m As Long
    y = CLng(Split(periodLabel, "-")(0))
    m = MonthNumFrom3(Split(periodLabel, "-")(1))
    'Month is primary, year secondary
    PeriodKeyOriginal = m * 10000 + y
End Function

'========================
' Summary
'========================
Private Sub WriteSummary(ByVal ws As Worksheet, ByVal product As String, ByVal dataArr As Variant, ByVal rows As Collection, _
                         ByVal col As Object, ByVal viewMode As String, ByVal anchorRow As Long)
    'Writes a small summary block at top (A12..), stacked
    Static nextSummaryRow As Long
    If nextSummaryRow = 0 Then nextSummaryRow = 12

    Dim r0 As Long: r0 = nextSummaryRow
    ws.Cells(r0, 1).Value = product
    ws.Cells(r0, 1).Font.Bold = True

    Dim latestLabel As String, prevLabel As String
    If LCase$(viewMode) = "yearly" Then
        Dim years As Variant: years = BuildYearPeriods(dataArr, rows, col)
        latestLabel = years(UBound(years))
        If UBound(years) >= 1 Then prevLabel = years(UBound(years) - 1) Else prevLabel = ""
    Else
        Dim periods As Variant: periods = BuildMonthPeriods(dataArr, rows, col, "chronological")
        latestLabel = periods(UBound(periods))
        'YoY same month previous year if exists
        prevLabel = CStr(CLng(Split(latestLabel, "-")(0)) - 1) & "-" & Split(latestLabel, "-")(1)
    End If

    ws.Cells(r0 + 1, 1).Value = "Latest period"
    ws.Cells(r0 + 1, 2).Value = latestLabel

    ws.Cells(r0 + 2, 1).Value = "Total Month (latest)"
    ws.Cells(r0 + 2, 2).Value = GetValueForPeriod(dataArr, rows, col, viewMode, "chronological", "Total Month", latestLabel)

    If Len(prevLabel) > 0 Then
        Dim vLatest As Variant, vPrev As Variant
        vLatest = GetValueForPeriod(dataArr, rows, col, viewMode, "chronological", "Total Month", latestLabel)
        vPrev = GetValueForPeriod(dataArr, rows, col, viewMode, "chronological", "Total Month", prevLabel)

        ws.Cells(r0 + 3, 1).Value = "YoY ? (Total Month)"
        If IsNumeric(vLatest) And IsNumeric(vPrev) Then
            ws.Cells(r0 + 3, 2).Value = CDbl(vLatest) - CDbl(vPrev)
        Else
            ws.Cells(r0 + 3, 2).Value = ""
        End If
    End If

    ws.Range(ws.Cells(r0, 1), ws.Cells(r0 + 3, 2)).Borders.LineStyle = xlContinuous
    ws.Columns("A:B").AutoFit

    nextSummaryRow = r0 + 5
End Sub

'========================
' Small utilities
'========================
Private Function Month3(ByVal m As String) As String
    Month3 = UCase$(Left$(Trim$(m), 3))
End Function

Private Function MonthNumFrom3(ByVal m3 As String) As Long
    Select Case UCase$(Left$(m3, 3))
        Case "JAN": MonthNumFrom3 = 1
        Case "FEB": MonthNumFrom3 = 2
        Case "MAR": MonthNumFrom3 = 3
        Case "APR": MonthNumFrom3 = 4
        Case "MAY": MonthNumFrom3 = 5
        Case "JUN": MonthNumFrom3 = 6
        Case "JUL": MonthNumFrom3 = 7
        Case "AUG": MonthNumFrom3 = 8
        Case "SEP": MonthNumFrom3 = 9
        Case "OCT": MonthNumFrom3 = 10
        Case "NOV": MonthNumFrom3 = 11
        Case "DEC": MonthNumFrom3 = 12
        Case Else: MonthNumFrom3 = 0
    End Select
End Function

Private Sub SortStringsNumeric(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(j)) < CLng(arr(i)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function UniqueListFromTable(ByVal lo As ListObject, ByVal colName As String) As Variant
    Dim idx As Long: idx = lo.ListColumns(colName).Index
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long, v As String
    For r = 1 To lo.DataBodyRange.rows.Count
        v = Trim$(CStr(lo.DataBodyRange.Cells(r, idx).Value))
        If Len(v) > 0 Then dict(v) = True
    Next r

    UniqueListFromTable = dict.keys
End Function

Private Sub WriteVerticalList(ByVal ws As Worksheet, ByVal startCell As Range, ByVal arr As Variant)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        startCell.Offset(i - LBound(arr), 0).Value = arr(i)
    Next i
End Sub

Private Sub ApplyDVByName(ByVal target As Range, ByVal listName As String)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, Formula1:="=" & listName
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub


Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.name = name
    End If
End Function

Private Function EnsureItemsListName(ByVal loDB As ListObject) As String
    'Creates/refreshes a very hidden sheet "_Lists" with unique items
    'and creates/updates a workbook-level named range "ItemsList".
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsL As Worksheet
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, v As String
    
    'Build unique list from tblDB[Item]
    If loDB.DataBodyRange Is Nothing Then
        EnsureItemsListName = "ItemsList"
        Exit Function
    End If
    
    Dim itemCol As Long: itemCol = loDB.ListColumns("Item").Index
    For r = 1 To loDB.DataBodyRange.rows.Count
        v = Trim$(CStr(loDB.DataBodyRange.Cells(r, itemCol).Value))
        If Len(v) > 0 Then
            If Not dict.Exists(UCase$(v)) Then dict.Add UCase$(v), v
        End If
    Next r
    
    'Create/get hidden list sheet
    Set wsL = GetOrCreateSheet(wb, "_Lists")
    wsL.Cells.Clear
    wsL.Visible = xlSheetVeryHidden
    
    'Write list to column A
    Dim i As Long, keys As Variant
    keys = dict.keys
    For i = 0 To dict.Count - 1
        wsL.Cells(i + 1, 1).Value = dict(keys(i))
    Next i
    
    'Define named range ItemsList to that range
    Dim rng As Range
    If dict.Count = 0 Then
        Set rng = wsL.Range("A1")
        rng.Value = ""
    Else
        Set rng = wsL.Range("A1").Resize(dict.Count, 1)
    End If
    
    On Error Resume Next
    wb.Names("ItemsList").Delete
    On Error GoTo 0
    
    wb.Names.Add name:="ItemsList", RefersTo:="=" & wsL.name & "!" & rng.Address
    
    EnsureItemsListName = "ItemsList"
End Function