Attribute VB_Name = "InventoryDashboard"
Option Explicit

'===========================================================================
' InventoryDashboard
' Builds a fully automated Inventory Performance Dashboard from a "Data"
' worksheet containing: Month, Category, Product, Supplier, Revenue,
' Profit, Expenses, Stock In, Stock Out
'===========================================================================

Private Const C_NAVY As Long = 8200479
Private Const C_BLUE As Long = 12874308
Private Const C_BG As Long = 16512756
Private Const C_WHITE As Long = 16777215

'===========================================================================
' 2. MAIN ENTRY POINT
'===========================================================================
Public Sub CreateInventoryDashboard()

    Dim wb As Workbook
    Dim dataWS As Worksheet
    Dim pivotWS As Worksheet
    Dim dashWS As Worksheet
    Dim pc As PivotCache
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable
    Dim pt4 As PivotTable, pt5 As PivotTable
    Dim sourceRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim nextCol As Long
    Dim headerCell As Range

    Dim savedScreenUpdating As Boolean
    Dim savedDisplayAlerts As Boolean
    Dim savedCalculation As XlCalculation

    savedScreenUpdating = Application.ScreenUpdating
    savedDisplayAlerts = Application.DisplayAlerts
    savedCalculation = Application.Calculation

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wb = ThisWorkbook

    ' Step 1 - Get the Data worksheet
    Set dataWS = GetWorksheet(wb, "Data")
    If dataWS Is Nothing Then
        Err.Raise vbObjectError + 100, "CreateInventoryDashboard", _
            "Required worksheet named 'Data' was not found in this workbook."
    End If

    ' Step 2 - Validate headers
    Call ValidateHeaders(dataWS)

    ' Step 3 - Remove any existing dashboard objects
    Call DeleteDashboardObjects(wb)

    ' Step 4 - Find last used row / column
    lastRow = LastUsedRow(dataWS)
    lastCol = LastUsedCol(dataWS)
    If lastRow < 2 Then
        Err.Raise vbObjectError + 101, "CreateInventoryDashboard", _
            "The 'Data' worksheet must contain at least one data row below the header row."
    End If

    ' Step 5 - Define source range
    Set sourceRange = dataWS.Range(dataWS.Cells(1, 1), dataWS.Cells(lastRow, lastCol))

    ' Step 6 - Add Pivot sheet
    Set pivotWS = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    pivotWS.Name = "Pivot"
    pivotWS.Tab.Color = RGB(68, 114, 196)
    pivotWS.Visible = xlSheetVisible

    ' Step 7 - Create PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=dataWS.Name & "!" & sourceRange.Address(ReferenceStyle:=xlR1C1))

    ' Step 8 - Create five pivot tables
    Set pt1 = CreatePT(pc, pivotWS, pivotWS.Range("A2"), "PT_Monthly")
    Call AddPTRowField(pt1, "Month")
    Call SafeAddDataField(pt1, "Revenue", "$#,##0", "Total Revenue")
    Call SafeAddDataField(pt1, "Profit", "$#,##0", "Total Profit")
    Call SafeAddDataField(pt1, "Expenses", "$#,##0", "Total Expenses")
    Call SafeAddDataField(pt1, "Stock Out", "#,##0", "UnitsSold")
    Set headerCell = pivotWS.Range("A1")
    Call FinishPT(pt1, headerCell, "Monthly Performance")

    nextCol = SafeNextCol(pt1)
    Set pt2 = CreatePT(pc, pivotWS, pivotWS.Cells(2, nextCol), "PT_Category")
    Call AddPTRowField(pt2, "Category")
    Call SafeAddDataField(pt2, "Revenue", "$#,##0", "Total Revenue")
    Call SafeAddDataField(pt2, "Profit", "$#,##0", "Total Profit")
    Set headerCell = pivotWS.Cells(1, nextCol)
    Call FinishPT(pt2, headerCell, "Category Performance")

    nextCol = SafeNextCol(pt2)
    Set pt3 = CreatePT(pc, pivotWS, pivotWS.Cells(2, nextCol), "PT_Product")
    Call AddPTRowField(pt3, "Product")
    Call SafeAddDataField(pt3, "Revenue", "$#,##0", "Total Revenue")
    Call SafeAddDataField(pt3, "Profit", "$#,##0", "Total Profit")
    Call SafeAddDataField(pt3, "Stock Out", "#,##0", "UnitsSold")
    Set headerCell = pivotWS.Cells(1, nextCol)
    Call FinishPT(pt3, headerCell, "Product Performance")

    nextCol = SafeNextCol(pt3)
    Set pt4 = CreatePT(pc, pivotWS, pivotWS.Cells(2, nextCol), "PT_Supplier")
    Call AddPTRowField(pt4, "Supplier")
    Call SafeAddDataField(pt4, "Revenue", "$#,##0", "Total Revenue")
    Call SafeAddDataField(pt4, "Profit", "$#,##0", "Total Profit")
    Set headerCell = pivotWS.Cells(1, nextCol)
    Call FinishPT(pt4, headerCell, "Supplier Performance")

    nextCol = SafeNextCol(pt4)
    Set pt5 = CreatePT(pc, pivotWS, pivotWS.Cells(2, nextCol), "PT_StockFlow")
    Call AddPTRowField(pt5, "Month")
    Call SafeAddDataField(pt5, "Stock In", "#,##0", "StockIn")
    Call SafeAddDataField(pt5, "Stock Out", "#,##0", "StockOut")
    Set headerCell = pivotWS.Cells(1, nextCol)
    Call FinishPT(pt5, headerCell, "Stock Flow")

    ' Step 9 - Autofit and format the pivot sheet
    pivotWS.Columns.AutoFit
    Call FormatPivotSheet(pivotWS)

    ' Step 10 - Add Dashboard sheet
    Set dashWS = wb.Worksheets.Add(After:=pivotWS)
    dashWS.Name = "Dashboard"
    dashWS.Tab.Color = RGB(31, 73, 125)

    ' Step 11 - Build the dashboard
    Call BuildDashboardLayout(dashWS)
    Call BuildKpis(dashWS, pt1)
    Call BuildCharts(dashWS, pt1, pt2, pt3, pt4, pt5)
    Call BuildSlicers(wb, dashWS, pt1, pt2, pt3, pt4, pt5)

    ' Step 12 - Activate Dashboard and finalize view
    dashWS.Activate
    dashWS.Range("A1").Select
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 70

    ' Step 13 - Restore calculation and recalc
    Application.Calculation = xlCalculationAutomatic
    Call CalculateFull

CleanupWithMessage:
    Application.ScreenUpdating = savedScreenUpdating
    Application.DisplayAlerts = savedDisplayAlerts
    Application.Calculation = savedCalculation
    Exit Sub

ErrHandler:
    Dim msg As String
    msg = "Inventory Dashboard build failed." & vbCrLf & vbCrLf & _
          "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
          "The 'Data' worksheet must contain these column headers in row 1:" & vbCrLf & _
          "Month, Category, Product, Supplier, Revenue, Profit, Expenses, Stock In, Stock Out"

    If Application.Visible = False Then
        Call StoreBuildStatus(msg)
    Else
        MsgBox msg, vbCritical, "Inventory Dashboard"
    End If

    Resume CleanupWithMessage

End Sub

Private Sub CalculateFull()
    On Error Resume Next
    Application.CalculateFullRebuild
    On Error GoTo 0
End Sub

'===========================================================================
' 4. HELPER FUNCTIONS & SUBS
'===========================================================================

Private Sub FormatPivotSheet(ws As Worksheet)
    On Error Resume Next

    ws.Range("A1:E1").Merge
    ws.Range("H1:J1").Merge
    ws.Range("M1:P1").Merge
    ws.Range("S1:U1").Merge
    ws.Range("X1:Z1").Merge

    ws.Columns(23).Delete   ' W
    ws.Columns(18).Delete   ' R
    ws.Columns(12).Delete   ' L
    ws.Columns(7).Delete    ' G

    On Error GoTo 0
End Sub

Private Sub ValidateHeaders(ws As Worksheet)
    Dim requiredHeaders As Variant
    requiredHeaders = Array("Month", "Category", "Product", "Supplier", _
        "Revenue", "Profit", "Expenses", "Stock In", "Stock Out")

    Dim headerRow As Range
    Set headerRow = ws.Rows(1)

    Dim missing As String
    Dim i As Long
    Dim m As Variant

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)
        m = Application.Match(requiredHeaders(i), headerRow, 0)
        If IsError(m) Then
            missing = missing & requiredHeaders(i) & ", "
        End If
    Next i

    If Len(missing) > 0 Then
        missing = Left(missing, Len(missing) - 2)
        Err.Raise vbObjectError + 102, "ValidateHeaders", _
            "The 'Data' worksheet is missing required column header(s): " & missing & _
            ". Required headers are: Month, Category, Product, Supplier, Revenue, Profit, Expenses, Stock In, Stock Out."
    End If
End Sub

Private Sub DeleteDashboardObjects(wb As Workbook)
    Dim sliceCacheNames As Variant
    sliceCacheNames = Array("SC_Month", "SC_Category", "SC_Product", "SC_Supplier")

    Dim i As Long
    On Error Resume Next
    For i = LBound(sliceCacheNames) To UBound(sliceCacheNames)
        wb.SlicerCaches(sliceCacheNames(i)).Delete
    Next i
    On Error GoTo 0

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    On Error Resume Next
    wb.Worksheets("Pivot").Delete
    wb.Worksheets("Dashboard").Delete
    On Error GoTo 0

    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub StoreBuildStatus(statusText As String)
    On Error Resume Next
    ThisWorkbook.Names("__DashboardBuildStatus").Delete
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Names.Add Name:="__DashboardBuildStatus", _
        RefersTo:="=""" & Replace(statusText, """", "'") & """"
    On Error GoTo 0
End Sub

Private Function GetWorksheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If foundCell Is Nothing Then
        LastUsedRow = 0
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

Private Function LastUsedCol(ws As Worksheet) As Long
    Dim foundCell As Range
    On Error Resume Next
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If foundCell Is Nothing Then
        LastUsedCol = 0
    Else
        LastUsedCol = foundCell.Column
    End If
End Function

Private Function CreatePT(pc As PivotCache, ws As Worksheet, target As Range, ptName As String) As PivotTable
    Dim pt As PivotTable
    Set pt = pc.CreatePivotTable(TableDestination:=target, TableName:=ptName)

    pt.ManualUpdate = True
    On Error Resume Next
    pt.TableStyle2 = "PivotStyleMedium9"
    On Error GoTo 0
    pt.RowAxisLayout xlTabularRow
    pt.DisplayErrorString = True
    pt.ErrorString = ""

    Set CreatePT = pt
End Function

Private Sub AddPTRowField(pt As PivotTable, fieldName As String)
    Dim fld As PivotField
    Set fld = pt.PivotFields(fieldName)
    fld.Orientation = xlRowField
    fld.Position = 1
End Sub

Private Sub SafeAddDataField(pt As PivotTable, srcField As String, numFmt As String, displayName As String)
    Dim df As PivotField
    Set df = pt.AddDataField(pt.PivotFields(srcField), displayName, xlSum)
    df.NumberFormat = numFmt

    On Error Resume Next
    df.Caption = displayName
    If Err.Number <> 0 Then
        Err.Clear
        df.Caption = displayName & " "
    End If
    On Error GoTo 0
End Sub

Private Sub FinishPT(pt As PivotTable, headerCell As Range, title As String)
    pt.ManualUpdate = False
    pt.RefreshTable

    With headerCell
        .Value = title
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = C_WHITE
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Function SafeNextCol(pt As PivotTable) As Long
    SafeNextCol = pt.TableRange2.Column + pt.TableRange2.Columns.Count + 2
End Function

'===========================================================================
' 5. BuildDashboardLayout
'===========================================================================
Private Sub BuildDashboardLayout(ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.Interior.Color = RGB(244, 246, 251)

    Dim c As Long
    For c = 1 To 30
        ws.Columns(c).ColumnWidth = 7.2
    Next c

    ws.Rows(1).RowHeight = 54
    ws.Rows(2).RowHeight = 8
    ws.Rows(3).RowHeight = 26
    ws.Rows(4).RowHeight = 34
    ws.Rows(5).RowHeight = 8

    Dim r As Long
    For r = 6 To 35
        ws.Rows(r).RowHeight = 14
    Next r

    With ws.Range("A1:Z1")
        .Merge
        .Interior.Color = RGB(41, 98, 181)
    End With

    With ws.Range("A1")
        .Value = "INVENTORY PERFORMANCE DASHBOARD"
        .Font.Name = "Calibri"
        .Font.Size = 36
        .Font.Bold = True
        .Font.Color = C_WHITE
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A2:Z5").Interior.Color = RGB(244, 246, 251)

    With ws.Range("A1:Z35").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(190, 200, 215)
    End With
End Sub

'===========================================================================
' 6. BuildKpis
'===========================================================================
Private Sub BuildKpis(ws As Worksheet, pt As PivotTable)

    Dim titles(0 To 4) As String
    Dim formulas(0 To 4) As String
    Dim numFmts(0 To 4) As String
    Dim colors(0 To 4) As Long

    titles(0) = "REVENUE"
    titles(1) = "PROFIT"
    titles(2) = "UNITS SOLD"
    titles(3) = "EXPENSES"
    titles(4) = "MARGIN"

    numFmts(0) = "$#,##0"
    numFmts(1) = "$#,##0"
    numFmts(2) = "#,##0"
    numFmts(3) = "$#,##0"
    numFmts(4) = "0.0%"

    colors(0) = RGB(41, 98, 181)
    colors(1) = RGB(46, 160, 100)
    colors(2) = RGB(230, 130, 30)
    colors(3) = RGB(190, 60, 60)
    colors(4) = RGB(120, 80, 200)

    Dim ptAnchor As String
    ptAnchor = "'" & pt.Parent.Name & "'!" & pt.TableRange1.Cells(1, 1).Address(False, False)

    Dim revGP As String, profGP As String, unitsGP As String, expGP As String
    revGP = "GETPIVOTDATA(""Total Revenue""," & ptAnchor & ")"
    profGP = "GETPIVOTDATA(""Total Profit""," & ptAnchor & ")"
    unitsGP = "GETPIVOTDATA(""UnitsSold""," & ptAnchor & ")"
    expGP = "GETPIVOTDATA(""Total Expenses""," & ptAnchor & ")"

    formulas(0) = "=IFERROR(" & revGP & ",0)"
    formulas(1) = "=IFERROR(" & profGP & ",0)"
    formulas(2) = "=IFERROR(" & unitsGP & ",0)"
    formulas(3) = "=IFERROR(" & expGP & ",0)"
    formulas(4) = "=IFERROR(" & profGP & "/" & revGP & ",0)"

    Dim i As Long
    Dim startCol As Long
    Dim titleRange As Range, valueRange As Range

    For i = 0 To 4
        startCol = 2 + i * 5

        Set titleRange = ws.Range(ws.Cells(3, startCol), ws.Cells(3, startCol + 3))
        With titleRange
            .Merge
            .Value = titles(i)
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = C_WHITE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = colors(i)
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = colors(i)
            End With
        End With

        Set valueRange = ws.Range(ws.Cells(4, startCol), ws.Cells(4, startCol + 3))
        With valueRange
            .Merge
            .Formula = formulas(i)
            .NumberFormat = numFmts(i)
            .Font.Name = "Calibri"
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = colors(i)
            .Interior.Color = C_WHITE
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = colors(i)
            End With
        End With
    Next i

End Sub

'===========================================================================
' 7. BuildCharts
'===========================================================================
Private Sub BuildCharts(ws As Worksheet, pt1 As PivotTable, pt2 As PivotTable, _
                         pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable)

    Const GAP_X As Double = 6
    Const GAP_Y As Double = 8
    Const PAD_LEFT As Double = 2
    Const PAD_RIGHT As Double = 2
    Const PAD_TOP As Double = 2
    Const PAD_BOT As Double = 2

    Dim areaLeft As Double, areaRight As Double, areaTop As Double, areaBottom As Double
    areaLeft = ws.Cells(6, 4).Left + PAD_LEFT
    areaRight = ws.Cells(6, 26).Left + ws.Cells(6, 26).Width - PAD_RIGHT
    areaTop = ws.Cells(6, 1).Top + PAD_TOP
    areaBottom = ws.Cells(35, 1).Top + ws.Cells(35, 1).Height - PAD_BOT

    Dim totalWidth As Double, totalHeight As Double
    totalWidth = areaRight - areaLeft
    totalHeight = areaBottom - areaTop

    Dim cellWidth As Double, cellHeight As Double
    cellWidth = (totalWidth - 2 * GAP_X) / 3
    cellHeight = (totalHeight - GAP_Y) / 2

    Dim colLefts(1 To 3) As Double
    Dim rowTops(1 To 2) As Double

    colLefts(1) = areaLeft
    colLefts(2) = areaLeft + cellWidth + GAP_X
    colLefts(3) = areaLeft + 2 * (cellWidth + GAP_X)

    rowTops(1) = areaTop
    rowTops(2) = areaTop + cellHeight + GAP_Y

    Call AddChart(ws, "LINE", pt1, "Revenue & Profit", colLefts(1), rowTops(1), cellWidth, cellHeight)
    Call AddChart(ws, "PIE", pt2, "Revenue by Category", colLefts(2), rowTops(1), cellWidth, cellHeight)
    Call AddChart(ws, "BAR_PRODUCT", pt3, "Revenue by Product", colLefts(3), rowTops(1), cellWidth, cellHeight)
    Call AddChart(ws, "AREA", pt4, "Revenue by Supplier", colLefts(1), rowTops(2), cellWidth, cellHeight)
    Call AddChart(ws, "COMBO", pt5, "Stock In vs Stock Out", colLefts(2), rowTops(2), cellWidth, cellHeight)
    Call AddChart(ws, "STACKED", pt1, "Revenue vs Expenses", colLefts(3), rowTops(2), cellWidth, cellHeight)

End Sub

'===========================================================================
' 8. AddChart - Per-Chart Formatting Rules
'===========================================================================
Private Sub AddChart(ws As Worksheet, chartType As String, pt As PivotTable, title As String, _
                      leftPos As Double, topPos As Double, widthVal As Double, heightVal As Double)

    Dim chObj As ChartObject
    Set chObj = ws.ChartObjects.Add(leftPos, topPos, widthVal, heightVal)

    Dim ch As Chart
    Set ch = chObj.Chart
    ch.SetSourceData Source:=pt.TableRange1

    On Error Resume Next
    ch.ShowAllFieldButtons = False
    On Error GoTo 0

    ch.HasLegend = False

    Select Case chartType
        Case "LINE"
            ch.ChartType = xlLineMarkers
            Call DeleteSeriesContaining(ch, "UnitsSold")
            Call DeleteSeriesContaining(ch, "Units Sold")
            Call RemoveHorizontalAxis(ch)

            On Error Resume Next
            ch.Axes(xlValue).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlValue).TickLabels.Font.Size = 12
            On Error GoTo 0

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
                .Width = 321.83
                .Left = 0.318
            End With

        Case "PIE"
            ch.ChartType = xlPie

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
            End With

            On Error Resume Next
            With ch.PlotArea
                .Left = 99.128
                .Top = 40.444
                .Width = 132.964
                .Height = 132.964
            End With
            On Error GoTo 0

            With ch.SeriesCollection(1)
                .HasDataLabels = True
                With .DataLabels
                    .ShowPercentage = True
                    .ShowValue = False
                    .ShowCategoryName = False
                    .ShowSeriesName = False
                    .Font.Name = "Calibri"
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Color = C_WHITE
                End With
            End With

        Case "BAR_PRODUCT"
            ch.ChartType = xlBarClustered
            Call DeleteSeriesContaining(ch, "UnitsSold")
            Call DeleteSeriesContaining(ch, "Units Sold")

            On Error Resume Next
            ch.ChartGroups(1).Overlap = 0
            ch.ChartGroups(1).GapWidth = 0
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlValue).Delete
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlCategory).TickLabels.Font.Size = 12
            On Error GoTo 0

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
            End With

        Case "AREA"
            ch.ChartType = xlArea

            On Error Resume Next
            ch.Axes(xlValue).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlValue).TickLabels.Font.Size = 12
            ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlCategory).TickLabels.Font.Size = 12
            On Error GoTo 0

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
            End With

        Case "COMBO"
            ch.ChartType = xlColumnClustered

            On Error Resume Next
            If ch.SeriesCollection.Count >= 2 Then
                ch.SeriesCollection(2).ChartType = xlLineMarkers
            End If
            On Error GoTo 0

            On Error Resume Next
            ch.SeriesCollection(1).ChartType = xlColumnClustered
            ch.ChartGroups(1).GapWidth = 97
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlValue).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlValue).TickLabels.Font.Size = 12
            ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlCategory).TickLabels.Font.Size = 12
            On Error GoTo 0

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
            End With

        Case "STACKED"
            ch.ChartType = xlColumnStacked
            Call DeleteSeriesContaining(ch, "UnitsSold")
            Call DeleteSeriesContaining(ch, "Units Sold")

            On Error Resume Next
            ch.ChartGroups(1).GapWidth = 100
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlValue).Delete
            On Error GoTo 0

            On Error Resume Next
            ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
            ch.Axes(xlCategory).TickLabels.Font.Size = 12
            On Error GoTo 0

            On Error Resume Next
            With ch.PlotArea
                .Left = 10
                .Top = 30.855
                .Width = ch.ChartArea.Width - 20
                .Height = 108.881
            End With
            On Error GoTo 0

            ch.HasLegend = True
            With ch.Legend
                .Position = xlLegendPositionBottom
                .Font.Name = "Calibri"
                .Font.Size = 12
                .Left = 0
                .Width = ch.ChartArea.Width
                .Height = 20
            End With

    End Select

    ' Common formatting for all chart types
    ch.HasTitle = True
    With ch.ChartTitle
        .Text = title
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Color = RGB(31, 73, 125)
        .Font.Size = 16
    End With

    With ch.ChartArea
        .Format.Fill.ForeColor.RGB = C_WHITE
        .Format.Line.ForeColor.RGB = RGB(31, 73, 125)
        .Format.Line.Weight = 1
    End With

    On Error Resume Next
    ch.PlotArea.Format.Fill.Visible = msoFalse
    On Error GoTo 0

    Call FormatSeries(ch, chartType)
    Call FormatAxes(ch)

End Sub

Private Sub DeleteSeriesContaining(ch As Chart, searchText As String)
    Dim i As Long
    On Error Resume Next
    For i = ch.SeriesCollection.Count To 1 Step -1
        If InStr(1, ch.SeriesCollection(i).Name, searchText, vbTextCompare) > 0 Then
            ch.SeriesCollection(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

'===========================================================================
' 9. SERIES COLOR PALETTE
'===========================================================================
Private Sub FormatSeries(ch As Chart, chartType As String)
    Dim palette(0 To 3) As Long
    palette(0) = RGB(41, 98, 181)
    palette(1) = RGB(46, 160, 100)
    palette(2) = RGB(190, 60, 60)
    palette(3) = RGB(230, 130, 30)

    Dim i As Long
    Dim colorIdx As Long
    Dim srs As Series

    On Error Resume Next

    Select Case chartType
        Case "PIE"
            Dim pt As Long
            For pt = 1 To ch.SeriesCollection(1).Points.Count
                colorIdx = (pt - 1) Mod 4
                ch.SeriesCollection(1).Points(pt).Format.Fill.ForeColor.RGB = palette(colorIdx)
            Next pt

        Case "LINE", "COMBO"
            For i = 1 To ch.SeriesCollection.Count
                colorIdx = (i - 1) Mod 4
                Set srs = ch.SeriesCollection(i)
                srs.Format.Line.ForeColor.RGB = palette(colorIdx)
                srs.Format.Line.Weight = 2
                srs.MarkerStyle = xlMarkerStyleCircle
                srs.MarkerSize = 5
                srs.MarkerForegroundColor = palette(colorIdx)
                srs.MarkerBackgroundColor = C_WHITE
            Next i

        Case Else
            For i = 1 To ch.SeriesCollection.Count
                colorIdx = (i - 1) Mod 4
                ch.SeriesCollection(i).Format.Fill.ForeColor.RGB = palette(colorIdx)
            Next i
    End Select

    On Error GoTo 0
End Sub

'===========================================================================
' 10. FormatAxes
'===========================================================================
Private Sub FormatAxes(ch As Chart)
    On Error Resume Next

    ch.Axes(xlValue).HasMajorGridlines = False
    ch.Axes(xlValue).HasMinorGridlines = False
    ch.Axes(xlValue).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlValue).TickLabels.Font.Size = 9

    ch.Axes(xlCategory).HasMajorGridlines = False
    ch.Axes(xlCategory).HasMinorGridlines = False
    ch.Axes(xlCategory).TickLabels.Font.Name = "Calibri"
    ch.Axes(xlCategory).TickLabels.Font.Size = 9

    On Error GoTo 0
End Sub

'===========================================================================
' 11. BuildSlicers
'===========================================================================
Private Sub BuildSlicers(wb As Workbook, dashWS As Worksheet, pt1 As PivotTable, _
                          pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable)

    Const GAP As Double = 4

    Dim railLeft As Double, railTop As Double, railWidth As Double, railHeight As Double
    railLeft = dashWS.Range("A6").Left
    railTop = dashWS.Range("A6").Top
    railWidth = dashWS.Range(dashWS.Cells(6, 1), dashWS.Cells(6, 3)).Width
    railHeight = dashWS.Range(dashWS.Cells(6, 1), dashWS.Cells(35, 1)).Height

    Dim slicerHeight As Double
    slicerHeight = (railHeight - 2 * GAP) / 3

    Dim fieldNames(0 To 2) As String
    Dim cacheNames(0 To 2) As String
    Dim sourcePTs(0 To 2) As Object

    fieldNames(0) = "Month"
    fieldNames(1) = "Supplier"
    fieldNames(2) = "Category"

    cacheNames(0) = "SC_Month"
    cacheNames(1) = "SC_Supplier"
    cacheNames(2) = "SC_Category"

    Set sourcePTs(0) = pt1
    Set sourcePTs(1) = pt4
    Set sourcePTs(2) = pt2

    Dim allPTs(0 To 4) As PivotTable
    Set allPTs(0) = pt1
    Set allPTs(1) = pt2
    Set allPTs(2) = pt3
    Set allPTs(3) = pt4
    Set allPTs(4) = pt5

    Dim i As Long, j As Long
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim topPos As Double

    For i = 0 To 2
        topPos = railTop + i * (slicerHeight + GAP)

        Set sc = wb.SlicerCaches.Add2(sourcePTs(i), fieldNames(i))
        sc.Name = cacheNames(i)

        For j = 0 To 4
            Call AddPivotToSlicer(sc, allPTs(j))
        Next j

        Set sl = sc.Slicers.Add(dashWS, , cacheNames(i), fieldNames(i), topPos, railLeft, slicerHeight, railWidth)

        On Error Resume Next
        sl.Style = "SlicerStyleDark1"
        If Err.Number <> 0 Then
            Err.Clear
            sl.Style = "SlicerStyleLight2"
        End If
        On Error GoTo 0

        sl.NumberOfColumns = 2
        sl.DisableMoveResizeUI = True

        Call FormatSlicerShape(sl, RGB(31, 73, 125), RGB(68, 114, 196))
    Next i

End Sub

Private Sub FormatSlicerShape(sl As Slicer, borderColor As Long, fillColor As Long)
    On Error Resume Next
    With sl.Shape
        .Fill.ForeColor.RGB = C_WHITE
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1
        .Shadow.Visible = msoFalse
    End With
    On Error GoTo 0
End Sub

Private Sub AddPivotToSlicer(cache As SlicerCache, pt As PivotTable)
    On Error Resume Next
    cache.PivotTables.AddPivotTable pt
    On Error GoTo 0
End Sub

'===========================================================================
' 13. ADDITIONAL STUBS
'===========================================================================
Private Sub RemoveHorizontalAxis(ch As Chart)
    On Error Resume Next
    ch.Axes(xlCategory).Delete
    On Error GoTo 0
End Sub

Private Sub ApplyChartTextSize(ch As Chart, fontSize As Long)
    On Error Resume Next

    ch.ChartTitle.Font.Size = fontSize
    ch.Legend.Font.Size = fontSize

    Dim i As Long
    For i = 1 To ch.SeriesCollection.Count
        If ch.SeriesCollection(i).HasDataLabels Then
            ch.SeriesCollection(i).DataLabels.Font.Size = fontSize
        End If
    Next i

    On Error GoTo 0
End Sub
