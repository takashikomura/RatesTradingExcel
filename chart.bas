Option Explicit

'============================================================
' JGB Multi Chart Controller (V12; V6 & V8 in-place migration)
' Data: selectable source sheet per series (A: Date, row 1: headers)
' Two source worksheets are selected independently for S1 and S2
' Time Series: calendar-date union of valid observations
' Scatter: date-matched intersection of valid observations
' Output: PlotData (four columns per chart)
' Dashboard: independent, movable control panels and charts
' MigrateExistingJGBCharts: upgrade V6 OR V8, preserve existing charts
' RepairJGBDropdowns: re-enable events and rebuild dropdowns
' SetupJGBChartController: initial setup/refresh existing charts
' DeleteChart: confirmation and complete per-chart cleanup
' Date axis: calendar-aligned month labels for time series
' Scatter: independent numeric XY axes
' Time-series Y-axis: data-driven min/max with 8% padding
'============================================================
' Only these three data worksheets are available as sources.
Private Const ALLOWED_SOURCE_1 As String = "JGBData"
Private Const ALLOWED_SOURCE_2 As String = "OISData"
Private Const ALLOWED_SOURCE_3 As String = "ASWData"

' Applied to NEW charts only. Migration never overwrites old selections.
Private Const DEFAULT_SOURCE_SHEET_1 As String = "JGBData"
Private Const DEFAULT_SOURCE_SHEET_2 As String = "OISData"
Private Const LIST_SHEET As String = "JGBChartLists"
Private Const DASH_SHEET As String = "Dashboard"
Private Const PLOT_SHEET As String = "PlotData"
Private Const CHART_PREFIX As String = "JGBChart_"
Private Const NAME_PREFIX As String = "JGB_"
Private Const HEADER_ROW As Long = 1
Private Const FIRST_ROW As Long = 2
Private Const BLOCK_WIDTH As Long = 4

'---------------- DISPLAY PARAMETERS ----------------
' Excel chart marker sizes: 2 to 72.
' These constants apply to both time-series and scatter charts.
Private Const NORMAL_MARKER_SIZE As Long = 4
Private Const LATEST_MARKER_SIZE As Long = 4
Private Const LINE_WEIGHT As Double = 1.5
'----------------------------------------------------

' Original Dashboard button names are retained.
Public Sub UpdateAll()
    RefreshAllJGBCharts
End Sub

Public Sub AddChart()
    AddJGBChart
End Sub

Public Sub RefreshJGBChart()
    RefreshAllJGBCharts
End Sub

Public Sub SetupJGBChartController()
    ' Fresh setup: create one chart. Existing V6/V8 workbook: migrate first.
    Dim ws As Worksheet, src1 As Worksheet, src2 As Worksheet
    Dim p As Range
    On Error GoTo Failed
    Set src1 = GetSourceSheet(DEFAULT_SOURCE_SHEET_1)
    Set src2 = GetSourceSheet(DEFAULT_SOURCE_SHEET_2)
    Set ws = EnsureSheet(DASH_SHEET)
    Call EnsureSheet(PLOT_SHEET)
    RefreshSourceSheetList
    EnsureDashboardButtons ws

    If RegisteredChartCount() = 0 Then
        ws.Activate
        Set p = FindFreePanelAnchor(ws)
        p.Select
        AddJGBChart
    Else
        ' Idempotent and non-destructive for V6 and V8 panels.
        MigrateExistingJGBCharts
        RefreshAllJGBCharts
    End If
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "SetupJGBChartController failed: " & Err.Description, vbExclamation
End Sub
Private Function FindFreePanelAnchor(ByVal ws As Worksheet) As Range
    Dim r As Long
    For r = 2 To ws.rows.Count - 20 Step 12
        If Application.CountA(ws.Cells(r, 1).Resize(8, 2)) = 0 Then
            Set FindFreePanelAnchor = ws.Cells(r, 1)
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 150, , "No empty space for a new chart panel."
End Function

' Buttons are created only if an equivalent macro button does not exist.
' Their positions can be freely moved by the user afterwards.
Private Sub EnsureDashboardButtons(ByVal ws As Worksheet)
    EnsureMacroButton ws, "UpdateAll", "UpdateAll", "T2", RGB(15, 92, 128)
    EnsureMacroButton ws, "AddChart", "AddChart", "T5", RGB(15, 92, 128)
    EnsureMacroButton ws, "DeleteChart", "DeleteChart", "T8", RGB(150, 45, 45)
End Sub

Private Sub EnsureMacroButton(ByVal ws As Worksheet, _
                              ByVal macroName As String, _
                              ByVal buttonText As String, _
                              ByVal anchor As String, _
                              ByVal fillColor As Long)
    Dim shp As Shape, existingAction As String
    For Each shp In ws.Shapes
        existingAction = vbNullString
        On Error Resume Next
        existingAction = shp.OnAction
        On Error GoTo 0
        If Len(existingAction) > 0 Then
            If LCase$(existingAction) = LCase$(macroName) Or _
               InStr(1, existingAction, "!" & macroName, vbTextCompare) > 0 Or _
               InStr(1, existingAction, "'!" & macroName, vbTextCompare) > 0 Then
                Exit Sub
            End If
        End If
    Next shp

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range(anchor).Left, ws.Range(anchor).Top, 105, 30)
    shp.Name = "JGB_Button_" & macroName
    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
    shp.TextFrame.Characters.text = buttonText
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.Visible = msoFalse
    shp.TextFrame.Characters.Font.Color = vbWhite
    shp.TextFrame.Characters.Font.Bold = True
    shp.Placement = xlFreeFloating
End Sub

Public Sub AddJGBChart()
    Dim ws As Worksheet, plotWs As Worksheet
    Dim srcWs1 As Worksheet, srcWs2 As Worksheet
    Dim p As Range, panel As Range, co As chartObject
    Dim id As String, s1 As String, s2 As String, errText As String
    Dim blockCol As Long
    Dim oldScreen As Boolean, oldEvents As Boolean

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo Failed

    Set ws = EnsureSheet(DASH_SHEET)
    Set plotWs = EnsureSheet(PLOT_SHEET)
    Set srcWs1 = GetSourceSheet(DEFAULT_SOURCE_SHEET_1)
    Set srcWs2 = GetSourceSheet(DEFAULT_SOURCE_SHEET_2)

    If Not ActiveSheet Is ws Then
        MsgBox "Dashboard上で配置開始セルを選択してください。", vbExclamation
        Exit Sub
    End If

    Set p = ActiveCell
    Set panel = p.Resize(8, 2)
    If Application.CountA(panel) > 0 Then
        MsgBox "設定パネルを置く8行×2列が空白ではありません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    RefreshSourceSheetList

    id = NextChartID(plotWs)
    blockCol = AllocatePlotBlock(plotWs, id)
    s1 = PreferredSeries(srcWs1, "2Y JGB", 2)
    s2 = PreferredSeries(srcWs2, "10Y JGB", 3)
    If StrComp(srcWs1.Name, srcWs2.Name, vbTextCompare) = 0 And _
       StrComp(s1, s2, vbTextCompare) = 0 Then s2 = vbNullString

    ' New panels use 8 rows: independent worksheet and header for each leg.
    p.value2 = "Chart ID"
    p.Offset(0, 1).value2 = id
    p.Offset(1, 0).value2 = "Chart Type"
    p.Offset(1, 1).value2 = "Time Series"
    p.Offset(2, 0).value2 = "Sheet 1 / X"
    p.Offset(2, 1).value2 = srcWs1.Name
    p.Offset(3, 0).value2 = "Series 1 / X"
    p.Offset(3, 1).value2 = s1
    p.Offset(4, 0).value2 = "Sheet 2 / Y"
    p.Offset(4, 1).value2 = srcWs2.Name
    p.Offset(5, 0).value2 = "Series 2 / Y"
    p.Offset(5, 1).value2 = s2
    p.Offset(6, 0).value2 = "Period"
    p.Offset(6, 1).value2 = "3Y"
    p.Offset(7, 0).value2 = "Status"
    p.Offset(7, 1).value2 = "Not updated"

    RegisterName id, "Panel", panel
    RegisterName id, "Type", p.Offset(1, 1)
    RegisterName id, "Sheet1", p.Offset(2, 1)
    RegisterName id, "S1", p.Offset(3, 1)
    RegisterName id, "Sheet2", p.Offset(4, 1)
    RegisterName id, "S2", p.Offset(5, 1)
    RegisterName id, "Period", p.Offset(6, 1)
    RegisterName id, "Status", p.Offset(7, 1)

    AddValidationList p.Offset(1, 1), "Time Series,Scatter", False
    AddValidationList p.Offset(6, 1), "1M,3M,6M,1Y,3Y,5Y,ALL", False
    AddSourceSheetValidation p.Offset(2, 1)
    AddSourceSheetValidation p.Offset(4, 1)
    AddSeriesValidationForChart p.Offset(3, 1), srcWs1, id, 1
    AddSeriesValidationForChart p.Offset(5, 1), srcWs2, id, 2

    With panel
        .Font.Name = "Arial"
        .Font.Size = 9
        .Interior.Color = RGB(245, 247, 250)
    End With
    With p.Resize(1, 2)
        .Interior.Color = RGB(42, 57, 76)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    p.Offset(1, 1).Resize(6, 1).Interior.Color = vbWhite
    ' Neither column widths nor row heights are changed.

    Set co = ws.ChartObjects.Add(p.Left, p.Offset(10, 0).Top, 720, 360)
    co.Name = CHART_PREFIX & id
    co.Placement = xlFreeFloating

    If Not TryRefreshChart(id, blockCol, errText) Then
        MsgBox id & ": " & errText, vbExclamation
    End If

CleanExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    MsgBox "AddChart failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RefreshAllJGBCharts()
    Dim ws As Worksheet, c As Long, lastCol As Long
    Dim id As String, errText As String, errorSummary As String
    Dim successCount As Long, failCount As Long
    Dim oldScreen As Boolean, oldEvents As Boolean

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    RefreshSourceSheetList
    Set ws = EnsureSheet(PLOT_SHEET)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        id = Trim$(CStr(ws.Cells(1, c).value2))
        If IsChartID(id) Then
            errText = vbNullString
            If TryRefreshChart(id, c, errText) Then
                successCount = successCount + 1
            Else
                failCount = failCount + 1
                errorSummary = errorSummary & id & ": " & errText & vbCrLf
            End If
        End If
    Next c

CleanExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    If failCount > 0 Then
        MsgBox successCount & " charts updated." & vbCrLf & _
               failCount & " charts failed." & vbCrLf & errorSummary, vbExclamation
    End If
    Exit Sub
FatalError:
    failCount = failCount + 1
    errorSummary = errorSummary & "Fatal error: " & Err.Description & vbCrLf
    Resume CleanExit
End Sub

Public Sub RefreshSelectedJGBChart()
    Dim ws As Worksheet, c As Long, lastCol As Long
    Dim id As String, errText As String
    RefreshSourceSheetList
    Set ws = EnsureSheet(PLOT_SHEET)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        id = Trim$(CStr(ws.Cells(1, c).value2))
        If IsChartID(id) Then
            If SelectionBelongsToChart(id) Then
                If Not TryRefreshChart(id, c, errText) Then
                    MsgBox id & ": " & errText, vbExclamation
                End If
                Exit Sub
            End If
        End If
    Next c
    MsgBox "対象チャートの設定セルを選択してください。", vbInformation
End Sub

Private Function TryRefreshChart(ByVal id As String, _
                                 ByVal blockCol As Long, _
                                 ByRef errText As String) As Boolean
    On Error GoTo Failed
    RefreshOneChart id, blockCol
    TryRefreshChart = True
    Exit Function
Failed:
    errText = Err.Description
    On Error Resume Next
    NamedCell(id, "Status").value2 = "ERROR: " & Left$(errText, 100)
    On Error GoTo 0
End Function

Private Sub RefreshOneChart(ByVal id As String, ByVal blockCol As Long)
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim dst As Worksheet, dash As Worksheet
    Dim chartType As String, s1 As String, s2 As String, periodKey As String
    Dim sheet1 As String, sheet2 As String, label1 As String, label2 As String
    Dim m1 As Object, m2 As Object, dateSet As Object
    Dim dateKeys() As Long, buffer() As Variant, output() As Variant
    Dim k As Variant, keyText As String
    Dim hasSecond As Boolean
    Dim n As Long, i As Long, j As Long, outCount As Long
    Dim earliest As Date, latest As Date, startDate As Date
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
    Dim v1 As Double, v2 As Double

    ' Upgrade legacy 6-row panels without moving old parameter cells.
    EnsureSourceFields id

    chartType = UCase$(Trim$(CStr(NamedCell(id, "Type").value2)))
    s1 = Trim$(CStr(NamedCell(id, "S1").value2))
    s2 = Trim$(CStr(NamedCell(id, "S2").value2))
    sheet1 = Trim$(CStr(NamedCell(id, "Sheet1").value2))
    sheet2 = Trim$(CStr(NamedCell(id, "Sheet2").value2))
    periodKey = UCase$(Trim$(CStr(NamedCell(id, "Period").value2)))

    If chartType <> "TIME SERIES" And chartType <> "SCATTER" Then
        Err.Raise vbObjectError + 101, , "Chart Type must be Time Series or Scatter."
    End If
    If Len(s1) = 0 Then Err.Raise vbObjectError + 102, , "Series 1 header is required."
    If Len(sheet1) = 0 Then Err.Raise vbObjectError + 103, , "Sheet 1 is blank."

    hasSecond = (Len(s2) > 0)
    If chartType = "SCATTER" And Not hasSecond Then
        Err.Raise vbObjectError + 104, , "Scatter requires Series 2 / Y."
    End If
    If hasSecond And Len(sheet2) = 0 Then
        Err.Raise vbObjectError + 105, , "Sheet 2 is blank."
    End If

    Set ws1 = GetSourceSheet(sheet1)
    If hasSecond Then Set ws2 = GetSourceSheet(sheet2)

    ' Refresh each header dropdown using its selected source sheet.
    AddSeriesValidationForChart NamedCell(id, "S1"), ws1, id, 1
    If hasSecond Then
        AddSeriesValidationForChart NamedCell(id, "S2"), ws2, id, 2
    ElseIf Len(sheet2) > 0 Then
        Set ws2 = GetSourceSheet(sheet2)
        AddSeriesValidationForChart NamedCell(id, "S2"), ws2, id, 2
    End If

    Set dst = ThisWorkbook.Worksheets(PLOT_SHEET)
    Set dash = ThisWorkbook.Worksheets(DASH_SHEET)
    Set m1 = ReadSeriesMap(ws1, s1)
    Set m2 = CreateObject("Scripting.Dictionary")
    If hasSecond Then Set m2 = ReadSeriesMap(ws2, s2)

    ' The union is used for Time Series; the intersection for Scatter.
    Set dateSet = CreateObject("Scripting.Dictionary")
    For Each k In m1.keys
        keyText = CStr(k)
        If chartType = "TIME SERIES" Then
            dateSet(keyText) = True
        ElseIf m2.Exists(keyText) Then
            dateSet(keyText) = True
        End If
    Next k
    If chartType = "TIME SERIES" And hasSecond Then
        For Each k In m2.keys
            dateSet(CStr(k)) = True
        Next k
    End If

    n = dateSet.Count
    If n = 0 Then
        If chartType = "SCATTER" Then
            Err.Raise vbObjectError + 106, , _
                "No same-date valid X/Y pairs between the two sheets."
        Else
            Err.Raise vbObjectError + 106, , "No numeric observations found."
        End If
    End If

    ' Sort dates chronologically before creating the PlotData array.
    ReDim dateKeys(1 To n)
    i = 0
    For Each k In dateSet.keys
        i = i + 1
        dateKeys(i) = CLng(k)
    Next k
    If n > 1 Then QuickSortDates dateKeys, 1, n

    earliest = CDate(dateKeys(1))
    latest = CDate(dateKeys(n))
    startDate = GetStartDate(latest, earliest, periodKey)

    ReDim buffer(1 To n, 1 To 3)
    For i = 1 To n
        If CDate(dateKeys(i)) >= startDate Then
            keyText = CStr(dateKeys(i))
            outCount = outCount + 1
            buffer(outCount, 1) = ToExcelSerial(CDate(dateKeys(i)))
            If m1.Exists(keyText) Then buffer(outCount, 2) = CDbl(m1(keyText))
            If hasSecond Then
                If m2.Exists(keyText) Then buffer(outCount, 3) = CDbl(m2(keyText))
            End If
        End If
    Next i
    If outCount = 0 Then
        Err.Raise vbObjectError + 107, , "No observations in selected period."
    End If

    ReDim output(1 To outCount, 1 To 3)
    For i = 1 To outCount
        For j = 1 To 3
            output(i, j) = buffer(i, j)
        Next j
    Next i

    If chartType = "SCATTER" Then
        ' All rows are valid date-matched pairs.
        xMin = CDbl(output(1, 2)): xMax = xMin
        yMin = CDbl(output(1, 3)): yMax = yMin
        For i = 2 To outCount
            v1 = CDbl(output(i, 2)): v2 = CDbl(output(i, 3))
            If v1 < xMin Then xMin = v1
            If v1 > xMax Then xMax = v1
            If v2 < yMin Then yMin = v2
            If v2 > yMax Then yMax = v2
        Next i
    End If

    WritePlotBlock dst, blockCol, output, outCount, s1, s2, startDate, latest
    UpdatePlotNames id, dst, blockCol, outCount

    label1 = s1
    label2 = s2
    If hasSecond Then
        If StrComp(ws1.Name, ws2.Name, vbTextCompare) <> 0 Then
            label1 = s1 & " (" & ws1.Name & ")"
            label2 = s2 & " (" & ws2.Name & ")"
        End If
    Else
        label1 = s1 & " (" & ws1.Name & ")"
    End If

    UpdateChartObject dash, dst, id, blockCol, chartType, label1, label2, _
                      hasSecond, startDate, latest, outCount, _
                      xMin, xMax, yMin, yMax

    NamedCell(id, "Status").value2 = "OK | " & outCount & " rows | " & _
        Format$(startDate, "yyyy/mm/dd") & " - " & Format$(latest, "yyyy/mm/dd")
End Sub

' Return valid numeric observations indexed by their actual calendar date.
' Duplicate dates: the last valid row in the source sheet wins.
Private Function ReadSeriesMap(ByVal ws As Worksheet, _
                               ByVal headerText As String) As Object
    Dim m As Object, dateData As Variant, values As Variant
    Dim lastRow As Long, seriesCol As Long, i As Long, n As Long
    Dim dt As Date, y As Double
    Set m = CreateObject("Scripting.Dictionary")
    seriesCol = FindHeaderColumn(ws, headerText)
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST_ROW Then
        Err.Raise vbObjectError + 108, , "No observations on source sheet: " & ws.Name
    End If
    n = lastRow - FIRST_ROW + 1
    dateData = ReadColumn(ws, 1, lastRow)
    values = ReadColumn(ws, seriesCol, lastRow)
    For i = 1 To n
        If ParseDateValue(dateData(i, 1), dt) Then
            If ParseYield(values(i, 1), y) Then
                m(CStr(CLng(dt))) = y
            End If
        End If
    Next i
    Set ReadSeriesMap = m
End Function

' QuickSort integer date keys without writing back to worksheet cells.
Private Sub QuickSortDates(ByRef a() As Long, _
                           ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, pivot As Long, temp As Long
    i = low: j = high
    pivot = a((low + high) \ 2)
    Do While i <= j
        Do While a(i) < pivot: i = i + 1: Loop
        Do While a(j) > pivot: j = j - 1: Loop
        If i <= j Then
            temp = a(i): a(i) = a(j): a(j) = temp
            i = i + 1: j = j - 1
        End If
    Loop
    If low < j Then QuickSortDates a, low, j
    If i < high Then QuickSortDates a, i, high
End Sub

Private Sub WritePlotBlock(ByVal ws As Worksheet, ByVal blockCol As Long, _
                           ByRef data() As Variant, ByVal rowCount As Long, _
                           ByVal s1 As String, ByVal s2 As String, _
                           ByVal startDate As Date, ByVal latest As Date)
    Dim lastOld As Long
    lastOld = ws.Cells(ws.rows.Count, blockCol).End(xlUp).Row
    If lastOld < 2 Then lastOld = 2
    ws.Range(ws.Cells(2, blockCol), ws.Cells(lastOld, blockCol + 2)).ClearContents

    ' Keep row 1 first cell as registry ID, row 1 next two cells as dates.
    ws.Cells(1, blockCol + 1).value2 = ToExcelSerial(startDate)
    ws.Cells(1, blockCol + 2).value2 = ToExcelSerial(latest)
    ws.Cells(1, blockCol + 1).Resize(1, 2).numberFormat = "yyyy/mm/dd"
    ws.Cells(2, blockCol).value2 = "Date"
    ws.Cells(2, blockCol + 1).value2 = s1
    ws.Cells(2, blockCol + 2).value2 = s2
    ws.Cells(3, blockCol).Resize(rowCount, 3).value2 = data

    If rowCount > 1 Then
        ws.Range(ws.Cells(2, blockCol), _
                 ws.Cells(rowCount + 2, blockCol + 2)).Sort _
                 Key1:=ws.Cells(3, blockCol), Order1:=xlAscending, Header:=xlYes
    End If
    ws.Cells(3, blockCol).Resize(rowCount, 1).numberFormat = "yyyy/mm/dd"
    ws.Cells(3, blockCol + 1).Resize(rowCount, 2).numberFormat = "0.000"
End Sub

' Configuration names end in _S1/_S2; data names use _Data_S1/_Data_S2.
Private Sub UpdatePlotNames(ByVal id As String, ByVal ws As Worksheet, _
                            ByVal blockCol As Long, ByVal rowCount As Long)
    Dim lastRow As Long
    lastRow = rowCount + 2
    UpsertWorkbookName NAME_PREFIX & id & "_Dates", _
        ws.Range(ws.Cells(3, blockCol), ws.Cells(lastRow, blockCol))
    UpsertWorkbookName NAME_PREFIX & id & "_Data_S1", _
        ws.Range(ws.Cells(3, blockCol + 1), ws.Cells(lastRow, blockCol + 1))
    UpsertWorkbookName NAME_PREFIX & id & "_Data_S2", _
        ws.Range(ws.Cells(3, blockCol + 2), ws.Cells(lastRow, blockCol + 2))
    UpsertWorkbookName NAME_PREFIX & id & "_S1_Label", ws.Cells(2, blockCol + 1)
    UpsertWorkbookName NAME_PREFIX & id & "_S2_Label", ws.Cells(2, blockCol + 2)
    UpsertWorkbookName NAME_PREFIX & id & "_StartDate", ws.Cells(1, blockCol + 1)
    UpsertWorkbookName NAME_PREFIX & id & "_EndDate", ws.Cells(1, blockCol + 2)
End Sub

Private Sub UpdateChartObject(ByVal dash As Worksheet, ByVal plotWs As Worksheet, _
                              ByVal id As String, ByVal blockCol As Long, _
                              ByVal chartType As String, ByVal s1 As String, _
                              ByVal s2 As String, ByVal hasSecond As Boolean, _
                              ByVal startDate As Date, ByVal latest As Date, _
                              ByVal rowCount As Long, ByVal xMin As Double, _
                              ByVal xMax As Double, ByVal yMin As Double, _
                              ByVal yMax As Double)
    Dim co As chartObject, ch As Chart, sr As series
    Dim p As Range, xRange As Range, yRange As Range
    Dim targetType As XlChartType, targetCount As Long
    Dim lastDataRow As Long, j As Long, lineColor As Long
    Dim latestIndex As Long, newChart As Boolean
    Dim xAxis As Axis, yAxis As Axis

    lastDataRow = rowCount + 2
    On Error Resume Next
    Set co = dash.ChartObjects(CHART_PREFIX & id)
    On Error GoTo 0
    If co Is Nothing Then
        Set p = NamedCell(id, "Panel")
        Set co = dash.ChartObjects.Add(p.Left, p.Offset(10, 0).Top, 720, 360)
        co.Name = CHART_PREFIX & id
        co.Placement = xlFreeFloating
        newChart = True
    End If
    Set ch = co.Chart

    ' Reset prior red point BEFORE rebinding and changing point count.
    If Not newChart Then ResetPreviousHighlights ch, plotWs, blockCol

    If chartType = "TIME SERIES" Then
        targetType = xlLineMarkers
        If hasSecond Then targetCount = 2 Else targetCount = 1
    Else
        targetType = xlXYScatter
        targetCount = 1
    End If

    With ch
        If .chartType <> targetType Then .chartType = targetType
        Do While .SeriesCollection.Count > targetCount
            .SeriesCollection(.SeriesCollection.Count).Delete
        Loop
        Do While .SeriesCollection.Count < targetCount
            .SeriesCollection.NewSeries
        Loop
        .PlotVisibleOnly = False
        .HasTitle = False

        If chartType = "TIME SERIES" Then
            .DisplayBlanksAs = xlInterpolated
            Set xRange = plotWs.Range(plotWs.Cells(3, blockCol), _
                                      plotWs.Cells(lastDataRow, blockCol))
            For j = 1 To targetCount
                Set sr = .SeriesCollection(j)
                Set yRange = plotWs.Range(plotWs.Cells(3, blockCol + j), _
                                          plotWs.Cells(lastDataRow, blockCol + j))
                lineColor = ChartSeriesColor(j)
                With sr
                    If .chartType <> xlLineMarkers Then .chartType = xlLineMarkers
                    .axisGroup = xlPrimary
                    If j = 1 Then .Name = s1 Else .Name = s2
                    .xValues = xRange
                    .values = yRange
                    With .Format.Line
                        .Visible = msoTrue
                        .ForeColor.RGB = lineColor
                        .Weight = LINE_WEIGHT
                        .DashStyle = msoLineSolid
                        .Transparency = 0
                    End With
                    ' Normal observations: visible markers (not None).
                    .MarkerStyle = xlMarkerStyleCircle
                    .markerSize = NORMAL_MARKER_SIZE
                    .MarkerForegroundColor = lineColor
                    .MarkerBackgroundColor = lineColor
                End With

                ' Only the latest valid observation gets a larger red marker.
                latestIndex = LastNumericPointIndex(plotWs, blockCol + j, rowCount)
                If latestIndex > 0 Then HighlightLatestPoint sr, latestIndex
                plotWs.Cells(j, blockCol + 3).value2 = latestIndex
            Next j
            If targetCount = 1 Then plotWs.Cells(2, blockCol + 3).ClearContents

            ' A true Excel date axis places month ticks at calendar boundaries.
            ' It also retains date-proportional spacing across holidays.
            ConfigureTimeSeriesDateAxis ch, startDate, latest
            ' Recalculate the time-series Y-axis from displayed data.
            ' Include both series when Series 2 is present.
            ConfigureTimeSeriesYAxis ch, plotWs, blockCol, _
                                     rowCount, targetCount

            ' Avoid point-by-point legend entries on single-series charts.
            If targetCount = 1 Then
                .HasLegend = False
                With .Axes(xlValue, xlPrimary)
                    .HasTitle = True
                    .AxisTitle.text = s1
                End With
            Else
                .HasLegend = True
                .Legend.Position = xlLegendPositionTop
                .Axes(xlValue, xlPrimary).HasTitle = False
            End If
        Else
            .DisplayBlanksAs = xlNotPlotted
            Set sr = .SeriesCollection(1)
            With sr
                If .chartType <> xlXYScatter Then .chartType = xlXYScatter
                .axisGroup = xlPrimary
                .Name = s2 & " vs " & s1
                .xValues = plotWs.Range(plotWs.Cells(3, blockCol + 1), _
                                          plotWs.Cells(lastDataRow, blockCol + 1))
                .values = plotWs.Range(plotWs.Cells(3, blockCol + 2), _
                                         plotWs.Cells(lastDataRow, blockCol + 2))
                .Format.Line.Visible = msoFalse
                .MarkerStyle = xlMarkerStyleCircle
                .markerSize = NORMAL_MARKER_SIZE
                .MarkerForegroundColor = ChartSeriesColor(1)
                .MarkerBackgroundColor = ChartSeriesColor(1)
            End With
            latestIndex = rowCount ' PlotData was sorted ascending by date.
            HighlightLatestPoint sr, latestIndex
            plotWs.Cells(1, blockCol + 3).value2 = latestIndex
            plotWs.Cells(2, blockCol + 3).ClearContents

            Set xAxis = ch.Axes(xlCategory, xlPrimary)
            ConfigureScatterAxis xAxis, xMin, xMax, s1
            Set yAxis = ch.Axes(xlValue, xlPrimary)
            ConfigureScatterAxis yAxis, yMin, yMax, s2
            .HasLegend = False
        End If
        If newChart Then
            .chartArea.Font.Name = "Arial"
            .chartArea.Font.Size = 8
        End If
    End With
    ' Existing chart position and dimensions remain unchanged.
End Sub

Private Function ChartSeriesColor(ByVal seriesIndex As Long) As Long
    If seriesIndex = 1 Then
        ChartSeriesColor = RGB(35, 105, 175)
    Else
        ChartSeriesColor = RGB(200, 90, 65)
    End If
End Function

Private Sub HighlightLatestPoint(ByVal sr As series, ByVal pointIndex As Long)
    If pointIndex <= 0 Or pointIndex > sr.Points.Count Then Exit Sub
    With sr.Points(pointIndex)
        .MarkerStyle = xlMarkerStyleCircle
        .markerSize = LATEST_MARKER_SIZE
        .MarkerForegroundColor = RGB(255, 0, 0)
        .MarkerBackgroundColor = RGB(255, 0, 0)
    End With
End Sub

Private Sub ResetPreviousHighlights(ByVal ch As Chart, _
                                    ByVal plotWs As Worksheet, _
                                    ByVal blockCol As Long)
    Dim j As Long, oldIndex As Long, v As Variant
    Dim sr As series, colorValue As Long
    On Error GoTo Done
    For j = 1 To 2
        If j > ch.SeriesCollection.Count Then Exit For
        v = plotWs.Cells(j, blockCol + 3).value2
        If IsNumeric(v) Then
            oldIndex = CLng(v)
            Set sr = ch.SeriesCollection(j)
            If oldIndex > 0 And oldIndex <= sr.Points.Count Then
                colorValue = ChartSeriesColor(j)
                With sr.Points(oldIndex)
                    ' Restore BOTH ordinary size and ordinary series color.
                    .MarkerStyle = xlMarkerStyleCircle
                    .markerSize = NORMAL_MARKER_SIZE
                    .MarkerForegroundColor = colorValue
                    .MarkerBackgroundColor = colorValue
                End With
            End If
        End If
    Next j
Done:
End Sub

Private Function LastNumericPointIndex(ByVal ws As Worksheet, _
                                       ByVal dataCol As Long, _
                                       ByVal rowCount As Long) As Long
    Dim data As Variant, v As Double, i As Long
    If rowCount <= 0 Then Exit Function
    data = ws.Cells(3, dataCol).Resize(rowCount, 1).value2
    If rowCount = 1 Then
        If ParseYield(data, v) Then LastNumericPointIndex = 1
        Exit Function
    End If
    For i = rowCount To 1 Step -1
        If ParseYield(data(i, 1), v) Then
            LastNumericPointIndex = i
            Exit Function
        End If
    Next i
End Function


'============================================================
' TIME SERIES DATE AXIS
' - Line chart with a true xlTimeScale category axis.
' - xlDays as the base unit preserves actual date spacing.
' - For >= 1 month, month ticks are anchored on day 1.
' - For short windows, weekly ticks show day and month.
'============================================================
Private Sub ConfigureTimeSeriesDateAxis(ByVal ch As Chart, _
                                        ByVal startDate As Date, _
                                        ByVal latestDate As Date)
    Dim ax As Axis
    Dim months As Long
    Dim firstMonth As Date, nextMonth As Date

    Set ax = ch.Axes(xlCategory, xlPrimary)
    firstMonth = dateSerial(Year(startDate), Month(startDate), 1)
    nextMonth = dateSerial(Year(latestDate), Month(latestDate) + 1, 1)
    months = DateDiff("m", startDate, latestDate)

    With ax
        .CategoryType = xlTimeScale
        .BaseUnit = xlDays
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True

        If months <= 1 Then
            .MinimumScale = ToExcelSerial(startDate)
            .MaximumScale = ToExcelSerial(latestDate) + 1#
            .MajorUnitScale = xlDays
            .majorUnit = 7
            .TickLabels.numberFormat = "m/d"
        Else
            .MinimumScale = ToExcelSerial(firstMonth)
            .MaximumScale = ToExcelSerial(nextMonth)
            .MajorUnitScale = xlMonths
            Select Case months
                Case Is <= 12
                    .majorUnit = 1
                Case Is <= 36
                    .majorUnit = 3
                Case Is <= 60
                    .majorUnit = 6
                Case Else
                    .majorUnit = 12
            End Select
            .TickLabels.numberFormat = "yyyy/mm"
        End If

        .HasMajorGridlines = True
        .TickLabelPosition = xlTickLabelPositionLow
        .HasTitle = True
        .AxisTitle.text = "Date"
    End With

    ' A negative/positive yield series should not move the date
    ' labels into the center of the plot around the zero line.
    With ch.Axes(xlValue, xlPrimary)
        .Crosses = xlMinimum
    End With
End Sub


'============================================================
' TIME SERIES Y-AXIS: DATA-DRIVEN SCALING
' - Uses the displayed PlotData rows, not all of JGBData.
' - Supports one or two time-series lines, missing values,
'   negative spreads, and nearly constant series.
' - Adds 8% padding to each side of the observed range.
' - Rounds boundaries and tick increments to nice numbers.
' - Leaves the Y-axis title to UpdateChartObject so that
'   the single-series and dual-series label logic is kept.
'============================================================
Private Sub ConfigureTimeSeriesYAxis( _
    ByVal ch As Chart, _
    ByVal ws As Worksheet, _
    ByVal blockCol As Long, _
    ByVal rowCount As Long, _
    ByVal seriesCount As Long)

    Dim seriesIndex As Long, i As Long
    Dim data As Variant, v As Variant
    Dim y As Double, dataMin As Double, dataMax As Double
    Dim span As Double, padding As Double
    Dim stepSize As Double, axisMin As Double, axisMax As Double
    Dim found As Boolean
    Dim ax As Axis

    If rowCount < 1 Or seriesCount < 1 Then Exit Sub

    ' Collect both Y series, excluding blank / invalid cells.
    For seriesIndex = 1 To seriesCount
        data = ws.Cells(3, blockCol + seriesIndex) _
                 .Resize(rowCount, 1).value2

        For i = 1 To rowCount
            If rowCount = 1 Then
                v = data
            Else
                v = data(i, 1)
            End If

            If ParseYield(v, y) Then
                If Not found Then
                    dataMin = y
                    dataMax = y
                    found = True
                Else
                    If y < dataMin Then dataMin = y
                    If y > dataMax Then dataMax = y
                End If
            End If
        Next i
    Next seriesIndex

    If Not found Then
        Err.Raise vbObjectError + 120, , _
            "Cannot scale Y axis: no numeric observations."
    End If

    span = dataMax - dataMin
    If span < 0.000000001 Then
        span = Abs(dataMax) * 0.1
        If span < 0.0001 Then span = 0.0001
        dataMin = dataMin - span / 2#
        dataMax = dataMax + span / 2#
    End If

    padding = (dataMax - dataMin) * 0.08
    stepSize = NiceAxisStep( _
        (dataMax - dataMin + 2# * padding) / 5#)

    axisMin = Int((dataMin - padding) / stepSize) * stepSize
    axisMax = -Int(-(dataMax + padding) / stepSize) * stepSize
    If axisMax <= axisMin Then axisMax = axisMin + stepSize

    Set ax = ch.Axes(xlValue, xlPrimary)
    With ax
        .ScaleType = xlScaleLinear
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .MajorUnitIsAuto = True

        ' Set the final bounds explicitly; Excel must not
        ' auto-anchor a yield/spread chart at zero.
        .MaximumScale = axisMax
        .MinimumScale = axisMin
        .majorUnit = stepSize
        .TickLabels.numberFormat = AxisNumberFormat(stepSize)
        .Crosses = xlMinimum
        .HasMajorGridlines = True
    End With
End Sub

Private Sub ConfigureScatterAxis(ByVal ax As Axis, _
                                 ByVal dataMin As Double, _
                                 ByVal dataMax As Double, _
                                 ByVal axisLabel As String)
    Dim span As Double, padding As Double, stepSize As Double
    Dim minScale As Double, maxScale As Double
    span = dataMax - dataMin
    If span < 0.000000001 Then
        span = Abs(dataMax) * 0.1
        If span < 0.0001 Then span = 0.0001
        dataMin = dataMin - span / 2#
        dataMax = dataMax + span / 2#
    End If
    padding = span * 0.08
    stepSize = NiceAxisStep((dataMax - dataMin + 2# * padding) / 5#)
    minScale = Int((dataMin - padding) / stepSize) * stepSize
    maxScale = -Int(-(dataMax + padding) / stepSize) * stepSize
    If maxScale <= minScale Then maxScale = minScale + stepSize
    With ax
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
        .ScaleType = xlScaleLinear
        .MinimumScale = minScale
        .MaximumScale = maxScale
        .majorUnit = stepSize
        .TickLabels.numberFormat = AxisNumberFormat(stepSize)
        .HasTitle = True
        .AxisTitle.text = axisLabel
    End With
End Sub

Private Function NiceAxisStep(ByVal rawStep As Double) As Double
    Dim magnitude As Double, fraction As Double, niceFraction As Double
    If rawStep <= 0 Then
        NiceAxisStep = 0.01
        Exit Function
    End If
    magnitude = 10# ^ Int(Log(rawStep) / Log(10#))
    fraction = rawStep / magnitude
    Select Case fraction
        Case Is <= 1#: niceFraction = 1#
        Case Is <= 2#: niceFraction = 2#
        Case Is <= 2.5: niceFraction = 2.5
        Case Is <= 5#: niceFraction = 5#
        Case Else: niceFraction = 10#
    End Select
    NiceAxisStep = niceFraction * magnitude
End Function

Private Function AxisNumberFormat(ByVal stepSize As Double) As String
    If stepSize >= 0.001 Then
        AxisNumberFormat = "0.000"
    ElseIf stepSize >= 0.0001 Then
        AxisNumberFormat = "0.0000"
    ElseIf stepSize >= 0.00001 Then
        AxisNumberFormat = "0.00000"
    Else
        AxisNumberFormat = "0.000000"
    End If
End Function

Private Function GetStartDate(ByVal latest As Date, ByVal earliest As Date, _
                              ByVal periodKey As String) As Date
    Select Case periodKey
        Case "1M": GetStartDate = DateAdd("m", -1, latest)
        Case "3M": GetStartDate = DateAdd("m", -3, latest)
        Case "6M": GetStartDate = DateAdd("m", -6, latest)
        Case "1Y": GetStartDate = DateAdd("yyyy", -1, latest)
        Case "3Y": GetStartDate = DateAdd("yyyy", -3, latest)
        Case "5Y": GetStartDate = DateAdd("yyyy", -5, latest)
        Case "ALL": GetStartDate = earliest
        Case Else
            Err.Raise vbObjectError + 107, , "Invalid period: " & periodKey
    End Select
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, _
                                  ByVal headerName As String) As Long
    Dim c As Long, lastCol As Long, found As Long, v As Variant
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 2 To lastCol
        v = ws.Cells(HEADER_ROW, c).value2
        If Not IsError(v) Then
            If StrComp(Trim$(CStr(v)), headerName, vbTextCompare) = 0 Then
                If found <> 0 Then
                    Err.Raise vbObjectError + 108, , "Duplicate header: " & headerName
                End If
                found = c
            End If
        End If
    Next c
    If found = 0 Then
        Err.Raise vbObjectError + 109, , "Header not found on " & ws.Name & ": " & headerName
    End If
    FindHeaderColumn = found
End Function

Private Function PreferredSeries(ByVal ws As Worksheet, _
                                 ByVal preferredName As String, _
                                 ByVal fallbackCol As Long) As String
    Dim c As Long, lastCol As Long, v As Variant
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    For c = 2 To lastCol
        v = ws.Cells(HEADER_ROW, c).value2
        If Not IsError(v) Then
            If StrComp(Trim$(CStr(v)), preferredName, vbTextCompare) = 0 Then
                PreferredSeries = preferredName
                Exit Function
            End If
        End If
    Next c
    If fallbackCol <= lastCol Then
        PreferredSeries = Trim$(CStr(ws.Cells(HEADER_ROW, fallbackCol).value2))
    End If
End Function

Private Function ReadColumn(ByVal ws As Worksheet, ByVal col As Long, _
                            ByVal lastRow As Long) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If lastRow = FIRST_ROW Then
        a(1, 1) = ws.Cells(FIRST_ROW, col).value2
        ReadColumn = a
    Else
        ReadColumn = ws.Range(ws.Cells(FIRST_ROW, col), _
                              ws.Cells(lastRow, col)).value2
    End If
End Function

Private Function ParseDateValue(ByVal v As Variant, ByRef result As Date) As Boolean
    Dim serial As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    On Error GoTo InvalidDate
    If IsNumeric(v) Then
        serial = CDbl(v)
        If ThisWorkbook.Date1904 Then serial = serial + 1462#
        If serial < 1000# Or serial > 2958465# Then Exit Function
        result = CDate(Fix(serial))
        ParseDateValue = True
    ElseIf IsDate(v) Then
        result = dateValue(CDate(v))
        ParseDateValue = True
    End If
    Exit Function
InvalidDate:
End Function

Private Function ParseYield(ByVal v As Variant, ByRef result As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    If IsNumeric(v) Then
        result = CDbl(v)
        ParseYield = True
    End If
End Function

Private Function ToExcelSerial(ByVal d As Date) As Double
    ToExcelSerial = CDbl(d)
    If ThisWorkbook.Date1904 Then ToExcelSerial = ToExcelSerial - 1462#
End Function

Private Function ControlName(ByVal id As String, ByVal field As String) As String
    ControlName = NAME_PREFIX & id & "_" & field
End Function

Private Sub RegisterName(ByVal id As String, ByVal field As String, _
                         ByVal Target As Range)
    UpsertWorkbookName ControlName(id, field), Target
End Sub

Private Sub UpsertWorkbookName(ByVal nameText As String, ByVal Target As Range)
    Dim nm As Name, refers As String
    refers = "='" & Replace(Target.Worksheet.Name, "'", "''") & _
             "'!" & Target.Address(True, True)
    On Error Resume Next
    Set nm = ThisWorkbook.names(nameText)
    On Error GoTo 0
    If nm Is Nothing Then
        ThisWorkbook.names.Add Name:=nameText, RefersTo:=refers
    Else
        nm.RefersTo = refers
    End If
End Sub

Private Function NamedCell(ByVal id As String, ByVal field As String) As Range
    On Error GoTo MissingName
    Set NamedCell = ThisWorkbook.names(ControlName(id, field)).RefersToRange
    If NamedCell.Cells.CountLarge <> 1 And field <> "Panel" Then
        Err.Raise vbObjectError + 110, , "Not a single parameter cell: " & ControlName(id, field)
    End If
    Exit Function
MissingName:
    Err.Raise vbObjectError + 110, , "Invalid configuration name: " & ControlName(id, field)
End Function

Private Function SelectionBelongsToChart(ByVal id As String) As Boolean
    Dim fields As Variant, f As Variant, r As Range
    If TypeName(Selection) <> "Range" Then Exit Function
    fields = Array("Panel", "Type", "Sheet1", "S1", "Sheet2", "S2", "Period", "Status")
    For Each f In fields
        Set r = Nothing
        On Error Resume Next
        Set r = ThisWorkbook.names(ControlName(id, CStr(f))).RefersToRange
        On Error GoTo 0
        If Not r Is Nothing Then
            If r.Worksheet Is ActiveSheet Then
                If Not Intersect(Selection, r) Is Nothing Then
                    SelectionBelongsToChart = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

Private Function IsChartID(ByVal id As String) As Boolean
    IsChartID = (id Like "JG####")
End Function

Private Function RegisteredChartCount() As Long
    Dim ws As Worksheet, c As Long, lastCol As Long
    Set ws = EnsureSheet(PLOT_SHEET)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        If IsChartID(CStr(ws.Cells(1, c).value2)) Then
            RegisteredChartCount = RegisteredChartCount + 1
        End If
    Next c
End Function

Private Function FindPlotBlock(ByVal ws As Worksheet, ByVal id As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        If StrComp(CStr(ws.Cells(1, c).value2), id, vbTextCompare) = 0 Then
            FindPlotBlock = c
            Exit Function
        End If
    Next c
End Function

Private Function NextChartID(ByVal ws As Worksheet) As String
    Dim i As Long, id As String
    For i = 1 To 9999
        id = "JG" & Format$(i, "0000")
        If FindPlotBlock(ws, id) = 0 Then
            NextChartID = id
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 111, , "Chart ID limit reached."
End Function

Private Function AllocatePlotBlock(ByVal ws As Worksheet, _
                                   ByVal id As String) As Long
    Dim lastCol As Long, c As Long, newCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 And Len(CStr(ws.Cells(1, 1).value2)) = 0 Then
        newCol = 1
    Else
        ' Reuse only a truly empty block, never move worksheet columns.
        For c = 1 To lastCol Step BLOCK_WIDTH
            If Application.CountA(ws.Range( _
                ws.Cells(1, c), ws.Cells( _
                Application.Max(2, ws.usedRange.Row + ws.usedRange.rows.Count), _
                c + 3))) = 0 Then
                newCol = c
                Exit For
            End If
        Next c
        If newCol = 0 Then
            newCol = ((lastCol - 1) \ BLOCK_WIDTH + 1) * BLOCK_WIDTH + 1
        End If
    End If
    If newCol + 3 > ws.Columns.Count Then
        Err.Raise vbObjectError + 112, , "PlotData column capacity exceeded."
    End If
    ws.Cells(1, newCol).value2 = id
    AllocatePlotBlock = newCol
End Function

Private Sub AddValidationList(ByVal Target As Range, ByVal items As String, _
                              ByVal allowBlank As Boolean)
    Dim sep As String
    sep = Application.International(xlListSeparator)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=Replace(items, ",", sep)
        .IgnoreBlank = allowBlank
        .InCellDropdown = True
    End With
End Sub

Private Sub AddSeriesValidationForChart(ByVal Target As Range, _
                                        ByVal src As Worksheet, _
                                        ByVal id As String, _
                                        ByVal leg As Long)
    Dim lastCol As Long, nameText As String
    lastCol = src.Cells(HEADER_ROW, src.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then
        Err.Raise vbObjectError + 130, , "No series headers on " & src.Name
    End If
    nameText = NAME_PREFIX & id & "_Headers_S" & CStr(leg)
    UpsertWorkbookName nameText, src.Range( _
        src.Cells(HEADER_ROW, 2), src.Cells(HEADER_ROW, lastCol))
    With Target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=" & nameText
        ' Allow Series 2 to be blank for a single-series line chart.
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' Put sheet names in a hidden, stable named list for the dropdowns.


'============================================================
' V8 -> V11: INSTANT DEPENDENT DROPDOWNS
' Dashboard.Worksheet_Change calls JGB_HandleDashboardSourceChange.
' This handler does NOT redraw charts or rewrite PlotData.
'============================================================

Public Sub JGB_HandleDashboardSourceChange( _
    ByVal dashboardSheet As Worksheet, _
    ByVal changedCells As Range)

    Dim plotWs As Worksheet
    Dim lastCol As Long, c As Long, chartID As String
    Dim selectedSource As Range
    Dim previousEvents As Boolean
    Dim errors As String
    Dim doLeg1 As Boolean, doLeg2 As Boolean

    If Not dashboardSheet Is ThisWorkbook.Worksheets(DASH_SHEET) Then Exit Sub
    If changedCells Is Nothing Then Exit Sub

    Set plotWs = ThisWorkbook.Worksheets(PLOT_SHEET)
    lastCol = plotWs.Cells(1, plotWs.Columns.Count).End(xlToLeft).Column

    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False

    For c = 1 To lastCol Step BLOCK_WIDTH
        chartID = Trim$(CStr(plotWs.Cells(1, c).value2))
        If IsChartID(chartID) Then
            doLeg1 = False: doLeg2 = False

            Set selectedSource = GetOptionalNamedCell(chartID, "Sheet1")
            If Not selectedSource Is Nothing Then
                If selectedSource.Worksheet Is dashboardSheet Then
                    doLeg1 = Not Intersect(changedCells, selectedSource) Is Nothing
                End If
            End If

            Set selectedSource = GetOptionalNamedCell(chartID, "Sheet2")
            If Not selectedSource Is Nothing Then
                If selectedSource.Worksheet Is dashboardSheet Then
                    doLeg2 = Not Intersect(changedCells, selectedSource) Is Nothing
                End If
            End If

            If doLeg1 Then UpdateOneDependentDropdown chartID, 1
            If doLeg2 Then UpdateOneDependentDropdown chartID, 2
        End If
    Next c

Done:
    Application.EnableEvents = previousEvents
    If Len(errors) > 0 Then
        MsgBox errors, vbExclamation, "JGB source dropdown update"
    End If
    Exit Sub

Failed:
    errors = "Could not update a dependent dropdown: " & Err.Description
    Resume Done
End Sub


Public Sub RefreshAllSourceDropdowns()
    Dim plotWs As Worksheet, c As Long, lastCol As Long
    Dim chartID As String, previousEvents As Boolean
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    RefreshSourceSheetList
    Set plotWs = ThisWorkbook.Worksheets(PLOT_SHEET)
    lastCol = plotWs.Cells(1, plotWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        chartID = Trim$(CStr(plotWs.Cells(1, c).value2))
        If IsChartID(chartID) Then
            UpdateOneDependentDropdown chartID, 1
            UpdateOneDependentDropdown chartID, 2
        End If
    Next c
CleanExit:
    Application.EnableEvents = previousEvents
    Exit Sub
Failed:
    MsgBox "RefreshAllSourceDropdowns failed: " & Err.Description, _
           vbExclamation
    Resume CleanExit
End Sub


Private Sub UpdateOneDependentDropdown( _
    ByVal chartID As String, ByVal leg As Long)

    Dim sourceCell As Range, seriesCell As Range, statusCell As Range
    Dim sourceSheet As Worksheet, headerText As String, sourceName As String
    Dim sourceField As String, seriesField As String, errorMessage As String

    If leg = 1 Then
        sourceField = "Sheet1": seriesField = "S1"
    Else
        sourceField = "Sheet2": seriesField = "S2"
    End If

    Set sourceCell = NamedCell(chartID, sourceField)
    Set seriesCell = NamedCell(chartID, seriesField)
    Set statusCell = NamedCell(chartID, "Status")
    sourceName = Trim$(CStr(sourceCell.value2))
    headerText = Trim$(CStr(seriesCell.value2))

    On Error GoTo BadSource
    Set sourceSheet = GetSourceSheet(sourceName)
    AddSeriesValidationForChart seriesCell, sourceSheet, chartID, leg

    ' Preserve a matching existing header; clear only stale choices.
    If Len(headerText) > 0 Then
        If Not SeriesHeaderExists(sourceSheet, headerText) Then
            seriesCell.ClearContents
            statusCell.value2 = "Choose Series " & CStr(leg) & _
                " column on " & sourceSheet.Name
            Exit Sub
        End If
    End If

    statusCell.value2 = "Source updated - click UpdateAll"
    Exit Sub

BadSource:
    ' Never offer a dropdown from the PREVIOUS data source.
    errorMessage = Err.Description
    On Error Resume Next
    seriesCell.Validation.Delete
    seriesCell.ClearContents
    statusCell.value2 = "ERROR: " & Left$(errorMessage, 110)
    On Error GoTo 0
End Sub


Private Function SeriesHeaderExists( _
    ByVal src As Worksheet, ByVal desiredHeader As String) As Boolean

    Dim c As Long, lastCol As Long, v As Variant, matches As Long
    lastCol = src.Cells(HEADER_ROW, src.Columns.Count).End(xlToLeft).Column
    For c = 2 To lastCol
        v = src.Cells(HEADER_ROW, c).value2
        If Not IsError(v) Then
            If StrComp(Trim$(CStr(v)), desiredHeader, vbTextCompare) = 0 Then
                matches = matches + 1
            End If
        End If
    Next c
    ' Duplicate headers are not usable: FindHeaderColumn will reject.
    SeriesHeaderExists = (matches = 1)
End Function


Private Function HasSourceLayout(ByVal ws As Worksheet) As Boolean
    Dim lastCol As Long, lastRow As Long, col As Long
    Dim v As Variant
    lastCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < FIRST_ROW Then Exit Function
    For col = 2 To lastCol
        v = ws.Cells(HEADER_ROW, col).value2
        If Not IsError(v) Then
            If Len(Trim$(CStr(v))) > 0 Then
                HasSourceLayout = True
                Exit Function
            End If
        End If
    Next col
End Function




'============================================================
' MIGRATE AN EXISTING V8 DASHBOARD IN PLACE
'
' Keeps all chart objects and their positions/sizes, the Chart IDs,
' the 8-row panels and their selected values, PlotData and exports.
' Only refreshes source/header dropdown validations and list names.
' If a V8 panel references a source outside the 3 allowed sheets,
' report it and leave its values untouched for manual resolution.
'
' This is safe to rerun.
'============================================================


'============================================================
' V6/V8 -> Unified Controller Migration
'============================================================
' Run MigrateExistingJGBCharts ONCE on a backup of a V6 or V8 book.
' Safe to rerun. Old panels/charts/PlotData are never erased or moved.
' The migration only adds missing Sheet selectors and updates validations.
Public Sub MigrateExistingJGBCharts()
    Dim dash As Worksheet, plotWs As Worksheet
    Dim lastCol As Long, c As Long, id As String
    Dim s1Cell As Range, s2Cell As Range, sh1Cell As Range, sh2Cell As Range
    Dim src1 As Worksheet, src2 As Worksheet
    Dim oldScreen As Boolean, succeeded As Long, Failed As Long
    Dim notes As String, errorForID As String
    Dim oldEnableEvents As Boolean
    On Error GoTo Fatal
    oldScreen = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set dash = ThisWorkbook.Worksheets(DASH_SHEET)
    Set plotWs = ThisWorkbook.Worksheets(PLOT_SHEET)
    RefreshSourceSheetList

    lastCol = plotWs.Cells(1, plotWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        id = Trim$(CStr(plotWs.Cells(1, c).value2))
        If IsChartID(id) Then
            errorForID = vbNullString
            If MigrateOneChartControls(id, errorForID) Then
                succeeded = succeeded + 1
                If Len(errorForID) > 0 Then notes = notes & id & ": " & errorForID & vbCrLf
            Else
                Failed = Failed + 1
                notes = notes & id & ": " & errorForID & vbCrLf
            End If
        End If
    Next c

Done:
    Application.ScreenUpdating = oldScreen
    ' Explicitly restore event handling after migration, even if a former
    ' VBA error left EnableEvents False. Required for instant dropdowns.
    Application.EnableEvents = True
    If Failed > 0 Or Len(notes) > 0 Then
        MsgBox "Migration completed: " & succeeded & " / " & _
               (succeeded + Failed) & " charts." & vbCrLf & notes, _
               vbExclamation, "JGB Chart Controller"
    Else
        MsgBox "Migration completed: " & succeeded & _
               " charts. Chart data and positions preserved.", _
               vbInformation, "JGB Chart Controller"
    End If
    Exit Sub
Fatal:
    Failed = Failed + 1
    notes = notes & "Migration stopped: " & Err.Description & vbCrLf
    Resume Done
End Sub

' Compatibility entry points for old usage or buttons.
Public Sub MigrateV6DashboardToV11()
    MigrateExistingJGBCharts
End Sub

Public Sub MigrateV8DashboardToV11()
    MigrateExistingJGBCharts
End Sub

' Named refs from very old builds sometimes pointed to PlotData arrays.
' Never reset a control that correctly points to a single Dashboard cell.
Private Function IsValidControlCell(ByVal r As Range) As Boolean
    If r Is Nothing Then Exit Function
    If r.Cells.CountLarge <> 1 Then Exit Function
    IsValidControlCell = (r.Worksheet.Name = DASH_SHEET)
End Function

Private Sub RepairLegacyControlNameIfNeeded( _
    ByVal id As String, ByVal field As String, _
    ByVal labelText As String)
    Dim r As Range, panel As Range, p As Range
    Dim rr As Long, Target As Range
    Set r = GetOptionalNamedCell(id, field)
    If IsValidControlCell(r) Then Exit Sub
    Set panel = GetOptionalNamedCell(id, "Panel")
    If panel Is Nothing Then Err.Raise vbObjectError + 191, , _
        "Missing Panel for " & id
    Set p = panel.Cells(1, 1)
    For rr = 1 To panel.rows.Count
        If StrComp(Trim$(CStr(p.Cells(rr, 1).value2)), _
                   labelText, vbTextCompare) = 0 Then
            Set Target = p.Cells(rr, 2)
            RegisterName id, field, Target
            Exit Sub
        End If
    Next rr
    Err.Raise vbObjectError + 192, , _
        "Cannot recover named control " & field & " for " & id
End Sub

Private Function MigrateOneChartControls( _
    ByVal id As String, ByRef info As String) As Boolean
    Dim src1 As Worksheet, src2 As Worksheet
    Dim s1Cell As Range, s2Cell As Range
    Dim sh1Cell As Range, sh2Cell As Range
    Dim srcName1 As String, srcName2 As String
    On Error GoTo Bad
    ' V6 names: Type, S1, S2, Period, Status (6-row Panel).
    ' V8 names: same + Sheet1/Sheet2 (8-row Panel).
    RepairLegacyControlNameIfNeeded id, "Type", "Chart Type"
    RepairLegacyControlNameIfNeeded id, "S1", "Series 1 / X"
    RepairLegacyControlNameIfNeeded id, "S2", "Series 2 / Y"
    RepairLegacyControlNameIfNeeded id, "Period", "Period"
    RepairLegacyControlNameIfNeeded id, "Status", "Status"

    EnsureSourceFields id
    Set sh1Cell = NamedCell(id, "Sheet1")
    Set sh2Cell = NamedCell(id, "Sheet2")
    Set s1Cell = NamedCell(id, "S1")
    Set s2Cell = NamedCell(id, "S2")
    AddSourceSheetValidation sh1Cell
    AddSourceSheetValidation sh2Cell
    AddValidationList NamedCell(id, "Type"), "Time Series,Scatter", False
    AddValidationList NamedCell(id, "Period"), _
                      "1M,3M,6M,1Y,3Y,5Y,ALL", False

    srcName1 = Trim$(CStr(sh1Cell.value2))
    srcName2 = Trim$(CStr(sh2Cell.value2))
    ' Invalid sources are reported; the existing selection is not overwritten.
    On Error GoTo InvalidSource
    Set src1 = GetSourceSheet(srcName1)
    Set src2 = GetSourceSheet(srcName2)
    On Error GoTo Bad
    AddSeriesValidationForChart s1Cell, src1, id, 1
    AddSeriesValidationForChart s2Cell, src2, id, 2
    If Len(Trim$(CStr(s1Cell.value2))) > 0 Then
        If Not SeriesHeaderExists(src1, CStr(s1Cell.value2)) Then
            info = "Series 1 header not found on " & src1.Name & ". "
        End If
    End If
    If Len(Trim$(CStr(s2Cell.value2))) > 0 Then
        If Not SeriesHeaderExists(src2, CStr(s2Cell.value2)) Then
            info = info & "Series 2 header not found on " & src2.Name & "."
        End If
    End If
    MigrateOneChartControls = True
    Exit Function
InvalidSource:
    info = "Invalid source: " & Err.Description
    Exit Function
Bad:
    info = Err.Description
End Function

' This is the recovery action used when source dropdowns stopped updating.
' Unlike a normal chart refresh, it also re-enables worksheet events.
Public Sub RepairJGBDropdowns()
    On Error GoTo Failed
    Application.EnableEvents = True
    MigrateExistingJGBCharts
    RefreshAllSourceDropdowns
    Application.EnableEvents = True
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "RepairJGBDropdowns failed: " & Err.Description, vbExclamation
End Sub

' Recreates parameter names only where broken; preserves separately moved cells.
Public Sub RepairJGBParameterNames()
    MigrateExistingJGBCharts
End Sub

Private Sub RefreshSourceSheetList()
    Dim wsList As Worksheet, sourceSheet As Worksheet
    Dim allowed As Variant, i As Long
    allowed = Array(ALLOWED_SOURCE_1, ALLOWED_SOURCE_2, ALLOWED_SOURCE_3)

    ' These 3 worksheets are required at initialization.
    ' Validate BEFORE writing the dependent list.
    For i = LBound(allowed) To UBound(allowed)
        Set sourceSheet = GetSourceSheet(CStr(allowed(i)))
    Next i

    Set wsList = EnsureSheet(LIST_SHEET)
    wsList.Range("A1:A4").ClearContents
    wsList.Cells(1, 1).value2 = "Available source sheets"
    For i = LBound(allowed) To UBound(allowed)
        wsList.Cells(i + 2, 1).value2 = CStr(allowed(i))
    Next i
    UpsertWorkbookName "JGBSourceSheetNames", wsList.Range("A2:A4")
    wsList.Visible = xlSheetVeryHidden
End Sub

Private Sub AddSourceSheetValidation(ByVal Target As Range)
    With Target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=JGBSourceSheetNames"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
End Sub


Private Function GetSourceSheet(ByVal sourceName As String) As Worksheet
    Select Case UCase$(Trim$(sourceName))
        Case UCase$(ALLOWED_SOURCE_1), _
             UCase$(ALLOWED_SOURCE_2), _
             UCase$(ALLOWED_SOURCE_3)
            ' Permitted data source.
        Case Else
            Err.Raise vbObjectError + 132, , _
                "Source must be JGBData, OISData or ASWData: " & sourceName
    End Select

    On Error Resume Next
    Set GetSourceSheet = ThisWorkbook.Worksheets(sourceName)
    On Error GoTo 0
    If GetSourceSheet Is Nothing Then
        Err.Raise vbObjectError + 133, , _
            "Source worksheet is missing: " & sourceName
    End If
    If Not HasSourceLayout(GetSourceSheet) Then
        Err.Raise vbObjectError + 137, , _
            "Invalid data on " & sourceName & _
            ": A2:A must contain dates; B1:... must contain headers."
    End If
End Function

' First refresh of a V6 chart upgrades its 6-row panel in place.
' Rows 7-8 are used only if all four cells are empty; no existing
' controls or worksheet contents are overwritten.
Private Sub EnsureSourceFields(ByVal id As String)
    ' V6: 6-row panel and no Sheet1/Sheet2 controls.
    ' V8: 8-row panel with independent Sheet1/Sheet2 controls.
    ' Preserve the original V6 meaning: BOTH series came from JGBData.
    ' Never overwrite occupied cells or resize the legacy Panel name.
    Dim panel As Range, p As Range, extra As Range
    Dim sheetCell1 As Range, sheetCell2 As Range
    Dim pos As Long, r As Long, c As Long
    Dim candidate As Range

    Set panel = GetOptionalNamedCell(id, "Panel")
    If panel Is Nothing Then Err.Raise vbObjectError + 133, , _
        "Missing Panel name for " & id
    If Not panel.Worksheet Is ThisWorkbook.Worksheets(DASH_SHEET) Then _
        Err.Raise vbObjectError + 133, , "Panel is not on Dashboard: " & id
    If panel.rows.Count < 6 Or panel.Columns.Count < 2 Then _
        Err.Raise vbObjectError + 133, , "Invalid Panel size: " & id

    Set p = panel.Cells(1, 1)
    Set sheetCell1 = GetOptionalNamedCell(id, "Sheet1")
    Set sheetCell2 = GetOptionalNamedCell(id, "Sheet2")

    ' Recover V8 names if the names alone are missing but labels exist.
    If panel.rows.Count >= 8 Then
        If CStr(p.Offset(2, 0).value2) = "Sheet 1 / X" And _
           CStr(p.Offset(4, 0).value2) = "Sheet 2 / Y" Then
            If Not IsValidControlCell(sheetCell1) Then
                RegisterName id, "Sheet1", p.Offset(2, 1)
            End If
            If Not IsValidControlCell(sheetCell2) Then
                RegisterName id, "Sheet2", p.Offset(4, 1)
            End If
            Exit Sub
        End If
    End If

    If IsValidControlCell(sheetCell1) And _
       IsValidControlCell(sheetCell2) Then Exit Sub

    ' Find 2x2 blank area below old V6 panel, or just to its right.
    Set extra = p.Offset(6, 0).Resize(2, 2)
    If Application.CountA(extra) <> 0 Then
        Set extra = Nothing
        For pos = 3 To 30 Step 3
            If p.Column + pos + 1 <= p.Worksheet.Columns.Count Then
                Set candidate = p.Offset(0, pos).Resize(2, 2)
                If Application.CountA(candidate) = 0 Then
                    Set extra = candidate
                    Exit For
                End If
            End If
        Next pos
    End If
    If extra Is Nothing Then Err.Raise vbObjectError + 134, , _
        "Cannot locate two empty rows/columns for Sheet controls: " & id

    extra.Cells(1, 1).value2 = "Sheet 1 / X"
    extra.Cells(2, 1).value2 = "Sheet 2 / Y"
    ' V6 originally had both legs on JGBData. Never silently switch S2 to OISData.
    extra.Cells(1, 2).value2 = ALLOWED_SOURCE_1
    extra.Cells(2, 2).value2 = ALLOWED_SOURCE_1
    RegisterName id, "Sheet1", extra.Cells(1, 2)
    RegisterName id, "Sheet2", extra.Cells(2, 2)
    With extra
        .Font.Name = "Arial"
        .Font.Size = 9
        .Interior.Color = RGB(245, 247, 250)
    End With
    extra.Columns(2).Interior.Color = vbWhite
    ' Names (not the old Panel bounding box) own the two extra controls.
End Sub
Private Function GetOptionalNamedCell(ByVal id As String, _
                                      ByVal field As String) As Range
    On Error Resume Next
    Set GetOptionalNamedCell = _
        ThisWorkbook.names(ControlName(id, field)).RefersToRange
    On Error GoTo 0
End Function

' Run this once if you want to add source selectors to every old
' chart without otherwise changing existing chart data/formatting.
Public Sub UpgradeExistingJGBCharts()
    ' Legacy entry point; retained for older macro buttons.
    MigrateExistingJGBCharts
End Sub
'============================================================
' DELETE ONE CHART (integrated; no separate delete module needed)
'============================================================
Public Sub DeleteChart()
    DeleteSelectedJGBChart
End Sub

Public Sub DeleteSelectedJGBChart()
    Dim dash As Worksheet, plotWs As Worksheet
    Dim id As String, col As Long, lastRow As Long, c As Long
    Dim co As chartObject, panel As Range, r As Range
    Dim fields As Variant, f As Variant, n As Long
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim didStart As Boolean
    On Error GoTo Failed
    Set dash = ThisWorkbook.Worksheets(DASH_SHEET)
    Set plotWs = ThisWorkbook.Worksheets(PLOT_SHEET)

    id = IdentifySelectedChartID(dash, plotWs)
    If Len(id) = 0 Then
        MsgBox "Select a chart or a cell in its configuration panel first.", _
            vbInformation, "DeleteChart"
        Exit Sub
    End If
    col = FindPlotBlock(plotWs, id)
    If col = 0 Then Err.Raise vbObjectError + 160, , _
        "PlotData ID not found: " & id
    If MsgBox("Delete chart " & id & " and its panel, PlotData block, " & _
              "and workbook names?" & vbCrLf & _
              "Formulas elsewhere referring to these names may break.", _
              vbQuestion + vbYesNo + vbDefaultButton2, "DeleteChart") <> vbYes Then Exit Sub

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    didStart = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    Set co = dash.ChartObjects(CHART_PREFIX & id)
    On Error GoTo Failed
    If Not co Is Nothing Then co.Delete

    ' Named parameters can be moved separately from the original panel.
    fields = Array("Type", "Sheet1", "S1", "Sheet2", "S2", "Period", "Status")
    For Each f In fields
        Set r = GetOptionalNamedCell(id, CStr(f))
        If Not r Is Nothing Then
            If r.Worksheet Is dash Then
                If r.Cells.CountLarge = 1 Then
                    ' V6 migration may put source selectors outside Panel.
                    If CStr(f) = "Sheet1" Or CStr(f) = "Sheet2" Then
                        If r.Column > 1 Then
                            If CStr(f) = "Sheet1" Then
                                If CStr(r.Offset(0, -1).value2) = "Sheet 1 / X" Then _
                                    r.Offset(0, -1).Clear
                            Else
                                If CStr(r.Offset(0, -1).value2) = "Sheet 2 / Y" Then _
                                    r.Offset(0, -1).Clear
                            End If
                        End If
                    End If
                    r.Clear
                End If
            End If
        End If
    Next f
    Set panel = GetOptionalNamedCell(id, "Panel")
    If Not panel Is Nothing Then
        If panel.Worksheet Is dash Then panel.Clear
    End If

    lastRow = 2
    For c = col To col + BLOCK_WIDTH - 1
        If plotWs.Cells(plotWs.rows.Count, c).End(xlUp).Row > lastRow Then
            lastRow = plotWs.Cells(plotWs.rows.Count, c).End(xlUp).Row
        End If
    Next c
    plotWs.Range(plotWs.Cells(1, col), _
                 plotWs.Cells(lastRow, col + BLOCK_WIDTH - 1)).Clear

    For n = ThisWorkbook.names.Count To 1 Step -1
        If IsNameForChart(ThisWorkbook.names(n).Name, id) Then
            ThisWorkbook.names(n).Delete
        End If
    Next n

CleanExit:
    If didStart Then
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    Exit Sub
Failed:
    MsgBox "DeleteChart failed (part of the deletion may have completed): " & _
        Err.Description, vbExclamation
    Resume CleanExit
End Sub
Private Function IsNameForChart(ByVal fullName As String, _
                                ByVal id As String) As Boolean
    Dim nm As String, bang As Long
    bang = InStrRev(fullName, "!")
    If bang > 0 Then nm = mid$(fullName, bang + 1) Else nm = fullName
    IsNameForChart = _
        (LCase$(Left$(nm, Len(NAME_PREFIX & id & "_"))) = _
         LCase$(NAME_PREFIX & id & "_"))
End Function

Private Function IdentifySelectedChartID(ByVal dash As Worksheet, _
                                         ByVal plotWs As Worksheet) As String
    Dim co As chartObject, id As String
    Dim c As Long, lastCol As Long, chosen As Range
    If Not ActiveSheet Is dash Then Exit Function

    ' Chart selection is preferred over a stale ActiveCell.
    If Not ActiveChart Is Nothing Then
        For Each co In dash.ChartObjects
            If co.Chart Is ActiveChart Then
                If Left$(co.Name, Len(CHART_PREFIX)) = CHART_PREFIX Then
                    id = mid$(co.Name, Len(CHART_PREFIX) + 1)
                    If IsChartID(id) Then
                        IdentifySelectedChartID = id
                        Exit Function
                    End If
                End If
            End If
        Next co
    End If

    If TypeName(Selection) = "Range" Then
        Set chosen = Selection
    Else
        Set chosen = ActiveCell
    End If
    If chosen Is Nothing Then Exit Function
    lastCol = plotWs.Cells(1, plotWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        id = Trim$(CStr(plotWs.Cells(1, c).value2))
        If IsChartID(id) Then
            If SelectionInChartPanel(id, chosen) Then
                IdentifySelectedChartID = id
                Exit Function
            End If
        End If
    Next c
End Function

Private Function SelectionInChartPanel(ByVal id As String, _
                                       ByVal chosen As Range) As Boolean
    Dim fields As Variant, f As Variant, r As Range
    fields = Array("Panel", "Type", "Sheet1", "S1", _
                   "Sheet2", "S2", "Period", "Status")
    For Each f In fields
        Set r = GetOptionalNamedCell(id, CStr(f))
        If Not r Is Nothing Then
            If r.Worksheet Is chosen.Worksheet Then
                If Not Intersect(r, chosen) Is Nothing Then
                    SelectionInChartPanel = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

' Repair legacy parameter names if they were overwritten by old PlotData aliases.
' Handles both the original 6-row layout and the new 8-row layout.

