Option Explicit

'============================================================
' JGB Multi Chart Controller (V6)
' Data: JGBData (A: Date, row 1: headers)
' Output: PlotData (four columns per chart)
' Dashboard: independent, movable control panels and charts
' Date axis: calendar-aligned month labels for time series
' Scatter: independent numeric XY axes
' Time-series Y-axis: data-driven min/max with 8% padding
'============================================================
Private Const SRC_SHEET As String = "JGBData"
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
    Dim ws As Worksheet
    Dim p As Range

    Set ws = EnsureSheet(DASH_SHEET)
    Call EnsureSheet(PLOT_SHEET)

    If RegisteredChartCount() > 0 Then
        RefreshAllJGBCharts
        Exit Sub
    End If

    ws.Activate
    Set p = ws.Range("A2")
    If Application.CountA(p.Resize(6, 2)) > 0 Then
        Set p = ws.Cells(ws.rows.Count, "A").End(xlUp).Offset(3, 0)
    End If
    p.Select
    AddJGBChart
End Sub

Public Sub AddJGBChart()
    Dim ws As Worksheet, plotWs As Worksheet, srcWs As Worksheet
    Dim p As Range, panel As Range, co As chartObject
    Dim id As String, s1 As String, s2 As String, errText As String
    Dim blockCol As Long
    Dim oldScreen As Boolean, oldEvents As Boolean

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo Failed

    Set ws = EnsureSheet(DASH_SHEET)
    Set plotWs = EnsureSheet(PLOT_SHEET)
    Set srcWs = ThisWorkbook.Worksheets(SRC_SHEET)

    If Not ActiveSheet Is ws Then
        MsgBox "Dashboard上で配置開始セルを選択してください。", vbExclamation
        Exit Sub
    End If

    Set p = ActiveCell
    Set panel = p.Resize(6, 2)
    If Application.CountA(panel) > 0 Then
        MsgBox "設定パネルを置く6行×2列が空白ではありません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    id = NextChartID(plotWs)
    blockCol = AllocatePlotBlock(plotWs, id)
    s1 = PreferredSeries(srcWs, "2Y JGB", 2)
    s2 = PreferredSeries(srcWs, "10Y JGB", 3)
    If StrComp(s1, s2, vbTextCompare) = 0 Then s2 = vbNullString

    p.value2 = "Chart ID"
    p.Offset(0, 1).value2 = id
    p.Offset(1, 0).value2 = "Chart Type"
    p.Offset(1, 1).value2 = "Time Series"
    p.Offset(2, 0).value2 = "Series 1 / X"
    p.Offset(2, 1).value2 = s1
    p.Offset(3, 0).value2 = "Series 2 / Y"
    p.Offset(3, 1).value2 = s2
    p.Offset(4, 0).value2 = "Period"
    p.Offset(4, 1).value2 = "3Y"
    p.Offset(5, 0).value2 = "Status"
    p.Offset(5, 1).value2 = "Not updated"

    RegisterName id, "Panel", panel
    RegisterName id, "Type", p.Offset(1, 1)
    RegisterName id, "S1", p.Offset(2, 1)
    RegisterName id, "S2", p.Offset(3, 1)
    RegisterName id, "Period", p.Offset(4, 1)
    RegisterName id, "Status", p.Offset(5, 1)

    AddValidationList p.Offset(1, 1), "Time Series,Scatter", False
    AddValidationList p.Offset(4, 1), "1M,3M,6M,1Y,3Y,5Y,ALL", False
    AddSeriesValidation p.Offset(2, 1), srcWs
    AddSeriesValidation p.Offset(3, 1), srcWs

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
    p.Offset(1, 1).Resize(4, 1).Interior.Color = vbWhite
    ' Neither column widths nor row heights are changed.

    Set co = ws.ChartObjects.Add(p.Left, p.Offset(8, 0).Top, 720, 360)
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
    Dim src As Worksheet, dst As Worksheet, dash As Worksheet
    Dim chartType As String, s1 As String, s2 As String, periodKey As String
    Dim col1 As Long, col2 As Long, lastRow As Long, n As Long
    Dim i As Long, j As Long, outCount As Long
    Dim dates As Variant, data1 As Variant, data2 As Variant
    Dim buffer() As Variant, output() As Variant
    Dim dt As Date, earliest As Date, latest As Date, startDate As Date
    Dim v1 As Double, v2 As Double
    Dim ok1 As Boolean, ok2 As Boolean, validRow As Boolean
    Dim foundDate As Boolean, hasSecond As Boolean
    Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

    Set src = ThisWorkbook.Worksheets(SRC_SHEET)
    Set dst = ThisWorkbook.Worksheets(PLOT_SHEET)
    Set dash = ThisWorkbook.Worksheets(DASH_SHEET)

    chartType = UCase$(Trim$(CStr(NamedCell(id, "Type").value2)))
    s1 = Trim$(CStr(NamedCell(id, "S1").value2))
    s2 = Trim$(CStr(NamedCell(id, "S2").value2))
    periodKey = UCase$(Trim$(CStr(NamedCell(id, "Period").value2)))

    If chartType <> "TIME SERIES" And chartType <> "SCATTER" Then
        Err.Raise vbObjectError + 101, , "Invalid chart type."
    End If
    If Len(s1) = 0 Then Err.Raise vbObjectError + 102, , "Series 1 is required."
    hasSecond = (Len(s2) > 0)
    If chartType = "SCATTER" And Not hasSecond Then
        Err.Raise vbObjectError + 103, , "Scatter requires Series 2."
    End If

    col1 = FindHeaderColumn(src, s1)
    If hasSecond Then col2 = FindHeaderColumn(src, s2)
    lastRow = src.Cells(src.rows.Count, 1).End(xlUp).Row
    If lastRow < FIRST_ROW Then
        Err.Raise vbObjectError + 104, , "JGBData contains no observations."
    End If

    n = lastRow - FIRST_ROW + 1
    dates = ReadColumn(src, 1, lastRow)
    data1 = ReadColumn(src, col1, lastRow)
    If hasSecond Then data2 = ReadColumn(src, col2, lastRow)

    foundDate = False
    For i = 1 To n
        If ParseDateValue(dates(i, 1), dt) Then
            ok1 = ParseYield(data1(i, 1), v1)
            ok2 = False
            If hasSecond Then ok2 = ParseYield(data2(i, 1), v2)
            If chartType = "SCATTER" Then
                validRow = ok1 And ok2
            Else
                validRow = ok1 Or ok2
            End If
            If validRow Then
                If Not foundDate Then
                    earliest = dt
                    latest = dt
                    foundDate = True
                Else
                    If dt < earliest Then earliest = dt
                    If dt > latest Then latest = dt
                End If
            End If
        End If
    Next i
    If Not foundDate Then Err.Raise vbObjectError + 105, , "No valid observations."
    startDate = GetStartDate(latest, earliest, periodKey)

    ReDim buffer(1 To n, 1 To 3)
    For i = 1 To n
        If ParseDateValue(dates(i, 1), dt) Then
            If dt >= startDate And dt <= latest Then
                ok1 = ParseYield(data1(i, 1), v1)
                ok2 = False
                If hasSecond Then ok2 = ParseYield(data2(i, 1), v2)
                If chartType = "SCATTER" Then
                    validRow = ok1 And ok2
                Else
                    validRow = ok1 Or ok2
                End If
                If validRow Then
                    outCount = outCount + 1
                    buffer(outCount, 1) = ToExcelSerial(dt)
                    If ok1 Then buffer(outCount, 2) = v1
                    If ok2 Then buffer(outCount, 3) = v2
                End If
            End If
        End If
    Next i
    If outCount = 0 Then
        Err.Raise vbObjectError + 106, , "No observations in selected period."
    End If

    ReDim output(1 To outCount, 1 To 3)
    For i = 1 To outCount
        For j = 1 To 3
            output(i, j) = buffer(i, j)
        Next j
    Next i

    ' Scatter X/Y axes are scaled using only valid paired observations.
    If chartType = "SCATTER" Then
        xMin = CDbl(output(1, 2))
        xMax = xMin
        yMin = CDbl(output(1, 3))
        yMax = yMin
        For i = 2 To outCount
            v1 = CDbl(output(i, 2))
            v2 = CDbl(output(i, 3))
            If v1 < xMin Then xMin = v1
            If v1 > xMax Then xMax = v1
            If v2 < yMin Then yMin = v2
            If v2 > yMax Then yMax = v2
        Next i
    End If

    WritePlotBlock dst, blockCol, output, outCount, s1, s2, startDate, latest
    UpdatePlotNames id, dst, blockCol, outCount
    UpdateChartObject dash, dst, id, blockCol, chartType, s1, s2, _
                      hasSecond, startDate, latest, outCount, _
                      xMin, xMax, yMin, yMax

    NamedCell(id, "Status").value2 = "OK | " & outCount & " rows | " & _
        Format$(startDate, "yyyy/mm/dd") & " - " & Format$(latest, "yyyy/mm/dd")
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
        Set co = dash.ChartObjects.Add(p.Left, p.Offset(8, 0).Top, 720, 360)
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
                    .Values = yRange
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
                .Values = plotWs.Range(plotWs.Cells(3, blockCol + 2), _
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
        Err.Raise vbObjectError + 109, , "JGBData header not found: " & headerName
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
                         ByVal target As Range)
    UpsertWorkbookName ControlName(id, field), target
End Sub

Private Sub UpsertWorkbookName(ByVal nameText As String, ByVal target As Range)
    Dim nm As Name, refers As String
    refers = "='" & Replace(target.Worksheet.Name, "'", "''") & _
             "'!" & target.Address(True, True)
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
    fields = Array("Panel", "Type", "S1", "S2", "Period", "Status")
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
    Dim lastCol As Long, newCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 And Len(CStr(ws.Cells(1, 1).value2)) = 0 Then
        newCol = 1
    Else
        newCol = ((lastCol - 1) \ BLOCK_WIDTH + 1) * BLOCK_WIDTH + 1
    End If
    If newCol + 3 > ws.Columns.Count Then
        Err.Raise vbObjectError + 112, , "PlotData column capacity exceeded."
    End If
    ws.Cells(1, newCol).value2 = id
    AllocatePlotBlock = newCol
End Function

Private Sub AddValidationList(ByVal target As Range, ByVal items As String, _
                              ByVal allowBlank As Boolean)
    Dim sep As String
    sep = Application.International(xlListSeparator)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=Replace(items, ",", sep)
        .IgnoreBlank = allowBlank
        .InCellDropdown = True
    End With
End Sub

Private Sub AddSeriesValidation(ByVal target As Range, ByVal src As Worksheet)
    Dim lastCol As Long, sourceAddress As String
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Sub
    sourceAddress = "='" & Replace(src.Name, "'", "''") & "'!" & _
                    src.Range(src.Cells(1, 2), src.Cells(1, lastCol)).Address
    On Error Resume Next
    ThisWorkbook.names("JGBSeriesHeaders").Delete
    On Error GoTo 0
    ThisWorkbook.names.Add Name:="JGBSeriesHeaders", RefersTo:=sourceAddress
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="=JGBSeriesHeaders"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

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

' Repair utility retained for legacy workbooks with overwritten S1/S2 names.
' Assumes the six-row configuration panel has not been split into separate locations.
Public Sub RepairJGBParameterNames()
    Dim plotWs As Worksheet, panel As Range
    Dim id As String, c As Long, lastCol As Long
    Dim repaired As Long, skipped As Long, details As String
    On Error GoTo Failed
    Set plotWs = ThisWorkbook.Worksheets(PLOT_SHEET)
    lastCol = plotWs.Cells(1, plotWs.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol Step BLOCK_WIDTH
        id = Trim$(CStr(plotWs.Cells(1, c).value2))
        If IsChartID(id) Then
            Set panel = Nothing
            On Error Resume Next
            Set panel = ThisWorkbook.names(ControlName(id, "Panel")).RefersToRange
            On Error GoTo Failed
            If panel Is Nothing Then
                skipped = skipped + 1
                details = details & id & ": Panel name missing." & vbCrLf
            ElseIf panel.rows.Count < 4 Or panel.Columns.Count < 2 Then
                skipped = skipped + 1
                details = details & id & ": Invalid panel." & vbCrLf
            ElseIf StrComp(CStr(panel.Cells(3, 1).value2), "Series 1 / X", vbTextCompare) <> 0 Or _
                   StrComp(CStr(panel.Cells(4, 1).value2), "Series 2 / Y", vbTextCompare) <> 0 Then
                skipped = skipped + 1
                details = details & id & ": Labels moved." & vbCrLf
            Else
                RegisterName id, "S1", panel.Cells(3, 2)
                RegisterName id, "S2", panel.Cells(4, 2)
                repaired = repaired + 1
            End If
        End If
    Next c
    MsgBox "Repaired: " & repaired & vbCrLf & "Skipped: " & skipped & vbCrLf & details, vbInformation
    Exit Sub
Failed:
    MsgBox "Repair failed: " & Err.Description, vbExclamation
End Sub

