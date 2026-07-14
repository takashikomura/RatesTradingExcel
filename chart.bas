Option Explicit

'===============================================================================
' TIME SERIES DASHBOARD
'
' Data sheet:
'   Column A    : Dates
'   Row 1       : Series names
'   Column B... : Time-series values
'
' Dashboard:
'   Selectors       : A2:B7
'   Controls        : A9:B12
'   Chart           : C2:R19
'   X Range         : C20:R20
'   Period buttons  : C21:R21
'
' Place this code in a standard module.
' Do not name the module "Chart".
'===============================================================================

'-------------------------------------------------------------------------------
' Main configuration
'-------------------------------------------------------------------------------
Private Const DASHBOARD_SHEET As String = "Chart"
Private Const CHART_OUTPUT_RANGE As String = "C2:R19"
Private Const CHART_OBJECT_NAME As String = "ts_MainChart"

Private Const MAX_SERIES_COUNT As Long = 3
Private Const NONE_SERIES_TEXT As String = "(None)"

'-------------------------------------------------------------------------------
' Visible input cells
'-------------------------------------------------------------------------------
Private Const CELL_DATA_SHEET_1 As String = "B2"
Private Const CELL_SERIES_1 As String = "B3"

Private Const CELL_DATA_SHEET_2 As String = "B4"
Private Const CELL_SERIES_2 As String = "B5"

Private Const CELL_DATA_SHEET_3 As String = "B6"
Private Const CELL_SERIES_3 As String = "B7"

'-------------------------------------------------------------------------------
' Hidden state cells
'-------------------------------------------------------------------------------
Private Const CELL_PREVIOUS_SHEET_1 As String = "BJ1"
Private Const CELL_PREVIOUS_SERIES_1 As String = "BJ2"

Private Const CELL_PREVIOUS_SHEET_2 As String = "BJ3"
Private Const CELL_PREVIOUS_SERIES_2 As String = "BJ4"

Private Const CELL_PREVIOUS_SHEET_3 As String = "BJ5"
Private Const CELL_PREVIOUS_SERIES_3 As String = "BJ6"

Private Const CELL_START_INDEX As String = "BJ7"
Private Const CELL_WINDOW_POINTS As String = "BJ8"
Private Const CELL_TOTAL_POINTS As String = "BJ9"

Private Const CELL_Y_ZOOM As String = "BJ10"
Private Const CELL_MARKER_SIZE As String = "BJ11"

Private Const CELL_AXIS_MIN As String = "BJ12"
Private Const CELL_AXIS_MAX As String = "BJ13"
Private Const CELL_PREVIOUS_WINDOW As String = "BJ14"

'-------------------------------------------------------------------------------
' Hidden helper columns
'-------------------------------------------------------------------------------
Private Const HELPER_DATE_COL As String = "AZ"
Private Const HELPER_SERIES_1_COL As String = "BA"
Private Const HELPER_SERIES_2_COL As String = "BB"
Private Const HELPER_SERIES_3_COL As String = "BC"

Private Const SHEET_LIST_COL As String = "BE"
Private Const SERIES_LIST_1_COL As String = "BF"
Private Const SERIES_LIST_2_COL As String = "BG"
Private Const SERIES_LIST_3_COL As String = "BH"

Private Const NAME_SHEET_LIST As String = "_TS_SheetList"
Private Const NAME_SERIES_LIST_1 As String = "_TS_SeriesList1"
Private Const NAME_SERIES_LIST_2 As String = "_TS_SeriesList2"
Private Const NAME_SERIES_LIST_3 As String = "_TS_SeriesList3"

'-------------------------------------------------------------------------------
' Shape names
'-------------------------------------------------------------------------------
Private Const CONTROL_X_RANGE As String = "ts_ctlXRange"

Private Const LABEL_X_ZOOM As String = "ts_lblXZoom"
Private Const LABEL_Y_ZOOM As String = "ts_lblYZoom"
Private Const LABEL_POINT_SIZE As String = "ts_lblPointSize"

Private Const READOUT_X_ZOOM As String = "ts_readoutXZoom"
Private Const READOUT_Y_ZOOM As String = "ts_readoutYZoom"
Private Const READOUT_POINT_SIZE As String = "ts_readoutPointSize"

Private Const BUTTON_UPDATE As String = "ts_btnUpdate"

Private Const BUTTON_X_MINUS As String = "ts_btnXMinus"
Private Const BUTTON_X_PLUS As String = "ts_btnXPlus"

Private Const BUTTON_Y_MINUS As String = "ts_btnYMinus"
Private Const BUTTON_Y_PLUS As String = "ts_btnYPlus"

Private Const BUTTON_POINT_MINUS As String = "ts_btnPointMinus"
Private Const BUTTON_POINT_PLUS As String = "ts_btnPointPlus"

Private Const BUTTON_1W As String = "ts_btn1W"
Private Const BUTTON_1M As String = "ts_btn1M"
Private Const BUTTON_3M As String = "ts_btn3M"
Private Const BUTTON_6M As String = "ts_btn6M"
Private Const BUTTON_1Y As String = "ts_btn1Y"
Private Const BUTTON_3Y As String = "ts_btn3Y"
Private Const BUTTON_5Y As String = "ts_btn5Y"
Private Const BUTTON_MAX As String = "ts_btnMax"
Private Const BUTTON_AUTO_Y As String = "ts_btnAutoY"

'-------------------------------------------------------------------------------
' Display-window presets
'-------------------------------------------------------------------------------
Private Const MIN_WINDOW_POINTS As Long = 5
Private Const DEFAULT_WINDOW_POINTS As Long = 66

Private Const POINTS_1W As Long = 5
Private Const POINTS_1M As Long = 22
Private Const POINTS_3M As Long = 66
Private Const POINTS_6M As Long = 132
Private Const POINTS_1Y As Long = 264
Private Const POINTS_3Y As Long = 792
Private Const POINTS_5Y As Long = 1320

'-------------------------------------------------------------------------------
' Y-axis and point settings
'-------------------------------------------------------------------------------
Private Const DEFAULT_Y_ZOOM As Long = 100
Private Const Y_ZOOM_MIN As Long = 25
Private Const Y_ZOOM_MAX As Long = 400

Private Const DEFAULT_MARKER_SIZE As Long = 3
Private Const MARKER_SIZE_MIN As Long = 0
Private Const MARKER_SIZE_MAX As Long = 12

'===============================================================================
' SETUP
'===============================================================================
Public Sub SetupTimeSeriesDashboard()

    Dim ws As Excel.Worksheet

    Dim oldSheet1 As String
    Dim oldSeries1 As String
    Dim oldSheet2 As String
    Dim oldSeries2 As String
    Dim oldSheet3 As String
    Dim oldSeries3 As String

    Dim oldWindow As Long
    Dim oldYZoom As Long
    Dim oldMarkerSize As Long

    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    On Error GoTo ErrHandler

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    oldSheet1 = Trim$(CStr(ws.Range(CELL_DATA_SHEET_1).Value))
    oldSeries1 = Trim$(CStr(ws.Range(CELL_SERIES_1).Value))

    oldSheet2 = Trim$(CStr(ws.Range(CELL_DATA_SHEET_2).Value))
    oldSeries2 = Trim$(CStr(ws.Range(CELL_SERIES_2).Value))

    oldSheet3 = Trim$(CStr(ws.Range(CELL_DATA_SHEET_3).Value))
    oldSeries3 = Trim$(CStr(ws.Range(CELL_SERIES_3).Value))

    oldWindow = GetLongOrDefault( _
        ws.Range(CELL_WINDOW_POINTS).Value, _
        DEFAULT_WINDOW_POINTS)

    oldYZoom = GetLongOrDefault( _
        ws.Range(CELL_Y_ZOOM).Value, _
        DEFAULT_Y_ZOOM)

    oldMarkerSize = GetLongOrDefault( _
        ws.Range(CELL_MARKER_SIZE).Value, _
        DEFAULT_MARKER_SIZE)

    DeleteDashboardObjects ws
    BuildDashboardLayout ws

    ws.Range(CELL_DATA_SHEET_1).Value = oldSheet1
    ws.Range(CELL_SERIES_1).Value = oldSeries1

    ws.Range(CELL_DATA_SHEET_2).Value = oldSheet2
    ws.Range(CELL_SERIES_2).Value = oldSeries2

    ws.Range(CELL_DATA_SHEET_3).Value = oldSheet3
    ws.Range(CELL_SERIES_3).Value = oldSeries3

    ws.Range(CELL_WINDOW_POINTS).Value = oldWindow

    ws.Range(CELL_Y_ZOOM).Value = ClampLong( _
        oldYZoom, _
        Y_ZOOM_MIN, _
        Y_ZOOM_MAX)

    ws.Range(CELL_MARKER_SIZE).Value = ClampLong( _
        oldMarkerSize, _
        MARKER_SIZE_MIN, _
        MARKER_SIZE_MAX)

    BuildSheetDropdown ws
    EnsureValidDataSheetSelections ws

    BuildAllSeriesDropdowns ws
    EnsureValidSeriesSelections ws

    CreateXRangeScrollBar ws
    CreateCompactControls ws
    CreatePeriodButtons ws

    RefreshTimeSeriesDashboard

    ws.Activate

    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
    On Error GoTo ErrHandler

CleanExit:

    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents

    Exit Sub

ErrHandler:

    MsgBox _
        "Dashboard setup failed." & vbCrLf & _
        Err.Description, _
        vbExclamation

    Resume CleanExit

End Sub

'===============================================================================
' UPDATE
'===============================================================================
Public Sub ReloadTimeSeriesDashboard()

    On Error GoTo ErrHandler

    Application.CalculateFull
    DoEvents

    RefreshTimeSeriesDashboard

    Exit Sub

ErrHandler:

    MsgBox _
        "Data update failed." & vbCrLf & _
        Err.Description, _
        vbExclamation

End Sub

'===============================================================================
' MAIN REFRESH
'===============================================================================
Public Sub RefreshTimeSeriesDashboard()

    Dim ws As Excel.Worksheet

    Dim seriesSheets(1 To MAX_SERIES_COUNT) As String
    Dim seriesNames(1 To MAX_SERIES_COUNT) As String
    Dim displayNames(1 To MAX_SERIES_COUNT) As String
    Dim seriesCount As Long

    Dim xAll() As Variant
    Dim yAll() As Variant

    Dim totalPoints As Long
    Dim startIndex As Long
    Dim windowPoints As Long
    Dim endIndex As Long
    Dim maximumStartIndex As Long

    Dim yZoom As Long
    Dim markerSize As Long

    Dim visibleMin As Double
    Dim visibleMax As Double

    Dim rawAxisMin As Double
    Dim rawAxisMax As Double

    Dim niceAxisMin As Double
    Dim niceAxisMax As Double
    Dim yMajorUnit As Double
    Dim yMinorUnit As Double

    Dim selectionChanged As Boolean
    Dim errorMessage As String

    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean

    On Error GoTo ErrHandler

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    BuildSheetDropdown ws
    EnsureValidDataSheetSelections ws

    BuildAllSeriesDropdowns ws
    EnsureValidSeriesSelections ws

    seriesCount = GetSelectedSeriesPairs( _
        ws:=ws, _
        seriesSheets:=seriesSheets, _
        seriesNames:=seriesNames)

    If seriesCount = 0 Then

        MsgBox "Select at least one series.", vbExclamation
        GoTo CleanExit

    End If

    BuildDisplayNames _
        seriesSheets:=seriesSheets, _
        seriesNames:=seriesNames, _
        seriesCount:=seriesCount, _
        displayNames:=displayNames

    If Not LoadAlignedDashboardData( _
        seriesSheets:=seriesSheets, _
        seriesNames:=seriesNames, _
        seriesCount:=seriesCount, _
        xAll:=xAll, _
        yAll:=yAll, _
        dataCount:=totalPoints, _
        errorMessage:=errorMessage) Then

        MsgBox errorMessage, vbExclamation
        GoTo CleanExit

    End If

    selectionChanged = HasSeriesSelectionChanged(ws)

    windowPoints = GetLongOrDefault( _
        ws.Range(CELL_WINDOW_POINTS).Value, _
        DEFAULT_WINDOW_POINTS)

    windowPoints = ClampLong( _
        windowPoints, _
        Application.WorksheetFunction.Min( _
            MIN_WINDOW_POINTS, totalPoints), _
        totalPoints)

    maximumStartIndex = Application.WorksheetFunction.Max( _
        1, _
        totalPoints - windowPoints + 1)

    startIndex = GetLongOrDefault( _
        ws.Range(CELL_START_INDEX).Value, _
        maximumStartIndex)

    If selectionChanged Then

        startIndex = maximumStartIndex
        ws.Range(CELL_Y_ZOOM).Value = DEFAULT_Y_ZOOM

    End If

    startIndex = ClampLong( _
        startIndex, _
        1, _
        maximumStartIndex)

    endIndex = startIndex + windowPoints - 1

    If endIndex > totalPoints Then
        endIndex = totalPoints
    End If

    yZoom = ClampLong( _
        GetLongOrDefault( _
            ws.Range(CELL_Y_ZOOM).Value, _
            DEFAULT_Y_ZOOM), _
        Y_ZOOM_MIN, _
        Y_ZOOM_MAX)

    markerSize = ClampLong( _
        GetLongOrDefault( _
            ws.Range(CELL_MARKER_SIZE).Value, _
            DEFAULT_MARKER_SIZE), _
        MARKER_SIZE_MIN, _
        MARKER_SIZE_MAX)

    GetVisibleGlobalMinMax _
        yAll:=yAll, _
        seriesCount:=seriesCount, _
        startIndex:=startIndex, _
        endIndex:=endIndex, _
        minimumValue:=visibleMin, _
        maximumValue:=visibleMax

    CalculateYAxisRange _
        dataMin:=visibleMin, _
        dataMax:=visibleMax, _
        yZoomPercent:=yZoom, _
        axisMin:=rawAxisMin, _
        axisMax:=rawAxisMax

    GetNiceYAxisScale _
        rawMinimum:=rawAxisMin, _
        rawMaximum:=rawAxisMax, _
        niceMinimum:=niceAxisMin, _
        niceMaximum:=niceAxisMax, _
        majorUnit:=yMajorUnit, _
        minorUnit:=yMinorUnit

    ws.Range(CELL_START_INDEX).Value = startIndex
    ws.Range(CELL_WINDOW_POINTS).Value = windowPoints
    ws.Range(CELL_TOTAL_POINTS).Value = totalPoints

    ws.Range(CELL_Y_ZOOM).Value = yZoom
    ws.Range(CELL_MARKER_SIZE).Value = markerSize

    ws.Range(CELL_AXIS_MIN).Value = niceAxisMin
    ws.Range(CELL_AXIS_MAX).Value = niceAxisMax
    ws.Range(CELL_PREVIOUS_WINDOW).Value = windowPoints

    SaveCurrentSelectionState ws

    UpdateXRangeStatus _
        ws:=ws, _
        startIndex:=startIndex, _
        endIndex:=endIndex, _
        totalPoints:=totalPoints

    DrawTimeSeriesChart _
        ws:=ws, _
        displayNames:=displayNames, _
        seriesCount:=seriesCount, _
        xAll:=xAll, _
        yAll:=yAll, _
        startIndex:=startIndex, _
        endIndex:=endIndex, _
        axisMin:=niceAxisMin, _
        axisMax:=niceAxisMax, _
        yMajorUnit:=yMajorUnit, _
        yMinorUnit:=yMinorUnit, _
        markerSize:=markerSize

    UpdateControlLimits _
        ws:=ws, _
        totalPoints:=totalPoints, _
        windowPoints:=windowPoints

    UpdateControlReadouts _
        ws:=ws, _
        windowPoints:=windowPoints, _
        yZoom:=yZoom, _
        markerSize:=markerSize

    UpdatePeriodButtonStyles _
        ws:=ws, _
        totalPoints:=totalPoints, _
        windowPoints:=windowPoints

CleanExit:

    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents

    Exit Sub

ErrHandler:

    MsgBox _
        "Dashboard refresh failed." & vbCrLf & _
        Err.Description, _
        vbExclamation

    Resume CleanExit

End Sub

'===============================================================================
' SELECTION STATE
'===============================================================================
Private Function HasSeriesSelectionChanged( _
    ByVal ws As Excel.Worksheet) As Boolean

    HasSeriesSelectionChanged = _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_DATA_SHEET_1).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SHEET_1).Value)), _
            vbTextCompare) <> 0 Or _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_SERIES_1).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SERIES_1).Value)), _
            vbTextCompare) <> 0 Or _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_DATA_SHEET_2).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SHEET_2).Value)), _
            vbTextCompare) <> 0 Or _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_SERIES_2).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SERIES_2).Value)), _
            vbTextCompare) <> 0 Or _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_DATA_SHEET_3).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SHEET_3).Value)), _
            vbTextCompare) <> 0 Or _
        StrComp( _
            Trim$(CStr(ws.Range(CELL_SERIES_3).Value)), _
            Trim$(CStr(ws.Range(CELL_PREVIOUS_SERIES_3).Value)), _
            vbTextCompare) <> 0

End Function

Private Sub SaveCurrentSelectionState(ByVal ws As Excel.Worksheet)

    ws.Range(CELL_PREVIOUS_SHEET_1).Value = _
        ws.Range(CELL_DATA_SHEET_1).Value

    ws.Range(CELL_PREVIOUS_SERIES_1).Value = _
        ws.Range(CELL_SERIES_1).Value

    ws.Range(CELL_PREVIOUS_SHEET_2).Value = _
        ws.Range(CELL_DATA_SHEET_2).Value

    ws.Range(CELL_PREVIOUS_SERIES_2).Value = _
        ws.Range(CELL_SERIES_2).Value

    ws.Range(CELL_PREVIOUS_SHEET_3).Value = _
        ws.Range(CELL_DATA_SHEET_3).Value

    ws.Range(CELL_PREVIOUS_SERIES_3).Value = _
        ws.Range(CELL_SERIES_3).Value

End Sub

'===============================================================================
' X RANGE
'===============================================================================
Public Sub TimeSeriesXRangeChanged()

    RefreshTimeSeriesDashboard

End Sub

'===============================================================================
' X ZOOM
'
' The newest currently displayed date remains fixed.
' Only the older-date side changes.
'===============================================================================
Public Sub XZoomMinus()

    Dim ws As Excel.Worksheet
    Dim currentWindow As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    currentWindow = GetLongOrDefault( _
        ws.Range(CELL_WINDOW_POINTS).Value, _
        DEFAULT_WINDOW_POINTS)

    ApplyWindowAnchoredToLatest _
        currentWindow - GetXZoomStep(currentWindow)

End Sub

Public Sub XZoomPlus()

    Dim ws As Excel.Worksheet
    Dim currentWindow As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    currentWindow = GetLongOrDefault( _
        ws.Range(CELL_WINDOW_POINTS).Value, _
        DEFAULT_WINDOW_POINTS)

    ApplyWindowAnchoredToLatest _
        currentWindow + GetXZoomStep(currentWindow)

End Sub

Private Function GetXZoomStep( _
    ByVal currentWindow As Long) As Long

    Select Case currentWindow

        Case Is <= 20
            GetXZoomStep = 1

        Case Is <= 100
            GetXZoomStep = 5

        Case Is <= 500
            GetXZoomStep = 20

        Case Is <= 1500
            GetXZoomStep = 50

        Case Else
            GetXZoomStep = 100

    End Select

End Function

Private Sub ApplyWindowAnchoredToLatest( _
    ByVal requestedWindow As Long)

    Dim ws As Excel.Worksheet

    Dim totalPoints As Long

    Dim oldStartIndex As Long
    Dim oldWindowPoints As Long
    Dim fixedEndIndex As Long

    Dim newWindowPoints As Long
    Dim newStartIndex As Long

    Dim minimumWindow As Long
    Dim maximumAnchoredWindow As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    totalPoints = GetLongOrDefault( _
        ws.Range(CELL_TOTAL_POINTS).Value, _
        0)

    If totalPoints <= 0 Then

        RefreshTimeSeriesDashboard
        Exit Sub

    End If

    oldWindowPoints = GetLongOrDefault( _
        ws.Range(CELL_WINDOW_POINTS).Value, _
        DEFAULT_WINDOW_POINTS)

    oldWindowPoints = ClampLong( _
        oldWindowPoints, _
        1, _
        totalPoints)

    oldStartIndex = GetLongOrDefault( _
        ws.Range(CELL_START_INDEX).Value, _
        Application.WorksheetFunction.Max( _
            1, _
            totalPoints - oldWindowPoints + 1))

    oldStartIndex = ClampLong( _
        oldStartIndex, _
        1, _
        totalPoints)

    fixedEndIndex = oldStartIndex + oldWindowPoints - 1

    fixedEndIndex = ClampLong( _
        fixedEndIndex, _
        1, _
        totalPoints)

    maximumAnchoredWindow = fixedEndIndex

    minimumWindow = Application.WorksheetFunction.Min( _
        MIN_WINDOW_POINTS, _
        maximumAnchoredWindow)

    newWindowPoints = ClampLong( _
        requestedWindow, _
        minimumWindow, _
        maximumAnchoredWindow)

    newStartIndex = fixedEndIndex - newWindowPoints + 1

    If newStartIndex < 1 Then
        newStartIndex = 1
    End If

    ws.Range(CELL_WINDOW_POINTS).Value = newWindowPoints
    ws.Range(CELL_START_INDEX).Value = newStartIndex
    ws.Range(CELL_PREVIOUS_WINDOW).Value = newWindowPoints

    RefreshTimeSeriesDashboard

End Sub

'===============================================================================
' Y ZOOM
'===============================================================================
Public Sub YZoomMinus()

    AdjustYZoom -10

End Sub

Public Sub YZoomPlus()

    AdjustYZoom 10

End Sub

Public Sub ResetYAxisDashboard()

    Dim ws As Excel.Worksheet

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    ws.Range(CELL_Y_ZOOM).Value = DEFAULT_Y_ZOOM

    RefreshTimeSeriesDashboard

End Sub

Private Sub AdjustYZoom(ByVal changeValue As Long)

    Dim ws As Excel.Worksheet
    Dim currentValue As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    currentValue = GetLongOrDefault( _
        ws.Range(CELL_Y_ZOOM).Value, _
        DEFAULT_Y_ZOOM)

    ws.Range(CELL_Y_ZOOM).Value = ClampLong( _
        currentValue + changeValue, _
        Y_ZOOM_MIN, _
        Y_ZOOM_MAX)

    RefreshTimeSeriesDashboard

End Sub

'===============================================================================
' POINT SIZE
'===============================================================================
Public Sub PointSizeMinus()

    AdjustPointSize -1

End Sub

Public Sub PointSizePlus()

    AdjustPointSize 1

End Sub

Private Sub AdjustPointSize(ByVal changeValue As Long)

    Dim ws As Excel.Worksheet
    Dim currentSize As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    currentSize = GetLongOrDefault( _
        ws.Range(CELL_MARKER_SIZE).Value, _
        DEFAULT_MARKER_SIZE)

    ws.Range(CELL_MARKER_SIZE).Value = ClampLong( _
        currentSize + changeValue, _
        MARKER_SIZE_MIN, _
        MARKER_SIZE_MAX)

    RefreshTimeSeriesDashboard

End Sub

'===============================================================================
' PERIOD BUTTONS
'===============================================================================
Public Sub SetWindow1W()
    ApplyLatestWindow POINTS_1W
End Sub

Public Sub SetWindow1M()
    ApplyLatestWindow POINTS_1M
End Sub

Public Sub SetWindow3M()
    ApplyLatestWindow POINTS_3M
End Sub

Public Sub SetWindow6M()
    ApplyLatestWindow POINTS_6M
End Sub

Public Sub SetWindow1Y()
    ApplyLatestWindow POINTS_1Y
End Sub

Public Sub SetWindow3Y()
    ApplyLatestWindow POINTS_3Y
End Sub

Public Sub SetWindow5Y()
    ApplyLatestWindow POINTS_5Y
End Sub

Public Sub SetWindowMaximum()
    ApplyLatestWindow 0
End Sub

Private Sub ApplyLatestWindow(ByVal requestedPoints As Long)

    Dim ws As Excel.Worksheet

    Dim totalPoints As Long
    Dim windowPoints As Long
    Dim startIndex As Long

    Set ws = GetOrCreateWorksheet(DASHBOARD_SHEET)

    totalPoints = GetLongOrDefault( _
        ws.Range(CELL_TOTAL_POINTS).Value, _
        0)

    If totalPoints <= 0 Then

        RefreshTimeSeriesDashboard

        totalPoints = GetLongOrDefault( _
            ws.Range(CELL_TOTAL_POINTS).Value, _
            0)

    End If

    If totalPoints <= 0 Then Exit Sub

    If requestedPoints <= 0 Then

        windowPoints = totalPoints

    Else

        windowPoints = Application.WorksheetFunction.Min( _
            requestedPoints, _
            totalPoints)

    End If

    startIndex = Application.WorksheetFunction.Max( _
        1, _
        totalPoints - windowPoints + 1)

    ws.Range(CELL_WINDOW_POINTS).Value = windowPoints
    ws.Range(CELL_START_INDEX).Value = startIndex
    ws.Range(CELL_PREVIOUS_WINDOW).Value = windowPoints

    RefreshTimeSeriesDashboard

End Sub

'===============================================================================
' LAYOUT
'===============================================================================
Private Sub BuildDashboardLayout(ByVal ws As Excel.Worksheet)

    ws.Range("A1:R21").UnMerge
    ws.Range("A1:R21").ClearContents

    ' Remove data left by older dashboard versions.
    ws.Range("Z1:AS5000").ClearContents
    ws.Range("AZ1:BJ5000").ClearContents

    ' Restore columns hidden by older versions.
    ws.Columns("Z:AY").Hidden = False

    ' Hide only the current helper area.
    ws.Columns("AZ:BJ").Hidden = True

    ' Row 1 remains blank.

    ws.Range("A2").Value = "Data Sheet 1"
    ws.Range("A3").Value = "Series 1"

    ws.Range("A4").Value = "Data Sheet 2"
    ws.Range("A5").Value = "Series 2"

    ws.Range("A6").Value = "Data Sheet 3"
    ws.Range("A7").Value = "Series 3"

    With ws.Range("A2:A7")

        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("B2:B7")

        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = False

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    ApplyThinBorders _
        target:=ws.Range("A2:B7"), _
        borderColor:=RGB(190, 190, 190)

    ws.Range("C20:E20").Merge

    With ws.Range("C20")

        .Value = "X RANGE"

        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    ws.Range("P20:R20").Merge

    With ws.Range("P20")

        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = False

        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub ApplyThinBorders( _
    ByVal target As Excel.Range, _
    ByVal borderColor As Long)

    With target.Borders

        .LineStyle = xlContinuous
        .Color = borderColor
        .Weight = xlThin

    End With

End Sub

'===============================================================================
' X RANGE SCROLLBAR
'===============================================================================
Private Sub CreateXRangeScrollBar(ByVal ws As Excel.Worksheet)

    Dim target As Excel.Range
    Dim shp As Shape

    Set target = ws.Range("F20:O20")

    Set shp = ws.Shapes.AddFormControl( _
        Type:=xlScrollBar, _
        Left:=target.Left, _
        Top:=target.Top + 1, _
        Width:=target.Width, _
        Height:=Application.WorksheetFunction.Max( _
            10, _
            target.Height - 2))

    shp.Name = CONTROL_X_RANGE
    shp.Placement = xlMoveAndSize
    shp.OnAction = MacroReference("TimeSeriesXRangeChanged")

    With shp.ControlFormat

        .LinkedCell = SheetCellReference( _
            ws, _
            CELL_START_INDEX)

        .Min = 1
        .Max = 1
        .SmallChange = 1
        .LargeChange = 20

    End With

End Sub

'===============================================================================
' CONTROLS IN A:B
'===============================================================================
Private Sub CreateCompactControls(ByVal ws As Excel.Worksheet)

    CreateLabeledThreePartControl _
        ws:=ws, _
        target:=ws.Range("A9:B9"), _
        labelShapeName:=LABEL_X_ZOOM, _
        labelCaption:="X ZOOM", _
        leftButtonName:=BUTTON_X_MINUS, _
        leftCaption:="-", _
        leftMacro:="XZoomMinus", _
        readoutName:=READOUT_X_ZOOM, _
        rightButtonName:=BUTTON_X_PLUS, _
        rightCaption:="+", _
        rightMacro:="XZoomPlus"

    CreateLabeledThreePartControl _
        ws:=ws, _
        target:=ws.Range("A10:B10"), _
        labelShapeName:=LABEL_Y_ZOOM, _
        labelCaption:="Y ZOOM", _
        leftButtonName:=BUTTON_Y_MINUS, _
        leftCaption:="-", _
        leftMacro:="YZoomMinus", _
        readoutName:=READOUT_Y_ZOOM, _
        rightButtonName:=BUTTON_Y_PLUS, _
        rightCaption:="+", _
        rightMacro:="YZoomPlus"

    CreateLabeledThreePartControl _
        ws:=ws, _
        target:=ws.Range("A11:B11"), _
        labelShapeName:=LABEL_POINT_SIZE, _
        labelCaption:="POINT", _
        leftButtonName:=BUTTON_POINT_MINUS, _
        leftCaption:="-", _
        leftMacro:="PointSizeMinus", _
        readoutName:=READOUT_POINT_SIZE, _
        rightButtonName:=BUTTON_POINT_PLUS, _
        rightCaption:="+", _
        rightMacro:="PointSizePlus"

    AddActionButton _
        ws:=ws, _
        shapeName:=BUTTON_UPDATE, _
        caption:="UPDATE", _
        macroName:="ReloadTimeSeriesDashboard", _
        leftPos:=ws.Range("A12:B12").Left, _
        topPos:=ws.Range("A12:B12").Top, _
        buttonWidth:=ws.Range("A12:B12").Width, _
        buttonHeight:=ws.Range("A12:B12").Height, _
        isBold:=True

End Sub

Private Sub CreateLabeledThreePartControl( _
    ByVal ws As Excel.Worksheet, _
    ByVal target As Excel.Range, _
    ByVal labelShapeName As String, _
    ByVal labelCaption As String, _
    ByVal leftButtonName As String, _
    ByVal leftCaption As String, _
    ByVal leftMacro As String, _
    ByVal readoutName As String, _
    ByVal rightButtonName As String, _
    ByVal rightCaption As String, _
    ByVal rightMacro As String)

    Dim gap As Double

    Dim labelWidth As Double
    Dim buttonWidth As Double
    Dim readoutWidth As Double

    Dim currentLeft As Double

    gap = 1

    labelWidth = target.Width * 0.34
    buttonWidth = target.Width * 0.14

    readoutWidth = _
        target.Width - _
        labelWidth - _
        buttonWidth * 2 - _
        gap * 3

    currentLeft = target.Left

    AddControlLabel _
        ws:=ws, _
        shapeName:=labelShapeName, _
        caption:=labelCaption, _
        leftPos:=currentLeft, _
        topPos:=target.Top, _
        labelWidth:=labelWidth, _
        labelHeight:=target.Height

    currentLeft = currentLeft + labelWidth + gap

    AddActionButton _
        ws:=ws, _
        shapeName:=leftButtonName, _
        caption:=leftCaption, _
        macroName:=leftMacro, _
        leftPos:=currentLeft, _
        topPos:=target.Top, _
        buttonWidth:=buttonWidth, _
        buttonHeight:=target.Height, _
        isBold:=False

    currentLeft = currentLeft + buttonWidth + gap

    AddReadoutBox _
        ws:=ws, _
        shapeName:=readoutName, _
        leftPos:=currentLeft, _
        topPos:=target.Top, _
        boxWidth:=readoutWidth, _
        boxHeight:=target.Height

    currentLeft = currentLeft + readoutWidth + gap

    AddActionButton _
        ws:=ws, _
        shapeName:=rightButtonName, _
        caption:=rightCaption, _
        macroName:=rightMacro, _
        leftPos:=currentLeft, _
        topPos:=target.Top, _
        buttonWidth:=buttonWidth, _
        buttonHeight:=target.Height, _
        isBold:=False

End Sub

Private Sub AddControlLabel( _
    ByVal ws As Excel.Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal labelWidth As Double, _
    ByVal labelHeight As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=labelWidth, _
        Height:=labelHeight)

    shp.Name = shapeName
    shp.Placement = xlMoveAndSize

    With shp

        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.Text = caption
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.WordWrap = msoFalse

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignLeft

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        With .TextFrame2.TextRange.Font

            .Name = "Arial"
            .Size = 8
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)

        End With

    End With

End Sub

'===============================================================================
' PERIOD BUTTONS
'===============================================================================
Private Sub CreatePeriodButtons(ByVal ws As Excel.Worksheet)

    Dim target As Excel.Range

    Dim captions As Variant
    Dim names As Variant
    Dim macros As Variant

    Dim i As Long
    Dim gap As Double
    Dim buttonWidth As Double
    Dim currentLeft As Double

    Set target = ws.Range("C21:R21")

    captions = Array( _
        "1W", "1M", "3M", "6M", "1Y", _
        "3Y", "5Y", "MAX", "AUTO Y")

    names = Array( _
        BUTTON_1W, BUTTON_1M, BUTTON_3M, _
        BUTTON_6M, BUTTON_1Y, BUTTON_3Y, _
        BUTTON_5Y, BUTTON_MAX, BUTTON_AUTO_Y)

    macros = Array( _
        "SetWindow1W", "SetWindow1M", "SetWindow3M", _
        "SetWindow6M", "SetWindow1Y", "SetWindow3Y", _
        "SetWindow5Y", "SetWindowMaximum", _
        "ResetYAxisDashboard")

    gap = 1

    buttonWidth = _
        (target.Width - _
         gap * (UBound(captions) - LBound(captions))) / _
        (UBound(captions) - LBound(captions) + 1)

    currentLeft = target.Left

    For i = LBound(captions) To UBound(captions)

        AddActionButton _
            ws:=ws, _
            shapeName:=CStr(names(i)), _
            caption:=CStr(captions(i)), _
            macroName:=CStr(macros(i)), _
            leftPos:=currentLeft, _
            topPos:=target.Top, _
            buttonWidth:=buttonWidth, _
            buttonHeight:=target.Height, _
            isBold:=False

        currentLeft = currentLeft + buttonWidth + gap

    Next i

End Sub

'===============================================================================
' BUTTONS AND READOUTS
'===============================================================================
Private Sub AddActionButton( _
    ByVal ws As Excel.Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal macroName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal buttonWidth As Double, _
    ByVal buttonHeight As Double, _
    ByVal isBold As Boolean)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=buttonWidth, _
        Height:=buttonHeight)

    shp.Name = shapeName
    shp.Placement = xlMoveAndSize
    shp.OnAction = MacroReference(macroName)

    With shp

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(242, 242, 242)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75

        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.Text = caption
        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        With .TextFrame2.TextRange.Font

            .Name = "Arial"
            .Size = 8

            If isBold Then
                .Bold = msoTrue
            Else
                .Bold = msoFalse
            End If

            .Fill.ForeColor.RGB = RGB(0, 0, 0)

        End With

    End With

End Sub

Private Sub AddReadoutBox( _
    ByVal ws As Excel.Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=leftPos, _
        Top:=topPos, _
        Width:=boxWidth, _
        Height:=boxHeight)

    shp.Name = shapeName
    shp.Placement = xlMoveAndSize

    With shp

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75

        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.Text = ""
        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        With .TextFrame2.TextRange.Font

            .Name = "Arial"
            .Size = 8
            .Bold = msoFalse
            .Fill.ForeColor.RGB = RGB(0, 0, 0)

        End With

    End With

End Sub

'===============================================================================
' CONTROL STATE
'===============================================================================
Private Sub UpdateControlLimits( _
    ByVal ws As Excel.Worksheet, _
    ByVal totalPoints As Long, _
    ByVal windowPoints As Long)

    Dim maximumStartIndex As Long

    maximumStartIndex = Application.WorksheetFunction.Max( _
        1, _
        totalPoints - windowPoints + 1)

    On Error Resume Next

    With ws.Shapes(CONTROL_X_RANGE).ControlFormat

        .Min = 1
        .Max = ScrollLimit(maximumStartIndex)
        .SmallChange = 1

        .LargeChange = Application.WorksheetFunction.Max( _
            1, _
            windowPoints \ 5)

    End With

    On Error GoTo 0

End Sub

Private Sub UpdateControlReadouts( _
    ByVal ws As Excel.Worksheet, _
    ByVal windowPoints As Long, _
    ByVal yZoom As Long, _
    ByVal markerSize As Long)

    SetShapeCaption _
        ws:=ws, _
        shapeName:=READOUT_X_ZOOM, _
        caption:=Format$(windowPoints, "#,##0")

    SetShapeCaption _
        ws:=ws, _
        shapeName:=READOUT_Y_ZOOM, _
        caption:=Format$(yZoom, "0") & "%"

    SetShapeCaption _
        ws:=ws, _
        shapeName:=READOUT_POINT_SIZE, _
        caption:=Format$(markerSize, "0")

End Sub

Private Sub SetShapeCaption( _
    ByVal ws As Excel.Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String)

    On Error Resume Next
    ws.Shapes(shapeName).TextFrame2.TextRange.Text = caption
    On Error GoTo 0

End Sub

Private Sub UpdatePeriodButtonStyles( _
    ByVal ws As Excel.Worksheet, _
    ByVal totalPoints As Long, _
    ByVal windowPoints As Long)

    SetPeriodButtonActive ws, BUTTON_1W, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_1W, totalPoints)

    SetPeriodButtonActive ws, BUTTON_1M, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_1M, totalPoints)

    SetPeriodButtonActive ws, BUTTON_3M, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_3M, totalPoints)

    SetPeriodButtonActive ws, BUTTON_6M, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_6M, totalPoints)

    SetPeriodButtonActive ws, BUTTON_1Y, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_1Y, totalPoints)

    SetPeriodButtonActive ws, BUTTON_3Y, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_3Y, totalPoints)

    SetPeriodButtonActive ws, BUTTON_5Y, _
        windowPoints = Application.WorksheetFunction.Min( _
            POINTS_5Y, totalPoints)

    SetPeriodButtonActive ws, BUTTON_MAX, _
        windowPoints = totalPoints

End Sub

Private Sub SetPeriodButtonActive( _
    ByVal ws As Excel.Worksheet, _
    ByVal shapeName As String, _
    ByVal isActive As Boolean)

    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If isActive Then

        shp.Fill.ForeColor.RGB = RGB(205, 205, 205)
        shp.Line.ForeColor.RGB = RGB(110, 110, 110)

        With shp.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With

    Else

        shp.Fill.ForeColor.RGB = RGB(242, 242, 242)
        shp.Line.ForeColor.RGB = RGB(170, 170, 170)

        With shp.TextFrame2.TextRange.Font
            .Bold = msoFalse
            .Fill.ForeColor.RGB = RGB(0, 0, 0)
        End With

    End If

End Sub

Private Sub UpdateXRangeStatus( _
    ByVal ws As Excel.Worksheet, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long, _
    ByVal totalPoints As Long)

    ws.Range("P20").Value = _
        Format$(startIndex, "#,##0") & _
        " - " & _
        Format$(endIndex, "#,##0") & _
        " / " & _
        Format$(totalPoints, "#,##0")

End Sub

'===============================================================================
' CHART
'===============================================================================
Private Sub DrawTimeSeriesChart( _
    ByVal ws As Excel.Worksheet, _
    ByRef displayNames() As String, _
    ByVal seriesCount As Long, _
    ByRef xAll() As Variant, _
    ByRef yAll() As Variant, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long, _
    ByVal axisMin As Double, _
    ByVal axisMax As Double, _
    ByVal yMajorUnit As Double, _
    ByVal yMinorUnit As Double, _
    ByVal markerSize As Long)

    Dim outputRange As Excel.Range
    Dim helperTop As Excel.Range

    Dim chartObject As Excel.ChartObject
    Dim ch As Excel.Chart

    Dim mainSeries As Excel.Series
    Dim latestSeries As Excel.Series

    Dim outputData() As Variant
    Dim xRange As Excel.Range
    Dim yRange As Excel.Range

    Dim seriesIndex As Long
    Dim pointIndex As Long
    Dim viewPoints As Long

    Dim firstDate As Double
    Dim lastDate As Double

    Dim xMajorUnit As Double
    Dim xMinorUnit As Double
    Dim xNumberFormat As String
    Dim xRotation As Long

    Dim yNumberFormat As String
    Dim seriesColor As Long
    Dim actualMarkerSize As Long

    Dim availablePlotWidth As Double
    Dim lastHelperRow As Long

    Dim latestIndex As Long
    Dim latestValue As Double
    Dim latestDate As Double

    Dim legendWidth As Double
    Dim legendHeight As Double

    viewPoints = endIndex - startIndex + 1

    If viewPoints <= 0 Then Exit Sub

    Set outputRange = ws.Range(CHART_OUTPUT_RANGE)
    Set helperTop = ws.Range(HELPER_DATE_COL & "1")

    lastHelperRow = ws.Range( _
        HELPER_DATE_COL & ws.Rows.Count).End(xlUp).Row

    If lastHelperRow < 2 Then
        lastHelperRow = 2
    End If

    ws.Range( _
        HELPER_DATE_COL & "1:" & _
        HELPER_SERIES_3_COL & lastHelperRow).ClearContents

    ReDim outputData( _
        1 To viewPoints, _
        1 To MAX_SERIES_COUNT + 1)

    For pointIndex = 1 To viewPoints

        outputData(pointIndex, 1) = _
            CDate(xAll(startIndex + pointIndex - 1))

        For seriesIndex = 1 To seriesCount

            outputData(pointIndex, seriesIndex + 1) = _
                yAll(startIndex + pointIndex - 1, seriesIndex)

        Next seriesIndex

    Next pointIndex

    helperTop.Value = "Date"

    For seriesIndex = 1 To seriesCount

        helperTop.Offset(0, seriesIndex).Value = _
            displayNames(seriesIndex)

    Next seriesIndex

    helperTop.Offset(1, 0).Resize( _
        viewPoints, _
        MAX_SERIES_COUNT + 1).Value = outputData

    Set xRange = helperTop.Offset(1, 0).Resize( _
        viewPoints, _
        1)

    ' Recreate the chart on each refresh.
    ' This guarantees that old titles and old data labels are removed.
    On Error Resume Next
    ws.ChartObjects(CHART_OBJECT_NAME).Delete
    On Error GoTo 0

    Set chartObject = ws.ChartObjects.Add( _
        Left:=outputRange.Left, _
        Top:=outputRange.Top, _
        Width:=outputRange.Width, _
        Height:=outputRange.Height)

    chartObject.Name = CHART_OBJECT_NAME
    chartObject.Placement = xlMoveAndSize

    Set ch = chartObject.Chart

    With ch

        .ChartType = xlXYScatterLines
        .PlotVisibleOnly = False

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        '-----------------------------------------------------------------------
        ' Main series
        '-----------------------------------------------------------------------
        For seriesIndex = 1 To seriesCount

            Set yRange = helperTop.Offset( _
                1, _
                seriesIndex).Resize(viewPoints, 1)

            Set mainSeries = .SeriesCollection.NewSeries

            seriesColor = GetSeriesColor(seriesIndex)

            With mainSeries

                ' Link the legend name explicitly to the header cell.
                .Name = "='" & _
                    Replace(ws.Name, "'", "''") & _
                    "'!" & _
                    helperTop.Offset(0, seriesIndex).Address

                .XValues = xRange
                .Values = yRange
                .Smooth = False

                On Error Resume Next
                .HasDataLabels = False
                On Error GoTo 0

                With .Format.Line

                    .Visible = msoTrue
                    .ForeColor.RGB = seriesColor
                    .Weight = 1.25

                End With

                If markerSize <= 0 Then

                    .MarkerStyle = xlMarkerStyleNone

                Else

                    actualMarkerSize = ClampLong( _
                        markerSize, _
                        2, _
                        MARKER_SIZE_MAX)

                    .MarkerStyle = xlMarkerStyleCircle
                    .MarkerSize = actualMarkerSize

                    .MarkerForegroundColor = seriesColor
                    .MarkerBackgroundColor = RGB(255, 255, 255)

                End If

            End With

        Next seriesIndex

        ' Completely remove the chart title.
        On Error Resume Next
        .ChartTitle.Delete
        On Error GoTo 0

        .HasTitle = False
        .HasLegend = True
        .DisplayBlanksAs = xlNotPlotted

        With .ChartArea.Format

            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(255, 255, 255)

            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(190, 190, 190)
            .Line.Weight = 0.5

        End With

        With .PlotArea.Format

            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Visible = msoFalse

        End With

    End With

    firstDate = CDbl(CDate(xRange.Cells(1, 1).Value))
    lastDate = CDbl(CDate(xRange.Cells(viewPoints, 1).Value))

    If lastDate <= firstDate Then
        lastDate = firstDate + 1
    End If

    yNumberFormat = NumberFormatFromStep(yMajorUnit)

    '---------------------------------------------------------------------------
    ' Y axis
    '---------------------------------------------------------------------------
    With ch.Axes(xlValue)

        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False

        .MinimumScale = axisMin
        .MaximumScale = axisMax

        .MajorUnitIsAuto = False
        .MajorUnit = yMajorUnit

        .MinorUnitIsAuto = False
        .MinorUnit = yMinorUnit

        .Crosses = xlAxisCrossesMinimum

        .TickLabels.NumberFormat = yNumberFormat
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8

        .HasMajorGridlines = True
        .HasMinorGridlines = False

        With .MajorGridlines.Format.Line

            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 220, 220)
            .Weight = 0.5
            .DashStyle = msoLineDash

        End With

        .Format.Line.ForeColor.RGB = RGB(150, 150, 150)
        .Format.Line.Weight = 0.5

    End With

    ch.Refresh
    DoEvents

    availablePlotWidth = ch.PlotArea.InsideWidth

    If availablePlotWidth <= 0 Then
        availablePlotWidth = chartObject.Width - 90
    End If

    GetAdaptiveDateAxisSettings _
        firstDate:=firstDate, _
        lastDate:=lastDate, _
        plotWidth:=availablePlotWidth, _
        majorUnit:=xMajorUnit, _
        numberFormat:=xNumberFormat, _
        labelRotation:=xRotation

    xMinorUnit = Application.WorksheetFunction.Max( _
        1, _
        xMajorUnit / 2)

    '---------------------------------------------------------------------------
    ' X axis
    '---------------------------------------------------------------------------
    With ch.Axes(xlCategory)

        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False

        .MinimumScale = firstDate
        .MaximumScale = lastDate

        .MajorUnitIsAuto = False
        .MajorUnit = xMajorUnit

        .MinorUnitIsAuto = False
        .MinorUnit = xMinorUnit

        .TickLabels.NumberFormat = xNumberFormat
        .TickLabels.Orientation = xRotation

        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 8

        .TickLabelPosition = xlTickLabelPositionLow

        .HasMajorGridlines = True
        .HasMinorGridlines = False

        With .MajorGridlines.Format.Line

            .Visible = msoTrue
            .ForeColor.RGB = RGB(225, 225, 225)
            .Weight = 0.4
            .DashStyle = msoLineDash

        End With

        .Format.Line.ForeColor.RGB = RGB(150, 150, 150)
        .Format.Line.Weight = 0.5

    End With

    '---------------------------------------------------------------------------
    ' Add one-point helper series for the latest visible value.
    '
    ' A separate series is used so data labels are never applied to all points.
    '---------------------------------------------------------------------------
    For seriesIndex = 1 To seriesCount

        latestIndex = FindLatestVisiblePointIndex( _
            yAll:=yAll, _
            seriesIndex:=seriesIndex, _
            startIndex:=startIndex, _
            endIndex:=endIndex)

        If latestIndex > 0 Then

            latestDate = CDbl(CDate(xAll(latestIndex)))
            latestValue = CDbl(yAll(latestIndex, seriesIndex))

            Set latestSeries = ch.SeriesCollection.NewSeries

            With latestSeries

                .Name = "__LATEST_" & CStr(seriesIndex)

                .XValues = Array(latestDate)
                .Values = Array(latestValue)

                .Format.Line.Visible = msoFalse

                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = Application.WorksheetFunction.Max( _
                    7, _
                    markerSize + 3)

                .MarkerForegroundColor = RGB(192, 0, 0)
                .MarkerBackgroundColor = RGB(192, 0, 0)

                .ApplyDataLabels

                With .DataLabels(1)

                    .ShowValue = True
                    .ShowSeriesName = False
                    .ShowCategoryName = False
                    .ShowLegendKey = False

                    .NumberFormat = yNumberFormat
                    .Text = Format$(latestValue, yNumberFormat)

                    .Position = xlLabelPositionLeft

                    .Font.Name = "Arial"
                    .Font.Size = 8
                    .Font.Bold = True
                    .Font.Color = RGB(192, 0, 0)

                    .Format.Fill.Visible = msoFalse
                    .Format.Line.Visible = msoFalse

                    Select Case seriesIndex

                        Case 2
                            .Top = .Top - 10

                        Case 3
                            .Top = .Top + 10

                    End Select

                End With

            End With

        End If

    Next seriesIndex

    ' Ensure no chart title was recreated.
    On Error Resume Next
    ch.ChartTitle.Delete
    ch.HasTitle = False
    On Error GoTo 0

    ch.Refresh
    DoEvents

    '---------------------------------------------------------------------------
    ' Legend
    '
    ' Delete legend entries created by the latest-point helper series.
    ' Only the original selected series remain in the legend.
    '---------------------------------------------------------------------------
    If ch.HasLegend Then

        On Error Resume Next

        Do While ch.Legend.LegendEntries.Count > seriesCount

            ch.Legend.LegendEntries( _
                ch.Legend.LegendEntries.Count).Delete

        Loop

        On Error GoTo 0

        With ch.Legend

            ' Right first forces a vertical legend.
            .Position = xlLegendPositionRight
            .IncludeInLayout = False

            .Font.Name = "Arial"
            .Font.Size = 8

            .Format.Fill.Visible = msoTrue
            .Format.Fill.Solid
            .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Format.Fill.Transparency = 0.12

            .Format.Line.Visible = msoFalse

        End With

        ch.Refresh
        DoEvents

        legendWidth = 125
        legendHeight = 8 + seriesCount * 15

        On Error Resume Next

        With ch.Legend

            .Left = ch.PlotArea.InsideLeft + 5
            .Top = ch.PlotArea.InsideTop + 5

            .Width = legendWidth
            .Height = legendHeight

        End With

        On Error GoTo 0

    End If

    ch.Refresh
    DoEvents

End Sub

'===============================================================================
' LATEST POINT
'===============================================================================
Private Function FindLatestVisiblePointIndex( _
    ByRef yAll() As Variant, _
    ByVal seriesIndex As Long, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long) As Long

    Dim pointIndex As Long

    For pointIndex = endIndex To startIndex Step -1

        If Not IsError(yAll(pointIndex, seriesIndex)) Then

            If IsNumeric(yAll(pointIndex, seriesIndex)) Then

                FindLatestVisiblePointIndex = pointIndex
                Exit Function

            End If

        End If

    Next pointIndex

End Function

'===============================================================================
' ADAPTIVE DATE AXIS
'===============================================================================
Private Sub GetAdaptiveDateAxisSettings( _
    ByVal firstDate As Double, _
    ByVal lastDate As Double, _
    ByVal plotWidth As Double, _
    ByRef majorUnit As Double, _
    ByRef numberFormat As String, _
    ByRef labelRotation As Long)

    Dim spanDays As Double
    Dim estimatedLabelWidth As Double
    Dim usableWidth As Double
    Dim maximumLabels As Long

    spanDays = lastDate - firstDate

    If spanDays <= 0 Then spanDays = 1

    If spanDays <= 45 Then

        numberFormat = "m/d"
        estimatedLabelWidth = 30

    ElseIf spanDays <= 400 Then

        If Year(CDate(firstDate)) = Year(CDate(lastDate)) Then

            numberFormat = "m/d"
            estimatedLabelWidth = 30

        Else

            numberFormat = "yyyy/m"
            estimatedLabelWidth = 45

        End If

    ElseIf spanDays <= 2200 Then

        numberFormat = "yyyy/m"
        estimatedLabelWidth = 45

    Else

        numberFormat = "yyyy"
        estimatedLabelWidth = 32

    End If

    usableWidth = Application.WorksheetFunction.Max( _
        120, _
        plotWidth)

    maximumLabels = CLng(Int( _
        usableWidth / (estimatedLabelWidth + 8)))

    maximumLabels = ClampLong( _
        maximumLabels, _
        4, _
        14)

    majorUnit = ChooseDateStepForMaximumLabels( _
        spanDays:=spanDays, _
        maximumLabels:=maximumLabels)

    labelRotation = 0

End Sub

Private Function ChooseDateStepForMaximumLabels( _
    ByVal spanDays As Double, _
    ByVal maximumLabels As Long) As Double

    Dim candidates As Variant

    Dim i As Long
    Dim candidate As Double
    Dim labelCount As Double

    candidates = Array( _
        1#, 2#, 3#, 5#, 7#, 10#, 14#, 21#, _
        30#, 45#, 60#, 90#, 120#, 180#, _
        365#, 730#, 1095#, 1825#)

    For i = LBound(candidates) To UBound(candidates)

        candidate = CDbl(candidates(i))
        labelCount = spanDays / candidate + 1

        If labelCount <= maximumLabels Then

            ChooseDateStepForMaximumLabels = candidate
            Exit Function

        End If

    Next i

    ChooseDateStepForMaximumLabels = _
        CDbl(candidates(UBound(candidates)))

End Function

'===============================================================================
' SERIES COLORS
'===============================================================================
Private Function GetSeriesColor(ByVal seriesIndex As Long) As Long

    Select Case seriesIndex

        Case 1
            GetSeriesColor = RGB(0, 102, 204)

        Case 2
            GetSeriesColor = RGB(255, 102, 0)

        Case 3
            GetSeriesColor = RGB(0, 153, 51)

        Case Else
            GetSeriesColor = RGB(80, 80, 80)

    End Select

End Function

'===============================================================================
' SELECTED SERIES
'===============================================================================
Private Function GetSelectedSeriesPairs( _
    ByVal ws As Excel.Worksheet, _
    ByRef seriesSheets() As String, _
    ByRef seriesNames() As String) As Long

    Dim candidateSheets(1 To MAX_SERIES_COUNT) As String
    Dim candidateSeries(1 To MAX_SERIES_COUNT) As String

    Dim candidateIndex As Long
    Dim acceptedIndex As Long
    Dim existingIndex As Long

    Dim isDuplicate As Boolean

    candidateSheets(1) = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_1).Value))

    candidateSeries(1) = Trim$(CStr( _
        ws.Range(CELL_SERIES_1).Value))

    candidateSheets(2) = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_2).Value))

    candidateSeries(2) = Trim$(CStr( _
        ws.Range(CELL_SERIES_2).Value))

    candidateSheets(3) = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_3).Value))

    candidateSeries(3) = Trim$(CStr( _
        ws.Range(CELL_SERIES_3).Value))

    For candidateIndex = 1 To MAX_SERIES_COUNT

        If Len(candidateSheets(candidateIndex)) > 0 And _
           Len(candidateSeries(candidateIndex)) > 0 And _
           StrComp( _
               candidateSeries(candidateIndex), _
               NONE_SERIES_TEXT, _
               vbTextCompare) <> 0 Then

            isDuplicate = False

            For existingIndex = 1 To acceptedIndex

                If StrComp( _
                    seriesSheets(existingIndex), _
                    candidateSheets(candidateIndex), _
                    vbTextCompare) = 0 And _
                   StrComp( _
                    seriesNames(existingIndex), _
                    candidateSeries(candidateIndex), _
                    vbTextCompare) = 0 Then

                    isDuplicate = True
                    Exit For

                End If

            Next existingIndex

            If Not isDuplicate Then

                acceptedIndex = acceptedIndex + 1

                seriesSheets(acceptedIndex) = _
                    candidateSheets(candidateIndex)

                seriesNames(acceptedIndex) = _
                    candidateSeries(candidateIndex)

            End If

        End If

    Next candidateIndex

    GetSelectedSeriesPairs = acceptedIndex

End Function

Private Sub BuildDisplayNames( _
    ByRef seriesSheets() As String, _
    ByRef seriesNames() As String, _
    ByVal seriesCount As Long, _
    ByRef displayNames() As String)

    Dim i As Long
    Dim j As Long
    Dim duplicateCount As Long

    For i = 1 To seriesCount

        duplicateCount = 0

        For j = 1 To seriesCount

            If StrComp( _
                seriesNames(i), _
                seriesNames(j), _
                vbTextCompare) = 0 Then

                duplicateCount = duplicateCount + 1

            End If

        Next j

        If duplicateCount > 1 Then

            displayNames(i) = _
                seriesNames(i) & _
                " [" & seriesSheets(i) & "]"

        Else

            displayNames(i) = seriesNames(i)

        End If

    Next i

End Sub

'===============================================================================
' LOAD AND ALIGN DATA
'===============================================================================
Private Function LoadAlignedDashboardData( _
    ByRef seriesSheets() As String, _
    ByRef seriesNames() As String, _
    ByVal seriesCount As Long, _
    ByRef xAll() As Variant, _
    ByRef yAll() As Variant, _
    ByRef dataCount As Long, _
    ByRef errorMessage As String) As Boolean

    Dim datesBySeries(1 To MAX_SERIES_COUNT) As Variant
    Dim valuesBySeries(1 To MAX_SERIES_COUNT) As Variant
    Dim counts(1 To MAX_SERIES_COUNT) As Long

    Dim dateDictionary As Object
    Dim dateIndexDictionary As Object

    Dim currentDates As Variant
    Dim currentValues As Variant
    Dim dictionaryItems As Variant

    Dim seriesIndex As Long
    Dim pointIndex As Long
    Dim unionIndex As Long

    Dim dateKeyText As String

    Set dateDictionary = CreateObject("Scripting.Dictionary")
    Set dateIndexDictionary = CreateObject("Scripting.Dictionary")

    For seriesIndex = 1 To seriesCount

        If Not LoadSingleSeriesData( _
            dataSheetName:=seriesSheets(seriesIndex), _
            seriesName:=seriesNames(seriesIndex), _
            datesOut:=datesBySeries(seriesIndex), _
            valuesOut:=valuesBySeries(seriesIndex), _
            itemCount:=counts(seriesIndex), _
            errorMessage:=errorMessage) Then

            Exit Function

        End If

        currentDates = datesBySeries(seriesIndex)

        For pointIndex = 1 To counts(seriesIndex)

            dateKeyText = DateKey(currentDates(pointIndex))

            If Not dateDictionary.Exists(dateKeyText) Then

                dateDictionary.Add _
                    dateKeyText, _
                    CDbl(CDate(currentDates(pointIndex)))

            End If

        Next pointIndex

    Next seriesIndex

    dataCount = dateDictionary.Count

    If dataCount = 0 Then

        errorMessage = "No valid dates were found."
        Exit Function

    End If

    ReDim xAll(1 To dataCount)
    ReDim yAll(1 To dataCount, 1 To MAX_SERIES_COUNT)

    dictionaryItems = dateDictionary.Items

    For pointIndex = 0 To dataCount - 1

        xAll(pointIndex + 1) = _
            CDate(CDbl(dictionaryItems(pointIndex)))

    Next pointIndex

    QuickSortDates _
        values:=xAll, _
        firstIndex:=1, _
        lastIndex:=dataCount

    For pointIndex = 1 To dataCount

        dateIndexDictionary.Add _
            DateKey(xAll(pointIndex)), _
            pointIndex

        For seriesIndex = 1 To MAX_SERIES_COUNT

            yAll(pointIndex, seriesIndex) = _
                CVErr(xlErrNA)

        Next seriesIndex

    Next pointIndex

    For seriesIndex = 1 To seriesCount

        currentDates = datesBySeries(seriesIndex)
        currentValues = valuesBySeries(seriesIndex)

        For pointIndex = 1 To counts(seriesIndex)

            dateKeyText = DateKey(currentDates(pointIndex))

            If dateIndexDictionary.Exists(dateKeyText) Then

                unionIndex = CLng( _
                    dateIndexDictionary(dateKeyText))

                yAll(unionIndex, seriesIndex) = _
                    currentValues(pointIndex)

            End If

        Next pointIndex

    Next seriesIndex

    LoadAlignedDashboardData = True

End Function

Private Function LoadSingleSeriesData( _
    ByVal dataSheetName As String, _
    ByVal seriesName As String, _
    ByRef datesOut As Variant, _
    ByRef valuesOut As Variant, _
    ByRef itemCount As Long, _
    ByRef errorMessage As String) As Boolean

    Dim ws As Excel.Worksheet

    Dim seriesColumn As Long
    Dim lastRow As Long
    Dim rowNumber As Long

    Dim dateValue As Variant
    Dim numericValue As Variant

    Dim dateArray() As Variant
    Dim valueArray() As Variant

    Set ws = GetWorksheet(dataSheetName)

    If ws Is Nothing Then

        errorMessage = _
            "Data sheet not found: " & dataSheetName

        Exit Function

    End If

    seriesColumn = FindSeriesColumn( _
        ws, _
        seriesName)

    If seriesColumn = 0 Then

        errorMessage = _
            "Series not found: " & _
            dataSheetName & " / " & seriesName

        Exit Function

    End If

    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, seriesColumn).End(xlUp).Row)

    If lastRow < 2 Then

        errorMessage = _
            "No data found: " & _
            dataSheetName & " / " & seriesName

        Exit Function

    End If

    ReDim dateArray(1 To lastRow - 1)
    ReDim valueArray(1 To lastRow - 1)

    itemCount = 0

    For rowNumber = 2 To lastRow

        dateValue = ws.Cells(rowNumber, 1).Value
        numericValue = ws.Cells(rowNumber, seriesColumn).Value

        If Not IsError(dateValue) And _
           IsDate(dateValue) And _
           IsValidNumericValue(numericValue) Then

            itemCount = itemCount + 1

            dateArray(itemCount) = _
                CDate(Int(CDbl(CDate(dateValue))))

            valueArray(itemCount) = _
                CDbl(numericValue)

        End If

    Next rowNumber

    If itemCount = 0 Then

        errorMessage = _
            "No valid observations found: " & _
            dataSheetName & " / " & seriesName

        Exit Function

    End If

    ReDim Preserve dateArray(1 To itemCount)
    ReDim Preserve valueArray(1 To itemCount)

    datesOut = dateArray
    valuesOut = valueArray

    LoadSingleSeriesData = True

End Function

Private Function IsValidNumericValue( _
    ByVal valueItem As Variant) As Boolean

    If IsError(valueItem) Then Exit Function
    If IsEmpty(valueItem) Then Exit Function
    If Len(Trim$(CStr(valueItem))) = 0 Then Exit Function

    IsValidNumericValue = IsNumeric(valueItem)

End Function

Private Function DateKey(ByVal dateValue As Variant) As String

    DateKey = CStr(CLng(Int(CDbl(CDate(dateValue)))))

End Function

Private Sub QuickSortDates( _
    ByRef values() As Variant, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long)

    Dim lowIndex As Long
    Dim highIndex As Long

    Dim pivotValue As Double
    Dim temporaryValue As Variant

    lowIndex = firstIndex
    highIndex = lastIndex

    pivotValue = CDbl(values( _
        (firstIndex + lastIndex) \ 2))

    Do While lowIndex <= highIndex

        Do While CDbl(values(lowIndex)) < pivotValue
            lowIndex = lowIndex + 1
        Loop

        Do While CDbl(values(highIndex)) > pivotValue
            highIndex = highIndex - 1
        Loop

        If lowIndex <= highIndex Then

            temporaryValue = values(lowIndex)
            values(lowIndex) = values(highIndex)
            values(highIndex) = temporaryValue

            lowIndex = lowIndex + 1
            highIndex = highIndex - 1

        End If

    Loop

    If firstIndex < highIndex Then

        QuickSortDates _
            values, _
            firstIndex, _
            highIndex

    End If

    If lowIndex < lastIndex Then

        QuickSortDates _
            values, _
            lowIndex, _
            lastIndex

    End If

End Sub

'===============================================================================
' VISIBLE MINIMUM AND MAXIMUM
'===============================================================================
Private Sub GetVisibleGlobalMinMax( _
    ByRef yAll() As Variant, _
    ByVal seriesCount As Long, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long, _
    ByRef minimumValue As Double, _
    ByRef maximumValue As Double)

    Dim pointIndex As Long
    Dim seriesIndex As Long

    Dim currentValue As Double
    Dim foundValue As Boolean

    For pointIndex = startIndex To endIndex

        For seriesIndex = 1 To seriesCount

            If Not IsError(yAll(pointIndex, seriesIndex)) Then

                If IsNumeric(yAll(pointIndex, seriesIndex)) Then

                    currentValue = _
                        CDbl(yAll(pointIndex, seriesIndex))

                    If Not foundValue Then

                        minimumValue = currentValue
                        maximumValue = currentValue
                        foundValue = True

                    Else

                        If currentValue < minimumValue Then
                            minimumValue = currentValue
                        End If

                        If currentValue > maximumValue Then
                            maximumValue = currentValue
                        End If

                    End If

                End If

            End If

        Next seriesIndex

    Next pointIndex

    If Not foundValue Then

        Err.Raise _
            vbObjectError + 2001, , _
            "No visible numeric data were found."

    End If

End Sub

'===============================================================================
' Y AXIS
'===============================================================================
Private Sub CalculateYAxisRange( _
    ByVal dataMin As Double, _
    ByVal dataMax As Double, _
    ByVal yZoomPercent As Long, _
    ByRef axisMin As Double, _
    ByRef axisMax As Double)

    Dim dataSpan As Double
    Dim baseSpan As Double
    Dim axisSpan As Double
    Dim axisCenter As Double

    dataSpan = dataMax - dataMin

    If dataSpan <= 0 Then

        dataSpan = Abs(dataMax) * 0.1

        If dataSpan <= 0 Then
            dataSpan = 0.01
        End If

    End If

    baseSpan = dataSpan * 1.12
    axisSpan = baseSpan * yZoomPercent / 100#

    axisCenter = (dataMin + dataMax) / 2

    axisMin = axisCenter - axisSpan / 2
    axisMax = axisCenter + axisSpan / 2

    If axisMax <= axisMin Then
        axisMax = axisMin + 0.01
    End If

End Sub

Private Sub GetNiceYAxisScale( _
    ByVal rawMinimum As Double, _
    ByVal rawMaximum As Double, _
    ByRef niceMinimum As Double, _
    ByRef niceMaximum As Double, _
    ByRef majorUnit As Double, _
    ByRef minorUnit As Double)

    Dim spanValue As Double

    spanValue = rawMaximum - rawMinimum

    If spanValue <= 0 Then

        spanValue = Abs(rawMaximum) * 0.1

        If spanValue <= 0 Then
            spanValue = 0.01
        End If

        rawMinimum = rawMinimum - spanValue / 2
        rawMaximum = rawMaximum + spanValue / 2

    End If

    majorUnit = ChooseNiceNumericStep( _
        spanValue:=rawMaximum - rawMinimum, _
        targetIntervals:=12)

    niceMinimum = FloorToStep( _
        rawMinimum, _
        majorUnit)

    niceMaximum = CeilingToStep( _
        rawMaximum, _
        majorUnit)

    If niceMaximum <= niceMinimum Then
        niceMaximum = niceMinimum + majorUnit
    End If

    minorUnit = majorUnit / 2

End Sub

Private Function ChooseNiceNumericStep( _
    ByVal spanValue As Double, _
    ByVal targetIntervals As Long) As Double

    Dim bases As Variant

    Dim baseIndex As Long
    Dim exponentValue As Long
    Dim centralExponent As Long

    Dim candidate As Double
    Dim intervalCount As Double
    Dim score As Double

    Dim bestCandidate As Double
    Dim bestScore As Double

    If spanValue <= 0 Then

        ChooseNiceNumericStep = 1
        Exit Function

    End If

    bases = Array(1#, 2#, 2.5, 5#, 10#)

    centralExponent = Int( _
        Log(spanValue / targetIntervals) / Log(10#))

    bestScore = 1E+99

    For exponentValue = _
        centralExponent - 2 To centralExponent + 2

        For baseIndex = LBound(bases) To UBound(bases)

            candidate = _
                CDbl(bases(baseIndex)) * _
                (10# ^ exponentValue)

            If candidate > 0 Then

                intervalCount = spanValue / candidate
                score = Abs(intervalCount - targetIntervals)

                If intervalCount < 9 Then

                    score = score + _
                        100 + _
                        (9 - intervalCount) * 10

                End If

                If intervalCount > 15 Then

                    score = score + _
                        100 + _
                        (intervalCount - 15) * 10

                End If

                If score < bestScore Then

                    bestScore = score
                    bestCandidate = candidate

                End If

            End If

        Next baseIndex

    Next exponentValue

    If bestCandidate <= 0 Then
        bestCandidate = spanValue / targetIntervals
    End If

    ChooseNiceNumericStep = bestCandidate

End Function

Private Function FloorToStep( _
    ByVal valueItem As Double, _
    ByVal stepValue As Double) As Double

    If stepValue <= 0 Then

        FloorToStep = valueItem

    Else

        FloorToStep = _
            Int(valueItem / stepValue) * stepValue

    End If

End Function

Private Function CeilingToStep( _
    ByVal valueItem As Double, _
    ByVal stepValue As Double) As Double

    If stepValue <= 0 Then

        CeilingToStep = valueItem

    Else

        CeilingToStep = _
            -Int(-valueItem / stepValue) * stepValue

    End If

End Function

Private Function NumberFormatFromStep( _
    ByVal stepValue As Double) As String

    Dim decimalPlaces As Long
    Dim scaledValue As Double

    If stepValue <= 0 Then

        NumberFormatFromStep = "0.000"
        Exit Function

    End If

    If stepValue >= 1000 Then

        NumberFormatFromStep = "#,##0"
        Exit Function

    End If

    decimalPlaces = _
        Application.WorksheetFunction.Max( _
            0, _
            -Int(Log(stepValue) / Log(10#)))

    scaledValue = _
        stepValue * (10# ^ decimalPlaces)

    If Abs( _
        scaledValue - Round(scaledValue, 0)) > 0.0000001 Then

        decimalPlaces = decimalPlaces + 1

    End If

    If stepValue < 1 And decimalPlaces < 3 Then
        decimalPlaces = 3
    End If

    If decimalPlaces > 8 Then
        decimalPlaces = 8
    End If

    If decimalPlaces = 0 Then

        NumberFormatFromStep = "0"

    Else

        NumberFormatFromStep = _
            "0." & String$(decimalPlaces, "0")

    End If

End Function

'===============================================================================
' SHEET DROPDOWNS
'===============================================================================
Private Sub BuildSheetDropdown(ByVal ws As Excel.Worksheet)

    Dim sourceSheet As Excel.Worksheet
    Dim listCount As Long

    ws.Range( _
        SHEET_LIST_COL & "1:" & _
        SHEET_LIST_COL & "5000").ClearContents

    For Each sourceSheet In ThisWorkbook.Worksheets

        If StrComp( _
            sourceSheet.Name, _
            DASHBOARD_SHEET, _
            vbTextCompare) <> 0 Then

            listCount = listCount + 1

            ws.Range( _
                SHEET_LIST_COL & listCount).Value = _
                sourceSheet.Name

        End If

    Next sourceSheet

    DeleteWorkbookName NAME_SHEET_LIST

    If listCount > 0 Then

        ThisWorkbook.Names.Add _
            Name:=NAME_SHEET_LIST, _
            RefersTo:="='" & _
                Replace(ws.Name, "'", "''") & _
                "'!$" & SHEET_LIST_COL & "$1:$" & _
                SHEET_LIST_COL & "$" & listCount

        SetListValidation _
            ws.Range(CELL_DATA_SHEET_1), _
            "=" & NAME_SHEET_LIST

        SetListValidation _
            ws.Range(CELL_DATA_SHEET_2), _
            "=" & NAME_SHEET_LIST

        SetListValidation _
            ws.Range(CELL_DATA_SHEET_3), _
            "=" & NAME_SHEET_LIST

    End If

End Sub

Private Sub EnsureValidDataSheetSelections( _
    ByVal ws As Excel.Worksheet)

    Dim fallbackSheet As String
    Dim sheet1 As String
    Dim sheet2 As String
    Dim sheet3 As String

    If WorksheetExists("JGBData") Then

        fallbackSheet = "JGBData"

    Else

        fallbackSheet = _
            Trim$(CStr(ws.Range( _
                SHEET_LIST_COL & "1").Value))

    End If

    sheet1 = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_1).Value))

    If Not WorksheetExists(sheet1) Then
        sheet1 = fallbackSheet
    End If

    ws.Range(CELL_DATA_SHEET_1).Value = sheet1

    sheet2 = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_2).Value))

    If Not WorksheetExists(sheet2) Then
        sheet2 = sheet1
    End If

    ws.Range(CELL_DATA_SHEET_2).Value = sheet2

    sheet3 = Trim$(CStr( _
        ws.Range(CELL_DATA_SHEET_3).Value))

    If Not WorksheetExists(sheet3) Then
        sheet3 = sheet1
    End If

    ws.Range(CELL_DATA_SHEET_3).Value = sheet3

End Sub

'===============================================================================
' SERIES DROPDOWNS
'===============================================================================
Private Sub BuildAllSeriesDropdowns(ByVal ws As Excel.Worksheet)

    BuildSeriesDropdownForSlot _
        ws:=ws, _
        dataSheetName:=Trim$(CStr( _
            ws.Range(CELL_DATA_SHEET_1).Value)), _
        targetCell:=CELL_SERIES_1, _
        listColumn:=SERIES_LIST_1_COL, _
        workbookName:=NAME_SERIES_LIST_1

    BuildSeriesDropdownForSlot _
        ws:=ws, _
        dataSheetName:=Trim$(CStr( _
            ws.Range(CELL_DATA_SHEET_2).Value)), _
        targetCell:=CELL_SERIES_2, _
        listColumn:=SERIES_LIST_2_COL, _
        workbookName:=NAME_SERIES_LIST_2

    BuildSeriesDropdownForSlot _
        ws:=ws, _
        dataSheetName:=Trim$(CStr( _
            ws.Range(CELL_DATA_SHEET_3).Value)), _
        targetCell:=CELL_SERIES_3, _
        listColumn:=SERIES_LIST_3_COL, _
        workbookName:=NAME_SERIES_LIST_3

End Sub

Private Sub BuildSeriesDropdownForSlot( _
    ByVal ws As Excel.Worksheet, _
    ByVal dataSheetName As String, _
    ByVal targetCell As String, _
    ByVal listColumn As String, _
    ByVal workbookName As String)

    Dim dataSheet As Excel.Worksheet

    Dim lastColumn As Long
    Dim columnNumber As Long
    Dim listCount As Long

    Dim headerText As String

    ws.Range( _
        listColumn & "1:" & _
        listColumn & "5000").ClearContents

    DeleteWorkbookName workbookName

    Set dataSheet = GetWorksheet(dataSheetName)

    If dataSheet Is Nothing Then Exit Sub

    listCount = 1
    ws.Range(listColumn & "1").Value = NONE_SERIES_TEXT

    lastColumn = dataSheet.Cells( _
        1, _
        dataSheet.Columns.Count).End(xlToLeft).Column

    For columnNumber = 2 To lastColumn

        headerText = Trim$(CStr( _
            dataSheet.Cells(1, columnNumber).Value))

        If Len(headerText) > 0 Then

            listCount = listCount + 1

            ws.Range( _
                listColumn & listCount).Value = _
                headerText

        End If

    Next columnNumber

    ThisWorkbook.Names.Add _
        Name:=workbookName, _
        RefersTo:="='" & _
            Replace(ws.Name, "'", "''") & _
            "'!$" & listColumn & "$1:$" & _
            listColumn & "$" & listCount

    SetListValidation _
        target:=ws.Range(targetCell), _
        formulaText:="=" & workbookName

End Sub

Private Sub EnsureValidSeriesSelections(ByVal ws As Excel.Worksheet)

    EnsureSeriesForSlot _
        ws:=ws, _
        dataSheetCell:=CELL_DATA_SHEET_1, _
        seriesCell:=CELL_SERIES_1, _
        listColumn:=SERIES_LIST_1_COL, _
        seriesRequired:=True

    EnsureSeriesForSlot _
        ws:=ws, _
        dataSheetCell:=CELL_DATA_SHEET_2, _
        seriesCell:=CELL_SERIES_2, _
        listColumn:=SERIES_LIST_2_COL, _
        seriesRequired:=False

    EnsureSeriesForSlot _
        ws:=ws, _
        dataSheetCell:=CELL_DATA_SHEET_3, _
        seriesCell:=CELL_SERIES_3, _
        listColumn:=SERIES_LIST_3_COL, _
        seriesRequired:=False

    RemoveDuplicateSeriesPair _
        ws:=ws, _
        earlierSheetCell:=CELL_DATA_SHEET_1, _
        earlierSeriesCell:=CELL_SERIES_1, _
        laterSheetCell:=CELL_DATA_SHEET_2, _
        laterSeriesCell:=CELL_SERIES_2

    RemoveDuplicateSeriesPair _
        ws:=ws, _
        earlierSheetCell:=CELL_DATA_SHEET_1, _
        earlierSeriesCell:=CELL_SERIES_1, _
        laterSheetCell:=CELL_DATA_SHEET_3, _
        laterSeriesCell:=CELL_SERIES_3

    RemoveDuplicateSeriesPair _
        ws:=ws, _
        earlierSheetCell:=CELL_DATA_SHEET_2, _
        earlierSeriesCell:=CELL_SERIES_2, _
        laterSheetCell:=CELL_DATA_SHEET_3, _
        laterSeriesCell:=CELL_SERIES_3

End Sub

Private Sub EnsureSeriesForSlot( _
    ByVal ws As Excel.Worksheet, _
    ByVal dataSheetCell As String, _
    ByVal seriesCell As String, _
    ByVal listColumn As String, _
    ByVal seriesRequired As Boolean)

    Dim dataSheet As Excel.Worksheet
    Dim seriesName As String

    Set dataSheet = GetWorksheet( _
        Trim$(CStr(ws.Range(dataSheetCell).Value)))

    seriesName = Trim$(CStr( _
        ws.Range(seriesCell).Value))

    If dataSheet Is Nothing Then

        ws.Range(seriesCell).Value = NONE_SERIES_TEXT
        Exit Sub

    End If

    If seriesRequired Then

        If Len(seriesName) = 0 Or _
           StrComp( _
               seriesName, _
               NONE_SERIES_TEXT, _
               vbTextCompare) = 0 Or _
           FindSeriesColumn(dataSheet, seriesName) = 0 Then

            If Len(Trim$(CStr( _
                ws.Range(listColumn & "2").Value))) > 0 Then

                ws.Range(seriesCell).Value = _
                    ws.Range(listColumn & "2").Value

            End If

        End If

    Else

        If Len(seriesName) = 0 Or _
           StrComp( _
               seriesName, _
               NONE_SERIES_TEXT, _
               vbTextCompare) = 0 Then

            ws.Range(seriesCell).Value = NONE_SERIES_TEXT

        ElseIf FindSeriesColumn( _
            dataSheet, _
            seriesName) = 0 Then

            ws.Range(seriesCell).Value = NONE_SERIES_TEXT

        End If

    End If

End Sub

Private Sub RemoveDuplicateSeriesPair( _
    ByVal ws As Excel.Worksheet, _
    ByVal earlierSheetCell As String, _
    ByVal earlierSeriesCell As String, _
    ByVal laterSheetCell As String, _
    ByVal laterSeriesCell As String)

    Dim earlierSheet As String
    Dim earlierSeries As String
    Dim laterSheet As String
    Dim laterSeries As String

    earlierSheet = Trim$(CStr( _
        ws.Range(earlierSheetCell).Value))

    earlierSeries = Trim$(CStr( _
        ws.Range(earlierSeriesCell).Value))

    laterSheet = Trim$(CStr( _
        ws.Range(laterSheetCell).Value))

    laterSeries = Trim$(CStr( _
        ws.Range(laterSeriesCell).Value))

    If StrComp( _
        laterSeries, _
        NONE_SERIES_TEXT, _
        vbTextCompare) = 0 Then Exit Sub

    If StrComp( _
        earlierSheet, _
        laterSheet, _
        vbTextCompare) = 0 And _
       StrComp( _
        earlierSeries, _
        laterSeries, _
        vbTextCompare) = 0 Then

        ws.Range(laterSeriesCell).Value = NONE_SERIES_TEXT

    End If

End Sub

Private Sub SetListValidation( _
    ByVal target As Excel.Range, _
    ByVal formulaText As String)

    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0

    target.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=formulaText

    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.ShowError = True

End Sub

'===============================================================================
' COMMON FUNCTIONS
'===============================================================================
Private Function FindSeriesColumn( _
    ByVal ws As Excel.Worksheet, _
    ByVal seriesName As String) As Long

    Dim lastColumn As Long
    Dim columnNumber As Long

    If Len(Trim$(seriesName)) = 0 Then Exit Function

    lastColumn = ws.Cells( _
        1, _
        ws.Columns.Count).End(xlToLeft).Column

    For columnNumber = 2 To lastColumn

        If StrComp( _
            Trim$(CStr( _
                ws.Cells(1, columnNumber).Value)), _
            Trim$(seriesName), _
            vbTextCompare) = 0 Then

            FindSeriesColumn = columnNumber
            Exit Function

        End If

    Next columnNumber

End Function

Private Function GetWorksheet( _
    ByVal sheetName As String) As Excel.Worksheet

    If Len(Trim$(sheetName)) = 0 Then Exit Function

    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

End Function

Private Function GetOrCreateWorksheet( _
    ByVal sheetName As String) As Excel.Worksheet

    Set GetOrCreateWorksheet = GetWorksheet(sheetName)

    If GetOrCreateWorksheet Is Nothing Then

        Set GetOrCreateWorksheet = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        GetOrCreateWorksheet.Name = sheetName

    End If

End Function

Private Function WorksheetExists( _
    ByVal sheetName As String) As Boolean

    Dim ws As Excel.Worksheet

    Set ws = GetWorksheet(sheetName)

    WorksheetExists = Not ws Is Nothing

End Function

Private Function ClampLong( _
    ByVal valueItem As Long, _
    ByVal minimumValue As Long, _
    ByVal maximumValue As Long) As Long

    If maximumValue < minimumValue Then
        maximumValue = minimumValue
    End If

    If valueItem < minimumValue Then

        ClampLong = minimumValue

    ElseIf valueItem > maximumValue Then

        ClampLong = maximumValue

    Else

        ClampLong = valueItem

    End If

End Function

Private Function GetLongOrDefault( _
    ByVal sourceValue As Variant, _
    ByVal defaultValue As Long) As Long

    If IsNumeric(sourceValue) Then

        GetLongOrDefault = CLng(sourceValue)

    Else

        GetLongOrDefault = defaultValue

    End If

End Function

Private Function SheetCellReference( _
    ByVal ws As Excel.Worksheet, _
    ByVal cellAddress As String) As String

    SheetCellReference = _
        "'" & _
        Replace(ws.Name, "'", "''") & _
        "'!" & _
        ws.Range(cellAddress).Address

End Function

Private Function MacroReference( _
    ByVal macroName As String) As String

    MacroReference = _
        "'" & _
        Replace(ThisWorkbook.Name, "'", "''") & _
        "'!" & _
        macroName

End Function

Private Function ScrollLimit( _
    ByVal valueItem As Long) As Long

    If valueItem < 1 Then

        ScrollLimit = 1

    ElseIf valueItem > 30000 Then

        ScrollLimit = 30000

    Else

        ScrollLimit = valueItem

    End If

End Function

'===============================================================================
' CLEANUP
'===============================================================================
Private Sub DeleteDashboardObjects(ByVal ws As Excel.Worksheet)

    Dim i As Long
    Dim shapeName As String

    On Error Resume Next
    ws.ChartObjects(CHART_OBJECT_NAME).Delete
    ws.ChartObjects("TimeSeriesChart_Main").Delete
    On Error GoTo 0

    For i = ws.Shapes.Count To 1 Step -1

        shapeName = ws.Shapes(i).Name

        If LCase$(Left$(shapeName, 3)) = "ts_" Or _
           StrComp(shapeName, "btnPlot", vbTextCompare) = 0 Or _
           StrComp(shapeName, "Plot", vbTextCompare) = 0 Then

            ws.Shapes(i).Delete

        End If

    Next i

End Sub

Private Sub DeleteWorkbookName(ByVal nameText As String)

    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0

End Sub
