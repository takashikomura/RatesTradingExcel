Option Explicit

'===============================================================================
' CHART GRID DASHBOARD
'
' Grid size:
'   Each dashboard occupies A1:S22 equivalent.
'   Horizontal stride = 19 columns.
'   Vertical stride   = 23 rows (one blank row between dashboard rows).
'
' Source-data convention:
'   Row 1       : series headers
'   Column A    : dates
'   Column B... : values
'
' Main features:
'   - InputBox specification such as 4*3, 4x3 or 4ﾃ・
'   - Maximum three independently selected series per grid
'   - Data sheet can be selected separately for every series
'   - Independent X range, X zoom, Y zoom and marker-size state per grid
'   - 1W / 1M / 3M / 6M / 1Y / 3Y / 5Y / MAX / AUTO Y buttons
'   - Latest visible point highlighted in red with a value label
'   - Compact internal legend
'   - No chart title
'   - Ordinary display operations use only the per-grid cache
'   - UPDATE recalculates only the source sheets used by that grid
'
' Recommended standard-module name:
'   modChartGrid
'===============================================================================

'===============================================================================
' WORKSHEETS
'===============================================================================

Private Const GRID_SHEET_NAME As String = "ChartGrid"
Private Const STATE_SHEET_NAME As String = "__CG_SYS"
Private Const DATA_SHEET_NAME As String = "__CG_DATA"

'===============================================================================
' GRID GEOMETRY
'===============================================================================

Private Const GRID_COLUMN_COUNT As Long = 19          ' A:S
Private Const GRID_ROW_COUNT As Long = 22             ' 1:22
Private Const GRID_COLUMN_STRIDE As Long = 19
Private Const GRID_ROW_STRIDE As Long = 23             ' one blank row between grids

Private Const GRID_CHART_FIRST_REL_COL As Long = 3     ' C
Private Const GRID_CHART_LAST_REL_COL As Long = 19     ' S
Private Const GRID_CHART_FIRST_REL_ROW As Long = 2
Private Const GRID_CHART_LAST_REL_ROW As Long = 19

Private Const GRID_X_RANGE_REL_ROW As Long = 20
Private Const GRID_PERIOD_REL_ROW As Long = 21

Private Const SERIES_SLOT_COUNT As Long = 3
Private Const NONE_SERIES_TEXT As String = "(None)"

'===============================================================================
' WINDOW PRESETS
'===============================================================================

Private Const MIN_WINDOW_POINTS As Long = 5
Private Const DEFAULT_WINDOW_POINTS As Long = 66

Private Const POINTS_1W As Long = 5
Private Const POINTS_1M As Long = 22
Private Const POINTS_3M As Long = 66
Private Const POINTS_6M As Long = 132
Private Const POINTS_1Y As Long = 264
Private Const POINTS_3Y As Long = 792
Private Const POINTS_5Y As Long = 1320

'===============================================================================
' Y-AXIS AND MARKERS
'===============================================================================

Private Const DEFAULT_Y_ZOOM As Long = 100
Private Const Y_ZOOM_MIN As Long = 25
Private Const Y_ZOOM_MAX As Long = 400

Private Const DEFAULT_MARKER_SIZE As Long = 0
Private Const MARKER_SIZE_MIN As Long = 0
Private Const MARKER_SIZE_MAX As Long = 12

'===============================================================================
' STATE-SHEET COLUMNS
'===============================================================================

Private Const ST_GRID_ID As Long = 1
Private Const ST_GRID_ROW As Long = 2
Private Const ST_GRID_COL As Long = 3

Private Const ST_DATA_SHEET_1 As Long = 4
Private Const ST_SERIES_1 As Long = 5
Private Const ST_DATA_SHEET_2 As Long = 6
Private Const ST_SERIES_2 As Long = 7
Private Const ST_DATA_SHEET_3 As Long = 8
Private Const ST_SERIES_3 As Long = 9

Private Const ST_START_INDEX As Long = 10
Private Const ST_WINDOW_POINTS As Long = 11
Private Const ST_TOTAL_POINTS As Long = 12
Private Const ST_Y_ZOOM As Long = 13
Private Const ST_MARKER_SIZE As Long = 14
Private Const ST_AXIS_MIN As Long = 15
Private Const ST_AXIS_MAX As Long = 16
Private Const ST_SELECTION_SIGNATURE As Long = 17
Private Const ST_SELECTED_COUNT As Long = 18
Private Const ST_LAST_ACTION As Long = 19

Private Const STATE_META_GRID_COLUMNS As String = "U1"
Private Const STATE_META_GRID_ROWS As String = "V1"
Private Const STATE_META_GRID_COUNT As String = "W1"

Private Const STATE_SHEET_LIST_COLUMN As Long = 27     ' AA
Private Const STATE_SERIES_LIST_FIRST_COLUMN As Long = 30

Private Const NAME_SHEET_LIST As String = "_CG_SheetList"

'===============================================================================
' DATA-CACHE LAYOUT
'===============================================================================

Private Const CACHE_COLUMNS_PER_GRID As Long = 4       ' Date + three series

'===============================================================================
' SHAPE / CHART NAME PREFIXES
'===============================================================================

Private Const SHAPE_PREFIX As String = "cg_g"
Private Const CHART_NAME_SUFFIX As String = "_chart"
Private Const X_RANGE_SUFFIX As String = "_xrange"

'===============================================================================
' MODULE STATE
'===============================================================================

Private mBusyDepth As Long
Private mCurrentStage As String

'===============================================================================
' PUBLIC ENTRY POINTS
'===============================================================================

Public Function IsChartGridBusy() As Boolean
    IsChartGridBusy = (mBusyDepth > 0)
End Function

Public Sub ChartGrid()

    Dim specification As Variant
    Dim gridColumns As Long
    Dim gridRows As Long
    Dim gridCount As Long

    Dim wsGrid As Worksheet
    Dim wsState As Worksheet
    Dim wsData As Worksheet

    Dim gridId As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    On Error GoTo ErrorHandler

    specification = Application.InputBox( _
        Prompt:="Enter the dashboard layout as Columns*Rows." & vbCrLf & _
               "Example: 4*3", _
        Title:="Chart Grid", _
        Default:="4*3", _
        Type:=2)

    If VarType(specification) = vbBoolean Then
        If specification = False Then Exit Sub
    End If

    If Not ParseGridSpecification(CStr(specification), gridColumns, gridRows) Then
        MsgBox "Enter a valid layout such as 4*3.", vbExclamation
        Exit Sub
    End If

    gridCount = gridColumns * gridRows

    If gridCount > 100 Then
        MsgBox "The maximum supported number of grids is 100.", vbExclamation
        Exit Sub
    End If

    EnterBusyMode

    SaveAndApplyFastMode _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    mCurrentStage = "Preparing worksheets"

    Set wsGrid = GetOrCreateWorksheet(GRID_SHEET_NAME)
    Set wsState = GetOrCreateWorksheet(STATE_SHEET_NAME)
    Set wsData = GetOrCreateWorksheet(DATA_SHEET_NAME)

    PrepareGridWorksheet wsGrid
    PrepareStateWorksheet wsState, gridColumns, gridRows, gridCount
    PrepareDataWorksheet wsData, gridCount

    mCurrentStage = "Building data-sheet selector list"
    BuildGlobalSheetList wsState

    For gridId = 1 To gridCount

        Application.StatusBar = _
            "Building chart grid " & CStr(gridId) & " / " & CStr(gridCount)

        mCurrentStage = "Building grid " & CStr(gridId)

        BuildOneGrid _
            wsGrid:=wsGrid, _
            wsState:=wsState, _
            wsData:=wsData, _
            gridId:=gridId, _
            gridColumns:=gridColumns

    Next gridId

    wsState.Visible = xlSheetVeryHidden
    wsData.Visible = xlSheetVeryHidden

    wsGrid.Activate

CleanExit:

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    LeaveBusyMode
    Exit Sub

ErrorHandler:

    ShowDetailedError _
        procedureName:="ChartGrid", _
        stageName:=mCurrentStage, _
        errorNumber:=Err.Number, _
        errorDescription:=Err.Description

    Resume CleanExit

End Sub

Public Sub RefreshAllChartGrids()

    Dim wsGrid As Worksheet
    Dim wsState As Worksheet
    Dim wsData As Worksheet

    Dim gridCount As Long
    Dim gridId As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    On Error GoTo ErrorHandler

    Set wsGrid = GetWorksheet(GRID_SHEET_NAME)
    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    Set wsData = GetWorksheet(DATA_SHEET_NAME)

    If wsGrid Is Nothing Or wsState Is Nothing Or wsData Is Nothing Then
        MsgBox "Run ChartGrid first.", vbExclamation
        Exit Sub
    End If

    gridCount = GetLongOrDefault(wsState.Range(STATE_META_GRID_COUNT).value2, 0)
    If gridCount <= 0 Then Exit Sub

    EnterBusyMode

    SaveAndApplyFastMode _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    For gridId = 1 To gridCount

        Application.StatusBar = _
            "Refreshing chart grid " & CStr(gridId) & " / " & CStr(gridCount)

        mCurrentStage = "Reloading cache for grid " & CStr(gridId)

        SyncGridSelectorsToState wsGrid, wsState, gridId
        ReloadGridCache wsGrid, wsState, wsData, gridId
        RenderGridFromCache wsGrid, wsState, wsData, gridId

    Next gridId

CleanExit:

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    LeaveBusyMode
    Exit Sub

ErrorHandler:

    ShowDetailedError _
        procedureName:="RefreshAllChartGrids", _
        stageName:=mCurrentStage, _
        errorNumber:=Err.Number, _
        errorDescription:=Err.Description

    Resume CleanExit

End Sub

Public Sub CG_ButtonDispatch()

    Dim wsGrid As Worksheet
    Dim wsState As Worksheet
    Dim wsData As Worksheet
    Dim callerShape As Shape

    Dim gridId As Long
    Dim actionName As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    On Error GoTo ErrorHandler

    Set wsGrid = ActiveSheet

    If StrComp(wsGrid.Name, GRID_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Set callerShape = wsGrid.Shapes(CStr(Application.Caller))

    If Not ParseCallerMetadata(callerShape.AlternativeText, gridId, actionName) Then
        MsgBox "The control metadata is invalid.", vbExclamation
        Exit Sub
    End If

    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    Set wsData = GetWorksheet(DATA_SHEET_NAME)

    If wsState Is Nothing Or wsData Is Nothing Then Exit Sub

    EnterBusyMode

    SaveAndApplyFastMode _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    mCurrentStage = "Processing " & actionName & " for grid " & CStr(gridId)

    wsState.Cells(StateRow(gridId), ST_LAST_ACTION).value2 = actionName

    Select Case UCase$(actionName)

        Case "X_MINUS"
            ChangeGridXWindow wsState, gridId, -1
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "X_PLUS"
            ChangeGridXWindow wsState, gridId, 1
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "Y_MINUS"
            AdjustGridYZoom wsState, gridId, -10
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "Y_PLUS"
            AdjustGridYZoom wsState, gridId, 10
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "POINT_MINUS"
            AdjustGridMarkerSize wsState, gridId, -1
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "POINT_PLUS"
            AdjustGridMarkerSize wsState, gridId, 1
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "1W"
            ApplyLatestGridWindow wsState, gridId, POINTS_1W
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "1M"
            ApplyLatestGridWindow wsState, gridId, POINTS_1M
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "3M"
            ApplyLatestGridWindow wsState, gridId, POINTS_3M
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "6M"
            ApplyLatestGridWindow wsState, gridId, POINTS_6M
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "1Y"
            ApplyLatestGridWindow wsState, gridId, POINTS_1Y
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "3Y"
            ApplyLatestGridWindow wsState, gridId, POINTS_3Y
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "5Y"
            ApplyLatestGridWindow wsState, gridId, POINTS_5Y
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "MAX"
            ApplyLatestGridWindow wsState, gridId, 0
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "AUTO_Y"
            wsState.Cells(StateRow(gridId), ST_Y_ZOOM).value2 = DEFAULT_Y_ZOOM
            RenderGridFromCache wsGrid, wsState, wsData, gridId

        Case "UPDATE"
            CalculateGridSourceSheets wsGrid, gridId
            SyncGridSelectorsToState wsGrid, wsState, gridId
            ReloadGridCache wsGrid, wsState, wsData, gridId
            RenderGridFromCache wsGrid, wsState, wsData, gridId

    End Select

CleanExit:

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    LeaveBusyMode
    Exit Sub

ErrorHandler:

    ShowDetailedError _
        procedureName:="CG_ButtonDispatch", _
        stageName:=mCurrentStage, _
        errorNumber:=Err.Number, _
        errorDescription:=Err.Description

    Resume CleanExit

End Sub

Public Sub CG_XRangeChanged()

    Dim wsGrid As Worksheet
    Dim wsState As Worksheet
    Dim wsData As Worksheet
    Dim callerShape As Shape

    Dim gridId As Long
    Dim actionName As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    On Error GoTo ErrorHandler

    If IsChartGridBusy Then Exit Sub

    Set wsGrid = ActiveSheet

    If StrComp(wsGrid.Name, GRID_SHEET_NAME, vbTextCompare) <> 0 Then Exit Sub

    Set callerShape = wsGrid.Shapes(CStr(Application.Caller))

    If Not ParseCallerMetadata(callerShape.AlternativeText, gridId, actionName) Then Exit Sub

    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    Set wsData = GetWorksheet(DATA_SHEET_NAME)

    If wsState Is Nothing Or wsData Is Nothing Then Exit Sub

    EnterBusyMode

    SaveAndApplyFastMode _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    mCurrentStage = "Moving X range for grid " & CStr(gridId)

    RenderGridFromCache wsGrid, wsState, wsData, gridId

CleanExit:

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    LeaveBusyMode
    Exit Sub

ErrorHandler:

    ShowDetailedError _
        procedureName:="CG_XRangeChanged", _
        stageName:=mCurrentStage, _
        errorNumber:=Err.Number, _
        errorDescription:=Err.Description

    Resume CleanExit

End Sub

Public Sub ChartGrid_HandleChange(ByVal wsGrid As Worksheet, ByVal target As Range)

    Dim wsState As Worksheet
    Dim wsData As Worksheet

    Dim gridColumns As Long
    Dim gridRows As Long
    Dim gridCount As Long

    Dim changedGrids As Object
    Dim cellItem As Range

    Dim gridId As Long
    Dim gridSlot As Long
    Dim isDataSheetCell As Boolean
    Dim keyItem As Variant

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim oldStatusBar As Variant

    On Error GoTo ErrorHandler

    If IsChartGridBusy Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.CountLarge > 100 Then Exit Sub

    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    Set wsData = GetWorksheet(DATA_SHEET_NAME)

    If wsState Is Nothing Or wsData Is Nothing Then Exit Sub

    gridColumns = GetLongOrDefault(wsState.Range(STATE_META_GRID_COLUMNS).value2, 0)
    gridRows = GetLongOrDefault(wsState.Range(STATE_META_GRID_ROWS).value2, 0)
    gridCount = GetLongOrDefault(wsState.Range(STATE_META_GRID_COUNT).value2, 0)

    If gridColumns <= 0 Or gridRows <= 0 Or gridCount <= 0 Then Exit Sub

    Set changedGrids = CreateObject("Scripting.Dictionary")

    For Each cellItem In target.Cells

        If GetGridSelectorLocation( _
            targetCell:=cellItem, _
            gridColumns:=gridColumns, _
            gridRows:=gridRows, _
            gridId:=gridId, _
            slot:=gridSlot, _
            isDataSheetCell:=isDataSheetCell) Then

            If Not changedGrids.Exists(CStr(gridId)) Then
                changedGrids.Add CStr(gridId), True
            End If

        End If

    Next cellItem

    If changedGrids.count = 0 Then Exit Sub

    EnterBusyMode

    SaveAndApplyFastMode _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    For Each keyItem In changedGrids.keys

        gridId = CLng(keyItem)

        mCurrentStage = "Updating selectors for grid " & CStr(gridId)

        RebuildGridSeriesValidations wsGrid, wsState, gridId
        EnsureGridSelectionsValid wsGrid, wsState, gridId
        SyncGridSelectorsToState wsGrid, wsState, gridId
        ReloadGridCache wsGrid, wsState, wsData, gridId
        RenderGridFromCache wsGrid, wsState, wsData, gridId

    Next keyItem

CleanExit:

    RestoreApplicationState _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        oldCalculation, _
        oldStatusBar

    LeaveBusyMode
    Exit Sub

ErrorHandler:

    ShowDetailedError _
        procedureName:="ChartGrid_HandleChange", _
        stageName:=mCurrentStage, _
        errorNumber:=Err.Number, _
        errorDescription:=Err.Description

    Resume CleanExit

End Sub

'===============================================================================
' GRID CONSTRUCTION
'===============================================================================

Private Sub BuildOneGrid( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal gridId As Long, _
    ByVal gridColumns As Long)

    Dim gridRowIndex As Long
    Dim gridColumnIndex As Long

    Dim defaultSheet As String
    Dim defaultSeries As String

    gridRowIndex = ((gridId - 1) \ gridColumns) + 1
    gridColumnIndex = ((gridId - 1) Mod gridColumns) + 1

    defaultSheet = GetDefaultDataSheet()
    defaultSeries = GetDefaultSeriesForGrid(defaultSheet, gridId)

    InitializeGridState _
        wsState:=wsState, _
        gridId:=gridId, _
        gridRowIndex:=gridRowIndex, _
        gridColumnIndex:=gridColumnIndex, _
        defaultSheet:=defaultSheet, _
        defaultSeries:=defaultSeries

    WriteGridSelectorArea wsGrid, gridId
    WriteGridSelectorValues wsGrid, wsState, gridId

    BuildGridSeriesValidations wsGrid, wsState, gridId
    EnsureGridSelectionsValid wsGrid, wsState, gridId
    SyncGridSelectorsToState wsGrid, wsState, gridId

    CreateGridControls wsGrid, wsState, gridId

    ReloadGridCache wsGrid, wsState, wsData, gridId
    RenderGridFromCache wsGrid, wsState, wsData, gridId

End Sub

Private Sub InitializeGridState( _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal gridRowIndex As Long, _
    ByVal gridColumnIndex As Long, _
    ByVal defaultSheet As String, _
    ByVal defaultSeries As String)

    Dim rowNumber As Long

    rowNumber = StateRow(gridId)

    wsState.Cells(rowNumber, ST_GRID_ID).value2 = gridId
    wsState.Cells(rowNumber, ST_GRID_ROW).value2 = gridRowIndex
    wsState.Cells(rowNumber, ST_GRID_COL).value2 = gridColumnIndex

    wsState.Cells(rowNumber, ST_DATA_SHEET_1).value2 = defaultSheet
    wsState.Cells(rowNumber, ST_SERIES_1).value2 = defaultSeries

    wsState.Cells(rowNumber, ST_DATA_SHEET_2).value2 = defaultSheet
    wsState.Cells(rowNumber, ST_SERIES_2).value2 = NONE_SERIES_TEXT

    wsState.Cells(rowNumber, ST_DATA_SHEET_3).value2 = defaultSheet
    wsState.Cells(rowNumber, ST_SERIES_3).value2 = NONE_SERIES_TEXT

    wsState.Cells(rowNumber, ST_START_INDEX).value2 = 1
    wsState.Cells(rowNumber, ST_WINDOW_POINTS).value2 = DEFAULT_WINDOW_POINTS
    wsState.Cells(rowNumber, ST_TOTAL_POINTS).value2 = 0
    wsState.Cells(rowNumber, ST_Y_ZOOM).value2 = DEFAULT_Y_ZOOM
    wsState.Cells(rowNumber, ST_MARKER_SIZE).value2 = DEFAULT_MARKER_SIZE
    wsState.Cells(rowNumber, ST_AXIS_MIN).ClearContents
    wsState.Cells(rowNumber, ST_AXIS_MAX).ClearContents
    wsState.Cells(rowNumber, ST_SELECTION_SIGNATURE).value2 = vbNullString
    wsState.Cells(rowNumber, ST_SELECTED_COUNT).value2 = 0
    wsState.Cells(rowNumber, ST_LAST_ACTION).value2 = "BUILD"

End Sub

Private Sub WriteGridSelectorArea(ByVal wsGrid As Worksheet, ByVal gridId As Long)

    Dim baseRow As Long
    Dim baseColumn As Long
    Dim selectorRange As Range

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    wsGrid.Cells(baseRow + 1, baseColumn).value2 = "Data Sheet 1"
    wsGrid.Cells(baseRow + 2, baseColumn).value2 = "Series 1"

    wsGrid.Cells(baseRow + 3, baseColumn).value2 = "Data Sheet 2"
    wsGrid.Cells(baseRow + 4, baseColumn).value2 = "Series 2"

    wsGrid.Cells(baseRow + 5, baseColumn).value2 = "Data Sheet 3"
    wsGrid.Cells(baseRow + 6, baseColumn).value2 = "Series 3"

    Set selectorRange = wsGrid.Range( _
        wsGrid.Cells(baseRow + 1, baseColumn), _
        wsGrid.Cells(baseRow + 6, baseColumn + 1))

    With selectorRange
        .Font.Name = "Arial"
        .Font.Size = 8
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    wsGrid.Range( _
        wsGrid.Cells(baseRow + 1, baseColumn), _
        wsGrid.Cells(baseRow + 6, baseColumn)).Font.Bold = True

    ApplyThinBorders selectorRange, RGB(190, 190, 190)

End Sub

Private Sub WriteGridSelectorValues( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim stateRowNumber As Long
    Dim slot As Long

    stateRowNumber = StateRow(gridId)

    For slot = 1 To SERIES_SLOT_COUNT

        GridDataSheetCell(wsGrid, gridId, slot).value2 = _
            wsState.Cells(stateRowNumber, DataSheetStateColumn(slot)).value2

        GridSeriesCell(wsGrid, gridId, slot).value2 = _
            wsState.Cells(stateRowNumber, SeriesStateColumn(slot)).value2

    Next slot

End Sub

Private Sub CreateGridControls( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim baseRow As Long
    Dim baseColumn As Long

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    CreateLabeledThreePartControl _
        ws:=wsGrid, _
        target:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 8, baseColumn), _
            wsGrid.Cells(baseRow + 8, baseColumn + 1)), _
        gridId:=gridId, _
        controlKey:="XZOOM", _
        labelCaption:="X ZOOM", _
        minusAction:="X_MINUS", _
        plusAction:="X_PLUS"

    CreateLabeledThreePartControl _
        ws:=wsGrid, _
        target:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 9, baseColumn), _
            wsGrid.Cells(baseRow + 9, baseColumn + 1)), _
        gridId:=gridId, _
        controlKey:="YZOOM", _
        labelCaption:="Y ZOOM", _
        minusAction:="Y_MINUS", _
        plusAction:="Y_PLUS"

    CreateLabeledThreePartControl _
        ws:=wsGrid, _
        target:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 10, baseColumn), _
            wsGrid.Cells(baseRow + 10, baseColumn + 1)), _
        gridId:=gridId, _
        controlKey:="POINT", _
        labelCaption:="POINT", _
        minusAction:="POINT_MINUS", _
        plusAction:="POINT_PLUS"

    AddActionButton _
        ws:=wsGrid, _
        shapeName:=GridShapeName(gridId, "update"), _
        caption:="UPDATE", _
        gridId:=gridId, _
        actionName:="UPDATE", _
        leftPosition:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 11, baseColumn), _
            wsGrid.Cells(baseRow + 11, baseColumn + 1)).Left, _
        topPosition:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 11, baseColumn), _
            wsGrid.Cells(baseRow + 11, baseColumn + 1)).Top, _
        shapeWidth:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 11, baseColumn), _
            wsGrid.Cells(baseRow + 11, baseColumn + 1)).Width, _
        shapeHeight:=wsGrid.Range( _
            wsGrid.Cells(baseRow + 11, baseColumn), _
            wsGrid.Cells(baseRow + 11, baseColumn + 1)).Height, _
        isBold:=True

    CreateGridXRangeControl wsGrid, wsState, gridId
    CreateGridPeriodButtons wsGrid, gridId

End Sub

Private Sub CreateLabeledThreePartControl( _
    ByVal ws As Worksheet, _
    ByVal target As Range, _
    ByVal gridId As Long, _
    ByVal controlKey As String, _
    ByVal labelCaption As String, _
    ByVal minusAction As String, _
    ByVal plusAction As String)

    Dim gap As Double
    Dim labelWidth As Double
    Dim buttonWidth As Double
    Dim readoutWidth As Double
    Dim currentLeft As Double

    gap = 1
    labelWidth = target.Width * 0.34
    buttonWidth = target.Width * 0.14
    readoutWidth = target.Width - labelWidth - buttonWidth * 2 - gap * 3

    currentLeft = target.Left

    AddControlLabel _
        ws:=ws, _
        shapeName:=GridShapeName(gridId, LCase$(controlKey) & "_label"), _
        caption:=labelCaption, _
        leftPosition:=currentLeft, _
        topPosition:=target.Top, _
        shapeWidth:=labelWidth, _
        shapeHeight:=target.Height

    currentLeft = currentLeft + labelWidth + gap

    AddActionButton _
        ws:=ws, _
        shapeName:=GridShapeName(gridId, LCase$(controlKey) & "_minus"), _
        caption:="-", _
        gridId:=gridId, _
        actionName:=minusAction, _
        leftPosition:=currentLeft, _
        topPosition:=target.Top, _
        shapeWidth:=buttonWidth, _
        shapeHeight:=target.Height, _
        isBold:=False

    currentLeft = currentLeft + buttonWidth + gap

    AddReadoutBox _
        ws:=ws, _
        shapeName:=GridReadoutName(gridId, controlKey), _
        leftPosition:=currentLeft, _
        topPosition:=target.Top, _
        shapeWidth:=readoutWidth, _
        shapeHeight:=target.Height

    currentLeft = currentLeft + readoutWidth + gap

    AddActionButton _
        ws:=ws, _
        shapeName:=GridShapeName(gridId, LCase$(controlKey) & "_plus"), _
        caption:="+", _
        gridId:=gridId, _
        actionName:=plusAction, _
        leftPosition:=currentLeft, _
        topPosition:=target.Top, _
        shapeWidth:=buttonWidth, _
        shapeHeight:=target.Height, _
        isBold:=False

End Sub

Private Sub CreateGridXRangeControl( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim baseRow As Long
    Dim baseColumn As Long
    Dim targetRange As Range
    Dim controlShape As Shape

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    With wsGrid.Range( _
        wsGrid.Cells(baseRow + 19, baseColumn + 2), _
        wsGrid.Cells(baseRow + 19, baseColumn + 4))

        .Merge
        .value2 = "X RANGE"
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    Set targetRange = wsGrid.Range( _
        wsGrid.Cells(baseRow + 19, baseColumn + 5), _
        wsGrid.Cells(baseRow + 19, baseColumn + 14))

    Set controlShape = wsGrid.Shapes.AddFormControl( _
        Type:=xlScrollBar, _
        Left:=targetRange.Left, _
        Top:=targetRange.Top + 1, _
        Width:=targetRange.Width, _
        Height:=MaxDouble(10, targetRange.Height - 2))

    controlShape.Name = GridShapeName(gridId, "xrange")
    controlShape.Placement = xlMoveAndSize
    controlShape.OnAction = "CG_XRangeChanged"
    controlShape.AlternativeText = BuildCallerMetadata(gridId, "XRANGE")

    With controlShape.ControlFormat

        .LinkedCell = SheetCellReference( _
            wsState, _
            wsState.Cells(StateRow(gridId), ST_START_INDEX))

        .Min = 1
        .Max = 1
        .SmallChange = 1
        .LargeChange = 20
        .value = 1

    End With

    With wsGrid.Range( _
        wsGrid.Cells(baseRow + 19, baseColumn + 15), _
        wsGrid.Cells(baseRow + 19, baseColumn + 18))

        .Merge
        .Font.Name = "Arial"
        .Font.Size = 8
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub CreateGridPeriodButtons(ByVal wsGrid As Worksheet, ByVal gridId As Long)

    Dim baseRow As Long
    Dim baseColumn As Long
    Dim targetRange As Range

    Dim captions As Variant
    Dim actions As Variant

    Dim itemIndex As Long
    Dim gap As Double
    Dim buttonWidth As Double
    Dim currentLeft As Double

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    Set targetRange = wsGrid.Range( _
        wsGrid.Cells(baseRow + 20, baseColumn + 2), _
        wsGrid.Cells(baseRow + 20, baseColumn + 18))

    captions = Array("1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "MAX", "AUTO Y")
    actions = Array("1W", "1M", "3M", "6M", "1Y", "3Y", "5Y", "MAX", "AUTO_Y")

    gap = 1
    buttonWidth = _
        (targetRange.Width - gap * (UBound(captions) - LBound(captions))) / _
        (UBound(captions) - LBound(captions) + 1)

    currentLeft = targetRange.Left

    For itemIndex = LBound(captions) To UBound(captions)

        AddActionButton _
            ws:=wsGrid, _
            shapeName:=GridPeriodButtonName(gridId, CStr(actions(itemIndex))), _
            caption:=CStr(captions(itemIndex)), _
            gridId:=gridId, _
            actionName:=CStr(actions(itemIndex)), _
            leftPosition:=currentLeft, _
            topPosition:=targetRange.Top, _
            shapeWidth:=buttonWidth, _
            shapeHeight:=targetRange.Height, _
            isBold:=False

        currentLeft = currentLeft + buttonWidth + gap

    Next itemIndex

End Sub

Private Sub AddControlLabel( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal leftPosition As Double, _
    ByVal topPosition As Double, _
    ByVal shapeWidth As Double, _
    ByVal shapeHeight As Double)

    Dim labelShape As Shape

    Set labelShape = ws.Shapes.AddTextbox( _
        Orientation:=msoTextOrientationHorizontal, _
        Left:=leftPosition, _
        Top:=topPosition, _
        Width:=shapeWidth, _
        Height:=shapeHeight)

    labelShape.Name = shapeName
    labelShape.Placement = xlMoveAndSize

    With labelShape

        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.text = caption
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

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

Private Sub AddReadoutBox( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal leftPosition As Double, _
    ByVal topPosition As Double, _
    ByVal shapeWidth As Double, _
    ByVal shapeHeight As Double)

    Dim readoutShape As Shape

    Set readoutShape = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=leftPosition, _
        Top:=topPosition, _
        Width:=shapeWidth, _
        Height:=shapeHeight)

    readoutShape.Name = shapeName
    readoutShape.Placement = xlMoveAndSize

    With readoutShape

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75

        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.text = vbNullString
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

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

Private Sub AddActionButton( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String, _
    ByVal gridId As Long, _
    ByVal actionName As String, _
    ByVal leftPosition As Double, _
    ByVal topPosition As Double, _
    ByVal shapeWidth As Double, _
    ByVal shapeHeight As Double, _
    ByVal isBold As Boolean)

    Dim buttonShape As Shape

    Set buttonShape = ws.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=leftPosition, _
        Top:=topPosition, _
        Width:=shapeWidth, _
        Height:=shapeHeight)

    buttonShape.Name = shapeName
    buttonShape.Placement = xlMoveAndSize
    buttonShape.OnAction = "CG_ButtonDispatch"
    buttonShape.AlternativeText = BuildCallerMetadata(gridId, actionName)

    With buttonShape

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(242, 242, 242)

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(170, 170, 170)
        .Line.Weight = 0.75

        .Shadow.Visible = msoFalse

        .TextFrame2.TextRange.text = caption
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

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

'===============================================================================
' SELECTOR VALIDATION
'===============================================================================

Private Sub BuildGlobalSheetList(ByVal wsState As Worksheet)

    Dim ws As Worksheet
    Dim sheetNames() As Variant
    Dim itemCount As Long

    DeleteWorkbookName NAME_SHEET_LIST

    For Each ws In ThisWorkbook.Worksheets

        If IsSelectableDataSheet(ws.Name) Then
            itemCount = itemCount + 1
        End If

    Next ws

    If itemCount <= 0 Then Exit Sub

    ReDim sheetNames(1 To itemCount, 1 To 1)

    itemCount = 0

    For Each ws In ThisWorkbook.Worksheets

        If IsSelectableDataSheet(ws.Name) Then
            itemCount = itemCount + 1
            sheetNames(itemCount, 1) = ws.Name
        End If

    Next ws

    ClearUsedColumn wsState, STATE_SHEET_LIST_COLUMN
    wsState.Cells(2, STATE_SHEET_LIST_COLUMN).Resize(itemCount, 1).value2 = sheetNames

    ThisWorkbook.names.Add _
        Name:=NAME_SHEET_LIST, _
        RefersTo:="=" & SheetRangeFormula( _
            wsState, _
            wsState.Cells(2, STATE_SHEET_LIST_COLUMN).Resize(itemCount, 1))

End Sub

Private Sub BuildGridSeriesValidations( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim slot As Long

    For slot = 1 To SERIES_SLOT_COUNT
        BuildSeriesValidationForSlot wsGrid, wsState, gridId, slot
    Next slot

End Sub

Private Sub RebuildGridSeriesValidations( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    BuildGridSeriesValidations wsGrid, wsState, gridId

End Sub

Private Sub BuildSeriesValidationForSlot( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal slot As Long)

    Dim sourceSheet As Worksheet
    Dim dataSheetName As String

    Dim listColumn As Long
    Dim listName As String

    Dim lastColumn As Long
    Dim headers As Variant
    Dim outputData() As Variant

    Dim headerIndex As Long
    Dim itemCount As Long
    Dim outputIndex As Long
    Dim headerText As String

    dataSheetName = SafeCellText(GridDataSheetCell(wsGrid, gridId, slot))
    Set sourceSheet = GetWorksheet(dataSheetName)

    listColumn = SeriesListColumn(gridId, slot)
    listName = GridSeriesListName(gridId, slot)

    DeleteWorkbookName listName
    ClearUsedColumn wsState, listColumn

    itemCount = 1

    If Not sourceSheet Is Nothing Then

        lastColumn = sourceSheet.Cells(1, sourceSheet.Columns.count).End(xlToLeft).Column

        If lastColumn >= 2 Then

            headers = sourceSheet.Range( _
                sourceSheet.Cells(1, 2), _
                sourceSheet.Cells(1, lastColumn)).value2

            For headerIndex = 1 To UBound(headers, 2)

                If Not IsError(headers(1, headerIndex)) Then
                    headerText = Trim$(CStr(headers(1, headerIndex)))
                    If Len(headerText) > 0 Then itemCount = itemCount + 1
                End If

            Next headerIndex

        End If

    End If

    ReDim outputData(1 To itemCount, 1 To 1)

    outputData(1, 1) = NONE_SERIES_TEXT
    outputIndex = 1

    If Not sourceSheet Is Nothing Then

        If lastColumn >= 2 Then

            For headerIndex = 1 To UBound(headers, 2)

                If Not IsError(headers(1, headerIndex)) Then

                    headerText = Trim$(CStr(headers(1, headerIndex)))

                    If Len(headerText) > 0 Then
                        outputIndex = outputIndex + 1
                        outputData(outputIndex, 1) = headerText
                    End If

                End If

            Next headerIndex

        End If

    End If

    wsState.Cells(2, listColumn).Resize(itemCount, 1).value2 = outputData

    ThisWorkbook.names.Add _
        Name:=listName, _
        RefersTo:="=" & SheetRangeFormula( _
            wsState, _
            wsState.Cells(2, listColumn).Resize(itemCount, 1))

    SetListValidation GridSeriesCell(wsGrid, gridId, slot), "=" & listName
    SetListValidation GridDataSheetCell(wsGrid, gridId, slot), "=" & NAME_SHEET_LIST

End Sub

Private Sub EnsureGridSelectionsValid( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim fallbackSheet As String
    Dim selectedSheet As String
    Dim selectedSeries As String

    Dim slot As Long
    Dim earlierSlot As Long
    Dim duplicateFound As Boolean

    fallbackSheet = GetDefaultDataSheet()

    For slot = 1 To SERIES_SLOT_COUNT

        selectedSheet = SafeCellText(GridDataSheetCell(wsGrid, gridId, slot))

        If Not IsSelectableDataSheet(selectedSheet) Then
            selectedSheet = fallbackSheet
            GridDataSheetCell(wsGrid, gridId, slot).value2 = selectedSheet
            BuildSeriesValidationForSlot wsGrid, wsState, gridId, slot
        End If

        selectedSeries = SafeCellText(GridSeriesCell(wsGrid, gridId, slot))

        If slot = 1 Then

            If StrComp(selectedSeries, NONE_SERIES_TEXT, vbTextCompare) = 0 _
                Or Len(selectedSeries) = 0 _
                Or Not IsSeriesAvailable(selectedSheet, selectedSeries) Then

                GridSeriesCell(wsGrid, gridId, slot).value2 = _
                    FirstSeriesName(selectedSheet)

            End If

        Else

            If Len(selectedSeries) = 0 _
                Or StrComp(selectedSeries, NONE_SERIES_TEXT, vbTextCompare) = 0 _
                Or Not IsSeriesAvailable(selectedSheet, selectedSeries) Then

                GridSeriesCell(wsGrid, gridId, slot).value2 = NONE_SERIES_TEXT

            End If

        End If

    Next slot

    For slot = 2 To SERIES_SLOT_COUNT

        duplicateFound = False

        If StrComp( _
            SafeCellText(GridSeriesCell(wsGrid, gridId, slot)), _
            NONE_SERIES_TEXT, _
            vbTextCompare) <> 0 Then

            For earlierSlot = 1 To slot - 1

                If StrComp( _
                    SafeCellText(GridDataSheetCell(wsGrid, gridId, slot)), _
                    SafeCellText(GridDataSheetCell(wsGrid, gridId, earlierSlot)), _
                    vbTextCompare) = 0 _
                    And StrComp( _
                        SafeCellText(GridSeriesCell(wsGrid, gridId, slot)), _
                        SafeCellText(GridSeriesCell(wsGrid, gridId, earlierSlot)), _
                        vbTextCompare) = 0 Then

                    duplicateFound = True
                    Exit For

                End If

            Next earlierSlot

        End If

        If duplicateFound Then
            GridSeriesCell(wsGrid, gridId, slot).value2 = NONE_SERIES_TEXT
        End If

    Next slot

End Sub

Private Sub SetListValidation(ByVal target As Range, ByVal formulaText As String)

    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0

    target.Validation.Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:=formulaText

    With target.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With

End Sub

'===============================================================================
' GRID CACHE
'===============================================================================

Private Sub ReloadGridCache( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal gridId As Long)

    Dim selectedSheets(1 To SERIES_SLOT_COUNT) As String
    Dim selectedSeries(1 To SERIES_SLOT_COUNT) As String
    Dim displayNames(1 To SERIES_SLOT_COUNT) As String

    Dim selectedCount As Long
    Dim dataCount As Long

    Dim xAll() As Variant
    Dim yAll() As Variant

    Dim errorMessage As String
    Dim selectionSignature As String

    Dim stateRowNumber As Long
    Dim oldTotalPoints As Long
    Dim windowPoints As Long
    Dim startIndex As Long

    stateRowNumber = StateRow(gridId)

    selectedCount = ReadSelectedSeriesFromGrid( _
        wsGrid:=wsGrid, _
        gridId:=gridId, _
        selectedSheets:=selectedSheets, _
        selectedSeries:=selectedSeries)

    If selectedCount <= 0 Then

        ClearGridCache wsData, wsState, gridId
        wsState.Cells(stateRowNumber, ST_SELECTED_COUNT).value2 = 0
        Exit Sub

    End If

    BuildDisplayNames _
        selectedSheets:=selectedSheets, _
        selectedSeries:=selectedSeries, _
        displayNames:=displayNames, _
        selectedCount:=selectedCount

    selectionSignature = BuildSelectionSignature( _
        selectedSheets, _
        selectedSeries, _
        selectedCount)

    If Not LoadAlignedDashboardData( _
        selectedSheets:=selectedSheets, _
        selectedSeries:=selectedSeries, _
        selectedCount:=selectedCount, _
        xAll:=xAll, _
        yAll:=yAll, _
        dataCount:=dataCount, _
        errorMessage:=errorMessage) Then

        ClearGridCache wsData, wsState, gridId
        wsState.Cells(stateRowNumber, ST_SELECTED_COUNT).value2 = 0

        If Len(errorMessage) > 0 Then
            MsgBox "Grid " & CStr(gridId) & ": " & errorMessage, vbExclamation
        End If

        Exit Sub

    End If

    oldTotalPoints = GetLongOrDefault( _
        wsState.Cells(stateRowNumber, ST_TOTAL_POINTS).value2, _
        0)

    WriteGridCache _
        wsData:=wsData, _
        gridId:=gridId, _
        displayNames:=displayNames, _
        selectedCount:=selectedCount, _
        xAll:=xAll, _
        yAll:=yAll, _
        dataCount:=dataCount, _
        previousDataCount:=oldTotalPoints

    windowPoints = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(stateRowNumber, ST_WINDOW_POINTS).value2, _
            DEFAULT_WINDOW_POINTS), _
        WorksheetFunction.Min(MIN_WINDOW_POINTS, dataCount), _
        dataCount)

    startIndex = MaxLong(1, dataCount - windowPoints + 1)

    wsState.Cells(stateRowNumber, ST_START_INDEX).value2 = startIndex
    wsState.Cells(stateRowNumber, ST_WINDOW_POINTS).value2 = windowPoints
    wsState.Cells(stateRowNumber, ST_TOTAL_POINTS).value2 = dataCount
    wsState.Cells(stateRowNumber, ST_Y_ZOOM).value2 = DEFAULT_Y_ZOOM
    wsState.Cells(stateRowNumber, ST_SELECTION_SIGNATURE).value2 = selectionSignature
    wsState.Cells(stateRowNumber, ST_SELECTED_COUNT).value2 = selectedCount

End Sub

Private Function ReadSelectedSeriesFromGrid( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByRef selectedSheets() As String, _
    ByRef selectedSeries() As String) As Long

    Dim slot As Long
    Dim existingIndex As Long
    Dim selectedCount As Long

    Dim candidateSheet As String
    Dim candidateSeries As String
    Dim duplicateFound As Boolean

    For slot = 1 To SERIES_SLOT_COUNT

        candidateSheet = SafeCellText(GridDataSheetCell(wsGrid, gridId, slot))
        candidateSeries = SafeCellText(GridSeriesCell(wsGrid, gridId, slot))

        If Len(candidateSheet) > 0 _
            And Len(candidateSeries) > 0 _
            And StrComp(candidateSeries, NONE_SERIES_TEXT, vbTextCompare) <> 0 Then

            duplicateFound = False

            For existingIndex = 1 To selectedCount

                If StrComp(selectedSheets(existingIndex), candidateSheet, vbTextCompare) = 0 _
                    And StrComp(selectedSeries(existingIndex), candidateSeries, vbTextCompare) = 0 Then

                    duplicateFound = True
                    Exit For

                End If

            Next existingIndex

            If Not duplicateFound Then

                selectedCount = selectedCount + 1
                selectedSheets(selectedCount) = candidateSheet
                selectedSeries(selectedCount) = candidateSeries

            End If

        End If

    Next slot

    ReadSelectedSeriesFromGrid = selectedCount

End Function

Private Sub BuildDisplayNames( _
    ByRef selectedSheets() As String, _
    ByRef selectedSeries() As String, _
    ByRef displayNames() As String, _
    ByVal selectedCount As Long)

    Dim seriesIndex As Long
    Dim comparisonIndex As Long
    Dim duplicateCount As Long

    For seriesIndex = 1 To selectedCount

        duplicateCount = 0

        For comparisonIndex = 1 To selectedCount

            If StrComp( _
                selectedSeries(seriesIndex), _
                selectedSeries(comparisonIndex), _
                vbTextCompare) = 0 Then

                duplicateCount = duplicateCount + 1

            End If

        Next comparisonIndex

        If duplicateCount > 1 Then

            displayNames(seriesIndex) = _
                selectedSeries(seriesIndex) & _
                " [" & selectedSheets(seriesIndex) & "]"

        Else

            displayNames(seriesIndex) = selectedSeries(seriesIndex)

        End If

    Next seriesIndex

End Sub

Private Function LoadAlignedDashboardData( _
    ByRef selectedSheets() As String, _
    ByRef selectedSeries() As String, _
    ByVal selectedCount As Long, _
    ByRef xAll() As Variant, _
    ByRef yAll() As Variant, _
    ByRef dataCount As Long, _
    ByRef errorMessage As String) As Boolean

    Dim datesBySeries(1 To SERIES_SLOT_COUNT) As Variant
    Dim valuesBySeries(1 To SERIES_SLOT_COUNT) As Variant
    Dim counts(1 To SERIES_SLOT_COUNT) As Long

    Dim uniqueDates As Object
    Dim dateIndexes As Object

    Dim currentDates As Variant
    Dim currentValues As Variant
    Dim dateItems As Variant

    Dim seriesIndex As Long
    Dim pointIndex As Long
    Dim alignedIndex As Long

    Dim dateKeyText As String

    Set uniqueDates = CreateObject("Scripting.Dictionary")
    Set dateIndexes = CreateObject("Scripting.Dictionary")

    For seriesIndex = 1 To selectedCount

        If Not LoadSingleSeriesData( _
            dataSheetName:=selectedSheets(seriesIndex), _
            seriesName:=selectedSeries(seriesIndex), _
            datesOut:=datesBySeries(seriesIndex), _
            valuesOut:=valuesBySeries(seriesIndex), _
            itemCount:=counts(seriesIndex), _
            errorMessage:=errorMessage) Then

            Exit Function

        End If

        currentDates = datesBySeries(seriesIndex)

        For pointIndex = 1 To counts(seriesIndex)

            dateKeyText = DateKey(currentDates(pointIndex))

            If Not uniqueDates.Exists(dateKeyText) Then
                uniqueDates.Add dateKeyText, CDbl(currentDates(pointIndex))
            End If

        Next pointIndex

    Next seriesIndex

    dataCount = uniqueDates.count

    If dataCount <= 0 Then
        errorMessage = "No valid dates were found."
        Exit Function
    End If

    ReDim xAll(1 To dataCount)
    ReDim yAll(1 To dataCount, 1 To SERIES_SLOT_COUNT)

    dateItems = uniqueDates.items

    For pointIndex = 0 To dataCount - 1
        xAll(pointIndex + 1) = CDbl(dateItems(pointIndex))
    Next pointIndex

    QuickSortDoubles xAll, 1, dataCount

    For pointIndex = 1 To dataCount

        dateIndexes.Add DateKey(xAll(pointIndex)), pointIndex

        For seriesIndex = 1 To SERIES_SLOT_COUNT
            yAll(pointIndex, seriesIndex) = CVErr(xlErrNA)
        Next seriesIndex

    Next pointIndex

    For seriesIndex = 1 To selectedCount

        currentDates = datesBySeries(seriesIndex)
        currentValues = valuesBySeries(seriesIndex)

        For pointIndex = 1 To counts(seriesIndex)

            dateKeyText = DateKey(currentDates(pointIndex))
            alignedIndex = CLng(dateIndexes(dateKeyText))
            yAll(alignedIndex, seriesIndex) = currentValues(pointIndex)

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

    Dim ws As Worksheet
    Dim seriesColumn As Long

    Dim lastDateRow As Long
    Dim lastValueRow As Long
    Dim lastRow As Long

    Dim dateData As Variant
    Dim valueData As Variant

    Dim dateArray() As Double
    Dim valueArray() As Double

    Dim sourceIndex As Long
    Dim normalizedDate As Double
    Dim normalizedValue As Double

    Set ws = GetWorksheet(dataSheetName)

    If ws Is Nothing Then
        errorMessage = "Data sheet not found: " & dataSheetName
        Exit Function
    End If

    seriesColumn = FindSeriesColumn(ws, seriesName)

    If seriesColumn = 0 Then
        errorMessage = "Series not found: " & dataSheetName & " / " & seriesName
        Exit Function
    End If

    lastDateRow = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    lastValueRow = ws.Cells(ws.rows.count, seriesColumn).End(xlUp).Row
    lastRow = MaxLong(lastDateRow, lastValueRow)

    If lastRow < 2 Then
        errorMessage = "No data found: " & dataSheetName & " / " & seriesName
        Exit Function
    End If

    dateData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).value2
    valueData = ws.Range(ws.Cells(2, seriesColumn), ws.Cells(lastRow, seriesColumn)).value2

    ReDim dateArray(1 To UBound(dateData, 1))
    ReDim valueArray(1 To UBound(valueData, 1))

    itemCount = 0

    For sourceIndex = 1 To UBound(dateData, 1)

        If TryGetDateSerial(dateData(sourceIndex, 1), normalizedDate) _
            And TryGetDouble(valueData(sourceIndex, 1), normalizedValue) Then

            itemCount = itemCount + 1
            dateArray(itemCount) = normalizedDate
            valueArray(itemCount) = normalizedValue

        End If

    Next sourceIndex

    If itemCount <= 0 Then
        errorMessage = "No valid observations found: " & dataSheetName & " / " & seriesName
        Exit Function
    End If

    ReDim Preserve dateArray(1 To itemCount)
    ReDim Preserve valueArray(1 To itemCount)

    datesOut = dateArray
    valuesOut = valueArray

    LoadSingleSeriesData = True

End Function

Private Sub WriteGridCache( _
    ByVal wsData As Worksheet, _
    ByVal gridId As Long, _
    ByRef displayNames() As String, _
    ByVal selectedCount As Long, _
    ByRef xAll() As Variant, _
    ByRef yAll() As Variant, _
    ByVal dataCount As Long, _
    ByVal previousDataCount As Long)

    Dim baseColumn As Long
    Dim outputData() As Variant

    Dim pointIndex As Long
    Dim seriesIndex As Long
    Dim clearRows As Long

    baseColumn = DataBaseColumn(gridId)
    clearRows = MaxLong(previousDataCount + 1, dataCount + 1)

    wsData.Cells(1, baseColumn).Resize(clearRows, CACHE_COLUMNS_PER_GRID).ClearContents

    ReDim outputData(1 To dataCount + 1, 1 To CACHE_COLUMNS_PER_GRID)

    outputData(1, 1) = "Date"

    For seriesIndex = 1 To selectedCount
        outputData(1, seriesIndex + 1) = displayNames(seriesIndex)
    Next seriesIndex

    For pointIndex = 1 To dataCount

        outputData(pointIndex + 1, 1) = CDbl(xAll(pointIndex))

        For seriesIndex = 1 To selectedCount
            outputData(pointIndex + 1, seriesIndex + 1) = yAll(pointIndex, seriesIndex)
        Next seriesIndex

    Next pointIndex

    wsData.Cells(1, baseColumn).Resize(dataCount + 1, CACHE_COLUMNS_PER_GRID).value2 = outputData
    wsData.Cells(2, baseColumn).Resize(dataCount, 1).numberFormat = "yyyy/mm/dd"

End Sub

Private Sub ClearGridCache( _
    ByVal wsData As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim stateRowNumber As Long
    Dim oldTotalPoints As Long
    Dim baseColumn As Long

    stateRowNumber = StateRow(gridId)
    oldTotalPoints = GetLongOrDefault( _
        wsState.Cells(stateRowNumber, ST_TOTAL_POINTS).value2, _
        0)

    baseColumn = DataBaseColumn(gridId)

    wsData.Cells(1, baseColumn).Resize(MaxLong(1, oldTotalPoints + 1), CACHE_COLUMNS_PER_GRID).ClearContents

    wsState.Cells(stateRowNumber, ST_START_INDEX).value2 = 1
    wsState.Cells(stateRowNumber, ST_TOTAL_POINTS).value2 = 0
    wsState.Cells(stateRowNumber, ST_AXIS_MIN).ClearContents
    wsState.Cells(stateRowNumber, ST_AXIS_MAX).ClearContents

End Sub

'===============================================================================
' CHART RENDERING
'===============================================================================

Private Sub RenderGridFromCache( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal wsData As Worksheet, _
    ByVal gridId As Long)

    Dim stateRowNumber As Long
    Dim selectedCount As Long
    Dim totalPoints As Long
    Dim windowPoints As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim maximumStartIndex As Long

    Dim yZoom As Long
    Dim markerSize As Long

    Dim baseDataColumn As Long
    Dim xRange As Range
    Dim yRange As Range
    Dim nameCell As Range

    Dim chartObject As chartObject
    Dim chartItem As chart
    Dim chartSeries As Series

    Dim seriesIndex As Long
    Dim displayName As String

    Dim firstDate As Double
    Dim lastDate As Double

    Dim visibleMinimum As Double
    Dim visibleMaximum As Double
    Dim rawAxisMinimum As Double
    Dim rawAxisMaximum As Double
    Dim axisMinimum As Double
    Dim axisMaximum As Double
    Dim yMajorUnit As Double
    Dim yMinorUnit As Double
    Dim yNumberFormat As String

    Dim xMajorUnit As Double
    Dim xMinorUnit As Double
    Dim xNumberFormat As String
    Dim xRotation As Long
    Dim plotWidth As Double

    stateRowNumber = StateRow(gridId)

    selectedCount = GetLongOrDefault( _
        wsState.Cells(stateRowNumber, ST_SELECTED_COUNT).value2, _
        0)

    totalPoints = GetLongOrDefault( _
        wsState.Cells(stateRowNumber, ST_TOTAL_POINTS).value2, _
        0)

    Set chartObject = GetOrCreateGridChart(wsGrid, gridId)
    Set chartItem = chartObject.chart

    InitializeGridChart chartItem

    If selectedCount <= 0 Or totalPoints <= 0 Then

        ResetGridSeries chartItem
        DeleteGridLegendShapes wsGrid, gridId
        SuppressGridChartTitle chartItem
        UpdateGridControlState wsGrid, wsState, gridId
        Exit Sub

    End If

    windowPoints = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(stateRowNumber, ST_WINDOW_POINTS).value2, _
            DEFAULT_WINDOW_POINTS), _
        WorksheetFunction.Min(MIN_WINDOW_POINTS, totalPoints), _
        totalPoints)

    maximumStartIndex = MaxLong(1, totalPoints - windowPoints + 1)

    startIndex = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(stateRowNumber, ST_START_INDEX).value2, _
            maximumStartIndex), _
        1, _
        maximumStartIndex)

    endIndex = startIndex + windowPoints - 1

    yZoom = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(stateRowNumber, ST_Y_ZOOM).value2, _
            DEFAULT_Y_ZOOM), _
        Y_ZOOM_MIN, _
        Y_ZOOM_MAX)

    markerSize = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(stateRowNumber, ST_MARKER_SIZE).value2, _
            DEFAULT_MARKER_SIZE), _
        MARKER_SIZE_MIN, _
        MARKER_SIZE_MAX)

    wsState.Cells(stateRowNumber, ST_START_INDEX).value2 = startIndex
    wsState.Cells(stateRowNumber, ST_WINDOW_POINTS).value2 = windowPoints
    wsState.Cells(stateRowNumber, ST_Y_ZOOM).value2 = yZoom
    wsState.Cells(stateRowNumber, ST_MARKER_SIZE).value2 = markerSize

    baseDataColumn = DataBaseColumn(gridId)

    Set xRange = wsData.Cells(startIndex + 1, baseDataColumn).Resize(windowPoints, 1)

    ResetGridSeries chartItem

    For seriesIndex = 1 To selectedCount

        Set nameCell = wsData.Cells(1, baseDataColumn + seriesIndex)
        displayName = Trim$(CStr(nameCell.value2))
        Set yRange = wsData.Cells(startIndex + 1, baseDataColumn + seriesIndex).Resize(windowPoints, 1)

        Set chartSeries = chartItem.SeriesCollection.NewSeries

        With chartSeries
            ' Force the name to a single helper cell. A bare name such as
            ' yield_1y can otherwise be resolved as a workbook defined name,
            ' causing dates from that range to appear as legend entries.
            .Name = SeriesNameCellFormula(nameCell)
            .xValues = xRange
            .values = yRange
            .AxisGroup = xlPrimary
            .Smooth = False
            On Error Resume Next
            .HasDataLabels = False
            On Error GoTo 0
        End With

        FormatMainSeries chartSeries, GetSeriesColor(seriesIndex), markerSize

    Next seriesIndex

    SuppressGridChartTitle chartItem

    FormatChartArea chartItem

    firstDate = CDbl(xRange.Cells(1, 1).value2)
    lastDate = CDbl(xRange.Cells(windowPoints, 1).value2)

    If lastDate <= firstDate Then lastDate = firstDate + 1

    GetVisibleCacheMinMax _
        wsData:=wsData, _
        baseDataColumn:=baseDataColumn, _
        selectedCount:=selectedCount, _
        startIndex:=startIndex, _
        windowPoints:=windowPoints, _
        minimumValue:=visibleMinimum, _
        maximumValue:=visibleMaximum

    CalculateYAxisRange _
        dataMinimum:=visibleMinimum, _
        dataMaximum:=visibleMaximum, _
        yZoomPercent:=yZoom, _
        axisMinimum:=rawAxisMinimum, _
        axisMaximum:=rawAxisMaximum

    GetNiceYAxisScale _
        rawMinimum:=rawAxisMinimum, _
        rawMaximum:=rawAxisMaximum, _
        niceMinimum:=axisMinimum, _
        niceMaximum:=axisMaximum, _
        majorUnit:=yMajorUnit, _
        minorUnit:=yMinorUnit

    yNumberFormat = NumberFormatFromStep(yMajorUnit)

    ConfigureYAxis _
        chartItem:=chartItem, _
        axisMinimum:=axisMinimum, _
        axisMaximum:=axisMaximum, _
        majorUnit:=yMajorUnit, _
        minorUnit:=yMinorUnit, _
        numberFormat:=yNumberFormat

    plotWidth = chartItem.PlotArea.InsideWidth
    If plotWidth <= 0 Then plotWidth = chartObject.Width - 80

    GetAdaptiveDateAxisSettings _
        firstDate:=firstDate, _
        lastDate:=lastDate, _
        plotWidth:=plotWidth, _
        majorUnit:=xMajorUnit, _
        numberFormat:=xNumberFormat, _
        labelRotation:=xRotation

    xMinorUnit = MaxDouble(1, xMajorUnit / 2)

    ConfigureXAxis _
        chartItem:=chartItem, _
        firstDate:=firstDate, _
        lastDate:=lastDate, _
        majorUnit:=xMajorUnit, _
        minorUnit:=xMinorUnit, _
        numberFormat:=xNumberFormat, _
        labelRotation:=xRotation

    DisableNativeGridLegend chartItem

    For seriesIndex = 1 To selectedCount

        Set yRange = wsData.Cells(startIndex + 1, baseDataColumn + seriesIndex).Resize(windowPoints, 1)

        HighlightLatestPoint _
            chartSeries:=chartItem.SeriesCollection(seriesIndex), _
            visibleValues:=yRange, _
            seriesIndex:=seriesIndex, _
            markerSize:=markerSize, _
            numberFormat:=yNumberFormat, _
            axisMinimum:=axisMinimum, _
            axisMaximum:=axisMaximum

    Next seriesIndex

    DrawGridCustomLegend _
        wsGrid:=wsGrid, _
        chartObject:=chartObject, _
        chartItem:=chartItem, _
        wsData:=wsData, _
        baseDataColumn:=baseDataColumn, _
        selectedCount:=selectedCount, _
        gridId:=gridId

    wsState.Cells(stateRowNumber, ST_AXIS_MIN).value2 = axisMinimum
    wsState.Cells(stateRowNumber, ST_AXIS_MAX).value2 = axisMaximum

    SuppressGridChartTitle chartItem

    On Error Resume Next
    chartItem.Refresh
    On Error GoTo 0

    SuppressGridChartTitle chartItem

    UpdateGridControlState wsGrid, wsState, gridId

End Sub

Private Function GetOrCreateGridChart( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long) As chartObject

    Dim chartObject As chartObject
    Dim targetRange As Range
    Dim chartName As String

    chartName = GridChartName(gridId)
    Set targetRange = GridChartRange(wsGrid, gridId)

    On Error Resume Next
    Set chartObject = wsGrid.ChartObjects(chartName)
    On Error GoTo 0

    If chartObject Is Nothing Then

        Set chartObject = wsGrid.ChartObjects.Add( _
            Left:=targetRange.Left, _
            Top:=targetRange.Top, _
            Width:=targetRange.Width, _
            Height:=targetRange.Height)

        chartObject.Name = chartName

    End If

    With chartObject
        .Left = targetRange.Left
        .Top = targetRange.Top
        .Width = targetRange.Width
        .Height = targetRange.Height
        .Placement = xlMoveAndSize
    End With

    Set GetOrCreateGridChart = chartObject

End Function

Private Sub InitializeGridChart(ByVal chartItem As chart)

    With chartItem
        .ChartType = xlXYScatterLinesNoMarkers
        .PlotVisibleOnly = False
        .DisplayBlanksAs = xlNotPlotted
        .HasLegend = False
        .HasTitle = False
    End With

    SuppressGridChartTitle chartItem

End Sub

Private Sub SuppressGridChartTitle(ByVal chartItem As chart)

    On Error Resume Next

    If chartItem.HasTitle Then
        chartItem.ChartTitle.Delete
    End If

    chartItem.HasTitle = False

    Err.Clear
    On Error GoTo 0

End Sub

Private Sub ResetGridSeries(ByVal chartItem As chart)

    Dim seriesIndex As Long

    On Error Resume Next
    chartItem.HasLegend = False
    On Error GoTo 0

    For seriesIndex = chartItem.SeriesCollection.count To 1 Step -1
        chartItem.SeriesCollection(seriesIndex).Delete
    Next seriesIndex

End Sub

Private Sub FormatMainSeries( _
    ByVal chartSeries As Series, _
    ByVal seriesColor As Long, _
    ByVal markerSize As Long)

    Dim actualMarkerSize As Long

    With chartSeries.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = seriesColor
        .Weight = 1.25
    End With

    If markerSize <= 0 Then

        chartSeries.MarkerStyle = xlMarkerStyleNone

    Else

        actualMarkerSize = ClampLong(markerSize, 2, MARKER_SIZE_MAX)

        With chartSeries
            .MarkerStyle = xlMarkerStyleCircle
            .markerSize = actualMarkerSize
            .MarkerForegroundColor = seriesColor
            .MarkerBackgroundColor = RGB(255, 255, 255)
        End With

    End If

End Sub

Private Sub FormatChartArea(ByVal chartItem As chart)

    With chartItem.chartArea.Format
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(190, 190, 190)
        .Line.Weight = 0.5
    End With

    With chartItem.PlotArea.Format
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
    End With

End Sub

Private Sub DisableNativeGridLegend(ByVal chartItem As chart)

    On Error Resume Next

    chartItem.HasLegend = False
    chartItem.Legend.Delete

    Err.Clear
    On Error GoTo 0

End Sub

Private Function SeriesNameCellFormula(ByVal nameCell As Range) As String

    ' Use a one-cell formula reference instead of assigning a bare string.
    ' This prevents names such as yield_1y from being interpreted as a
    ' workbook-level defined name or a multi-cell range.
    SeriesNameCellFormula = _
        "='" & Replace(nameCell.Worksheet.Name, "'", "''") & "'!" & _
        nameCell.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1)

End Function

Private Sub DrawGridCustomLegend( _
    ByVal wsGrid As Worksheet, _
    ByVal chartObject As chartObject, _
    ByVal chartItem As chart, _
    ByVal wsData As Worksheet, _
    ByVal baseDataColumn As Long, _
    ByVal selectedCount As Long, _
    ByVal gridId As Long)

    Dim backgroundShape As Shape
    Dim lineShape As Shape
    Dim textShape As Shape

    Dim seriesIndex As Long
    Dim maximumNameLength As Long

    Dim displayName As String
    Dim legendPrefix As String

    Dim legendLeft As Double
    Dim legendTop As Double
    Dim legendWidth As Double
    Dim legendHeight As Double
    Dim rowTop As Double
    Dim lineCenterY As Double

    DeleteGridLegendShapes wsGrid, gridId
    DisableNativeGridLegend chartItem

    If selectedCount <= 0 Then Exit Sub

    legendPrefix = GridShapeName(gridId, "legend_")

    For seriesIndex = 1 To selectedCount

        displayName = Trim$(CStr( _
            wsData.Cells(1, baseDataColumn + seriesIndex).value2))

        maximumNameLength = MaxLong( _
            maximumNameLength, _
            Len(displayName))

    Next seriesIndex

    On Error Resume Next

    legendLeft = chartObject.Left + chartItem.PlotArea.InsideLeft + 6
    legendTop = chartObject.Top + chartItem.PlotArea.InsideTop + 5

    If Err.Number <> 0 Then
        Err.Clear
        legendLeft = chartObject.Left + 10
        legendTop = chartObject.Top + 8
    End If

    On Error GoTo 0

    legendWidth = 44 + maximumNameLength * 5

    If legendWidth < 80 Then legendWidth = 80
    If legendWidth > 190 Then legendWidth = 190

    legendHeight = 6 + selectedCount * 14

    Set backgroundShape = wsGrid.Shapes.AddShape( _
        Type:=msoShapeRectangle, _
        Left:=legendLeft, _
        Top:=legendTop, _
        Width:=legendWidth, _
        Height:=legendHeight)

    backgroundShape.Name = legendPrefix & "background"
    backgroundShape.Placement = xlMoveAndSize

    With backgroundShape

        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.Transparency = 0.05

        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(210, 210, 210)
        .Line.Weight = 0.5

        .Shadow.Visible = msoFalse
        .ZOrder msoBringToFront

    End With

    For seriesIndex = 1 To selectedCount

        displayName = Trim$(CStr( _
            wsData.Cells(1, baseDataColumn + seriesIndex).value2))

        rowTop = legendTop + 2 + (seriesIndex - 1) * 14
        lineCenterY = rowTop + 6

        Set lineShape = wsGrid.Shapes.AddLine( _
            BeginX:=legendLeft + 6, _
            BeginY:=lineCenterY, _
            EndX:=legendLeft + 25, _
            EndY:=lineCenterY)

        lineShape.Name = legendPrefix & "line_" & CStr(seriesIndex)
        lineShape.Placement = xlMoveAndSize

        With lineShape.Line
            .Visible = msoTrue
            .ForeColor.RGB = GetSeriesColor(seriesIndex)
            .Weight = 1.25
        End With

        lineShape.ZOrder msoBringToFront

        Set textShape = wsGrid.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=legendLeft + 29, _
            Top:=rowTop, _
            Width:=legendWidth - 34, _
            Height:=12)

        textShape.Name = legendPrefix & "text_" & CStr(seriesIndex)
        textShape.Placement = xlMoveAndSize

        With textShape

            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            .Shadow.Visible = msoFalse

            .TextFrame2.TextRange.text = displayName
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.WordWrap = msoFalse

            .TextFrame2.MarginLeft = 0
            .TextFrame2.MarginRight = 0
            .TextFrame2.MarginTop = 0
            .TextFrame2.MarginBottom = 0

            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

            With .TextFrame2.TextRange.Font
                .Name = "Arial"
                .Size = 8
                .Bold = msoFalse
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With

            .ZOrder msoBringToFront

        End With

    Next seriesIndex

End Sub

Private Sub DeleteGridLegendShapes( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long)

    Dim shapeIndex As Long
    Dim legendPrefix As String
    Dim shapeName As String

    legendPrefix = LCase$(GridShapeName(gridId, "legend_"))

    For shapeIndex = wsGrid.Shapes.count To 1 Step -1

        shapeName = LCase$(wsGrid.Shapes(shapeIndex).Name)

        If Left$(shapeName, Len(legendPrefix)) = legendPrefix Then
            wsGrid.Shapes(shapeIndex).Delete
        End If

    Next shapeIndex

End Sub

Private Sub HighlightLatestPoint( _
    ByVal chartSeries As Series, _
    ByVal visibleValues As Range, _
    ByVal seriesIndex As Long, _
    ByVal markerSize As Long, _
    ByVal numberFormat As String, _
    ByVal axisMinimum As Double, _
    ByVal axisMaximum As Double)

    Dim valuesData As Variant
    Dim pointIndex As Long
    Dim latestPointIndex As Long
    Dim latestValue As Double
    Dim highlightSize As Long
    Dim latestPoint As Point

    valuesData = visibleValues.value2

    For pointIndex = UBound(valuesData, 1) To 1 Step -1

        If Not IsError(valuesData(pointIndex, 1)) Then

            If IsNumeric(valuesData(pointIndex, 1)) Then
                latestPointIndex = pointIndex
                latestValue = CDbl(valuesData(pointIndex, 1))
                Exit For
            End If

        End If

    Next pointIndex

    If latestPointIndex <= 0 Then Exit Sub
    If latestPointIndex > chartSeries.Points.count Then Exit Sub

    highlightSize = MaxLong(7, markerSize + 3)

    Set latestPoint = chartSeries.Points(latestPointIndex)

    With latestPoint
        .MarkerStyle = xlMarkerStyleCircle
        .markerSize = highlightSize
        .MarkerForegroundColor = RGB(192, 0, 0)
        .MarkerBackgroundColor = RGB(192, 0, 0)
    End With

    On Error Resume Next

    chartSeries.HasDataLabels = False
    latestPoint.ApplyDataLabels Type:=xlDataLabelsShowValue

    If Err.Number = 0 Then

        With latestPoint.DataLabel
            .ShowValue = True
            .ShowSeriesName = False
            .ShowCategoryName = False
            .ShowLegendKey = False
            .numberFormat = numberFormat
            .Position = GetLatestLabelPosition( _
                seriesIndex, _
                latestValue, _
                axisMinimum, _
                axisMaximum)
            .Font.Name = "Arial"
            .Font.Size = 8
            .Font.Bold = True
            .Font.Color = RGB(192, 0, 0)
            .Format.Fill.Visible = msoFalse
            .Format.Line.Visible = msoFalse
        End With

    End If

    Err.Clear
    On Error GoTo 0

End Sub

Private Function GetLatestLabelPosition( _
    ByVal seriesIndex As Long, _
    ByVal latestValue As Double, _
    ByVal axisMinimum As Double, _
    ByVal axisMaximum As Double) As XlDataLabelPosition

    Dim axisSpan As Double
    Dim relativePosition As Double

    axisSpan = axisMaximum - axisMinimum

    If axisSpan <= 0 Then
        GetLatestLabelPosition = xlLabelPositionLeft
        Exit Function
    End If

    relativePosition = (latestValue - axisMinimum) / axisSpan

    If relativePosition >= 0.85 Then
        GetLatestLabelPosition = xlLabelPositionBelow
        Exit Function
    End If

    If relativePosition <= 0.15 Then
        GetLatestLabelPosition = xlLabelPositionAbove
        Exit Function
    End If

    Select Case seriesIndex
        Case 1
            GetLatestLabelPosition = xlLabelPositionLeft
        Case 2
            GetLatestLabelPosition = xlLabelPositionAbove
        Case 3
            GetLatestLabelPosition = xlLabelPositionBelow
        Case Else
            GetLatestLabelPosition = xlLabelPositionLeft
    End Select

End Function

'===============================================================================
' AXES
'===============================================================================

Private Sub ConfigureYAxis( _
    ByVal chartItem As chart, _
    ByVal axisMinimum As Double, _
    ByVal axisMaximum As Double, _
    ByVal majorUnit As Double, _
    ByVal minorUnit As Double, _
    ByVal numberFormat As String)

    With chartItem.Axes(xlValue)

        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = axisMinimum
        .MaximumScale = axisMaximum

        .MajorUnitIsAuto = False
        .majorUnit = majorUnit

        .MinorUnitIsAuto = False
        .minorUnit = minorUnit

        .Crosses = xlAxisCrossesMinimum

        .TickLabels.numberFormat = numberFormat
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

End Sub

Private Sub ConfigureXAxis( _
    ByVal chartItem As chart, _
    ByVal firstDate As Double, _
    ByVal lastDate As Double, _
    ByVal majorUnit As Double, _
    ByVal minorUnit As Double, _
    ByVal numberFormat As String, _
    ByVal labelRotation As Long)

    With chartItem.Axes(xlCategory)

        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = firstDate
        .MaximumScale = lastDate

        .MajorUnitIsAuto = False
        .majorUnit = majorUnit

        .MinorUnitIsAuto = False
        .minorUnit = minorUnit

        .TickLabels.numberFormat = numberFormat
        .TickLabels.Orientation = labelRotation
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

End Sub

Private Sub GetAdaptiveDateAxisSettings( _
    ByVal firstDate As Double, _
    ByVal lastDate As Double, _
    ByVal plotWidth As Double, _
    ByRef majorUnit As Double, _
    ByRef numberFormat As String, _
    ByRef labelRotation As Long)

    Dim spanDays As Double
    Dim estimatedLabelWidth As Double
    Dim maximumLabels As Long

    spanDays = lastDate - firstDate
    If spanDays <= 0 Then spanDays = 1

    Select Case spanDays

        Case Is <= 45
            numberFormat = "m/d"
            estimatedLabelWidth = 30

        Case Is <= 400

            If Year(CDate(firstDate)) = Year(CDate(lastDate)) Then
                numberFormat = "m/d"
                estimatedLabelWidth = 30
            Else
                numberFormat = "yyyy/m"
                estimatedLabelWidth = 45
            End If

        Case Is <= 2200
            numberFormat = "yyyy/m"
            estimatedLabelWidth = 45

        Case Else
            numberFormat = "yyyy"
            estimatedLabelWidth = 32

    End Select

    maximumLabels = CLng(Int(MaxDouble(120, plotWidth) / (estimatedLabelWidth + 8)))
    maximumLabels = ClampLong(maximumLabels, 4, 14)

    majorUnit = ChooseDateStepForMaximumLabels(spanDays, maximumLabels)
    labelRotation = 0

End Sub

Private Function ChooseDateStepForMaximumLabels( _
    ByVal spanDays As Double, _
    ByVal maximumLabels As Long) As Double

    Dim candidates As Variant
    Dim candidateIndex As Long
    Dim candidate As Double

    candidates = Array( _
        1#, 2#, 3#, 5#, 7#, 10#, 14#, 21#, _
        30#, 45#, 60#, 90#, 120#, 180#, _
        365#, 730#, 1095#, 1825#)

    For candidateIndex = LBound(candidates) To UBound(candidates)

        candidate = CDbl(candidates(candidateIndex))

        If spanDays / candidate + 1 <= maximumLabels Then
            ChooseDateStepForMaximumLabels = candidate
            Exit Function
        End If

    Next candidateIndex

    ChooseDateStepForMaximumLabels = CDbl(candidates(UBound(candidates)))

End Function

Private Sub GetVisibleCacheMinMax( _
    ByVal wsData As Worksheet, _
    ByVal baseDataColumn As Long, _
    ByVal selectedCount As Long, _
    ByVal startIndex As Long, _
    ByVal windowPoints As Long, _
    ByRef minimumValue As Double, _
    ByRef maximumValue As Double)

    Dim seriesIndex As Long
    Dim pointIndex As Long
    Dim valuesData As Variant
    Dim currentValue As Double
    Dim valueFound As Boolean

    For seriesIndex = 1 To selectedCount

        valuesData = wsData.Cells( _
            startIndex + 1, _
            baseDataColumn + seriesIndex).Resize( _
                windowPoints, _
                1).value2

        For pointIndex = 1 To UBound(valuesData, 1)

            If Not IsError(valuesData(pointIndex, 1)) Then

                If IsNumeric(valuesData(pointIndex, 1)) Then

                    currentValue = CDbl(valuesData(pointIndex, 1))

                    If Not valueFound Then
                        minimumValue = currentValue
                        maximumValue = currentValue
                        valueFound = True
                    Else
                        If currentValue < minimumValue Then minimumValue = currentValue
                        If currentValue > maximumValue Then maximumValue = currentValue
                    End If

                End If

            End If

        Next pointIndex

    Next seriesIndex

    If Not valueFound Then
        Err.Raise vbObjectError + 2001, , "No visible numeric data were found."
    End If

End Sub

Private Sub CalculateYAxisRange( _
    ByVal dataMinimum As Double, _
    ByVal dataMaximum As Double, _
    ByVal yZoomPercent As Long, _
    ByRef axisMinimum As Double, _
    ByRef axisMaximum As Double)

    Dim dataSpan As Double
    Dim axisSpan As Double
    Dim axisCenter As Double

    dataSpan = dataMaximum - dataMinimum

    If dataSpan <= 0 Then
        dataSpan = Abs(dataMaximum) * 0.1
        If dataSpan <= 0 Then dataSpan = 0.01
    End If

    axisSpan = dataSpan * 1.12 * yZoomPercent / 100#
    axisCenter = (dataMinimum + dataMaximum) / 2

    axisMinimum = axisCenter - axisSpan / 2
    axisMaximum = axisCenter + axisSpan / 2

    If axisMaximum <= axisMinimum Then axisMaximum = axisMinimum + 0.01

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
        spanValue = MaxDouble(Abs(rawMaximum) * 0.1, 0.01)
        rawMinimum = rawMinimum - spanValue / 2
        rawMaximum = rawMaximum + spanValue / 2
    End If

    majorUnit = ChooseNiceNumericStep(rawMaximum - rawMinimum, 10)
    niceMinimum = FloorToStep(rawMinimum, majorUnit)
    niceMaximum = CeilingToStep(rawMaximum, majorUnit)

    If niceMaximum <= niceMinimum Then niceMaximum = niceMinimum + majorUnit

    minorUnit = majorUnit / 2

End Sub

Private Function ChooseNiceNumericStep( _
    ByVal spanValue As Double, _
    ByVal targetIntervals As Long) As Double

    Dim baseValues As Variant
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

    baseValues = Array(1#, 2#, 2.5, 5#, 10#)
    centralExponent = Int(Log(spanValue / targetIntervals) / Log(10#))
    bestScore = 1E+99

    For exponentValue = centralExponent - 2 To centralExponent + 2

        For baseIndex = LBound(baseValues) To UBound(baseValues)

            candidate = CDbl(baseValues(baseIndex)) * 10# ^ exponentValue

            If candidate > 0 Then

                intervalCount = spanValue / candidate
                score = Abs(intervalCount - targetIntervals)

                If intervalCount < 7 Then
                    score = score + 100 + (7 - intervalCount) * 10
                End If

                If intervalCount > 13 Then
                    score = score + 100 + (intervalCount - 13) * 10
                End If

                If score < bestScore Then
                    bestScore = score
                    bestCandidate = candidate
                End If

            End If

        Next baseIndex

    Next exponentValue

    If bestCandidate <= 0 Then bestCandidate = spanValue / targetIntervals

    ChooseNiceNumericStep = bestCandidate

End Function

Private Function FloorToStep(ByVal valueItem As Double, ByVal stepValue As Double) As Double

    If stepValue <= 0 Then
        FloorToStep = valueItem
    Else
        FloorToStep = Int(valueItem / stepValue) * stepValue
    End If

End Function

Private Function CeilingToStep(ByVal valueItem As Double, ByVal stepValue As Double) As Double

    If stepValue <= 0 Then
        CeilingToStep = valueItem
    Else
        CeilingToStep = -Int(-valueItem / stepValue) * stepValue
    End If

End Function

Private Function NumberFormatFromStep(ByVal stepValue As Double) As String

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

    decimalPlaces = MaxLong(0, -Int(Log(stepValue) / Log(10#)))
    scaledValue = stepValue * 10# ^ decimalPlaces

    If Abs(scaledValue - Round(scaledValue, 0)) > 0.0000001 Then
        decimalPlaces = decimalPlaces + 1
    End If

    If stepValue < 1 And decimalPlaces < 3 Then decimalPlaces = 3
    If decimalPlaces > 8 Then decimalPlaces = 8

    If decimalPlaces = 0 Then
        NumberFormatFromStep = "0"
    Else
        NumberFormatFromStep = "0." & String$(decimalPlaces, "0")
    End If

End Function

'===============================================================================
' GRID ACTIONS
'===============================================================================

Private Sub ChangeGridXWindow( _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal direction As Long)

    Dim rowNumber As Long
    Dim totalPoints As Long
    Dim oldWindowPoints As Long
    Dim oldStartIndex As Long
    Dim fixedEndIndex As Long

    Dim requestedWindow As Long
    Dim newWindowPoints As Long
    Dim newStartIndex As Long
    Dim minimumWindow As Long

    rowNumber = StateRow(gridId)

    totalPoints = GetLongOrDefault(wsState.Cells(rowNumber, ST_TOTAL_POINTS).value2, 0)
    If totalPoints <= 0 Then Exit Sub

    oldWindowPoints = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(rowNumber, ST_WINDOW_POINTS).value2, _
            DEFAULT_WINDOW_POINTS), _
        1, _
        totalPoints)

    oldStartIndex = ClampLong( _
        GetLongOrDefault( _
            wsState.Cells(rowNumber, ST_START_INDEX).value2, _
            totalPoints - oldWindowPoints + 1), _
        1, _
        totalPoints)

    fixedEndIndex = ClampLong(oldStartIndex + oldWindowPoints - 1, 1, totalPoints)

    requestedWindow = oldWindowPoints + direction * GetXZoomStep(oldWindowPoints)
    minimumWindow = WorksheetFunction.Min(MIN_WINDOW_POINTS, fixedEndIndex)

    newWindowPoints = ClampLong(requestedWindow, minimumWindow, fixedEndIndex)
    newStartIndex = fixedEndIndex - newWindowPoints + 1

    wsState.Cells(rowNumber, ST_WINDOW_POINTS).value2 = newWindowPoints
    wsState.Cells(rowNumber, ST_START_INDEX).value2 = newStartIndex

End Sub

Private Function GetXZoomStep(ByVal currentWindow As Long) As Long

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

Private Sub ApplyLatestGridWindow( _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal requestedPoints As Long)

    Dim rowNumber As Long
    Dim totalPoints As Long
    Dim windowPoints As Long
    Dim startIndex As Long

    rowNumber = StateRow(gridId)
    totalPoints = GetLongOrDefault(wsState.Cells(rowNumber, ST_TOTAL_POINTS).value2, 0)

    If totalPoints <= 0 Then Exit Sub

    If requestedPoints <= 0 Then
        windowPoints = totalPoints
    Else
        windowPoints = WorksheetFunction.Min(requestedPoints, totalPoints)
    End If

    startIndex = MaxLong(1, totalPoints - windowPoints + 1)

    wsState.Cells(rowNumber, ST_WINDOW_POINTS).value2 = windowPoints
    wsState.Cells(rowNumber, ST_START_INDEX).value2 = startIndex

End Sub

Private Sub AdjustGridYZoom( _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal changeValue As Long)

    Dim rowNumber As Long
    Dim currentValue As Long

    rowNumber = StateRow(gridId)
    currentValue = GetLongOrDefault( _
        wsState.Cells(rowNumber, ST_Y_ZOOM).value2, _
        DEFAULT_Y_ZOOM)

    wsState.Cells(rowNumber, ST_Y_ZOOM).value2 = ClampLong( _
        currentValue + changeValue, _
        Y_ZOOM_MIN, _
        Y_ZOOM_MAX)

End Sub

Private Sub AdjustGridMarkerSize( _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long, _
    ByVal changeValue As Long)

    Dim rowNumber As Long
    Dim currentValue As Long

    rowNumber = StateRow(gridId)
    currentValue = GetLongOrDefault( _
        wsState.Cells(rowNumber, ST_MARKER_SIZE).value2, _
        DEFAULT_MARKER_SIZE)

    wsState.Cells(rowNumber, ST_MARKER_SIZE).value2 = ClampLong( _
        currentValue + changeValue, _
        MARKER_SIZE_MIN, _
        MARKER_SIZE_MAX)

End Sub

Private Sub CalculateGridSourceSheets(ByVal wsGrid As Worksheet, ByVal gridId As Long)

    Dim sourceSheets As Object
    Dim slot As Long
    Dim sheetName As String
    Dim keyItem As Variant
    Dim sourceSheet As Worksheet

    Set sourceSheets = CreateObject("Scripting.Dictionary")

    For slot = 1 To SERIES_SLOT_COUNT

        sheetName = SafeCellText(GridDataSheetCell(wsGrid, gridId, slot))

        If Len(sheetName) > 0 Then
            If Not sourceSheets.Exists(sheetName) Then sourceSheets.Add sheetName, True
        End If

    Next slot

    For Each keyItem In sourceSheets.keys

        Set sourceSheet = GetWorksheet(CStr(keyItem))

        If Not sourceSheet Is Nothing Then
            sourceSheet.Calculate
        End If

    Next keyItem

    ' Intentionally do not call CalculateUntilAsyncQueriesDone.
    ' Bloomberg and other asynchronous functions may otherwise block indefinitely.

End Sub

'===============================================================================
' CONTROL STATE
'===============================================================================

Private Sub UpdateGridControlState( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim rowNumber As Long
    Dim totalPoints As Long
    Dim windowPoints As Long
    Dim startIndex As Long
    Dim endIndex As Long
    Dim maximumStartIndex As Long
    Dim yZoom As Long
    Dim markerSize As Long

    Dim scrollShape As Shape

    rowNumber = StateRow(gridId)

    totalPoints = GetLongOrDefault(wsState.Cells(rowNumber, ST_TOTAL_POINTS).value2, 0)
    windowPoints = GetLongOrDefault(wsState.Cells(rowNumber, ST_WINDOW_POINTS).value2, DEFAULT_WINDOW_POINTS)
    startIndex = GetLongOrDefault(wsState.Cells(rowNumber, ST_START_INDEX).value2, 1)
    yZoom = GetLongOrDefault(wsState.Cells(rowNumber, ST_Y_ZOOM).value2, DEFAULT_Y_ZOOM)
    markerSize = GetLongOrDefault(wsState.Cells(rowNumber, ST_MARKER_SIZE).value2, DEFAULT_MARKER_SIZE)

    If totalPoints > 0 Then

        windowPoints = ClampLong( _
            windowPoints, _
            WorksheetFunction.Min(MIN_WINDOW_POINTS, totalPoints), _
            totalPoints)

        maximumStartIndex = MaxLong(1, totalPoints - windowPoints + 1)
        startIndex = ClampLong(startIndex, 1, maximumStartIndex)
        endIndex = startIndex + windowPoints - 1

    Else

        maximumStartIndex = 1
        startIndex = 1
        endIndex = 0

    End If

    SetShapeCaption wsGrid, GridReadoutName(gridId, "XZOOM"), Format$(windowPoints, "#,##0")
    SetShapeCaption wsGrid, GridReadoutName(gridId, "YZOOM"), Format$(yZoom, "0") & "%"
    SetShapeCaption wsGrid, GridReadoutName(gridId, "POINT"), Format$(markerSize, "0")

    On Error Resume Next

    Set scrollShape = wsGrid.Shapes(GridShapeName(gridId, "xrange"))

    With scrollShape.ControlFormat
        .Min = 1
        .Max = ScrollLimit(maximumStartIndex)
        .SmallChange = 1
        .LargeChange = MaxLong(1, windowPoints \ 5)
        .value = ClampLong(startIndex, 1, ScrollLimit(maximumStartIndex))
    End With

    On Error GoTo 0

    UpdateGridRangeStatus wsGrid, gridId, startIndex, endIndex, totalPoints
    UpdateGridPeriodButtonStyles wsGrid, gridId, totalPoints, windowPoints

End Sub

Private Sub SetShapeCaption( _
    ByVal ws As Worksheet, _
    ByVal shapeName As String, _
    ByVal caption As String)

    On Error Resume Next
    ws.Shapes(shapeName).TextFrame2.TextRange.text = caption
    On Error GoTo 0

End Sub

Private Sub UpdateGridRangeStatus( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long, _
    ByVal totalPoints As Long)

    Dim baseRow As Long
    Dim baseColumn As Long
    Dim statusCell As Range

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    Set statusCell = wsGrid.Cells(baseRow + 19, baseColumn + 15)

    If totalPoints <= 0 Then
        statusCell.value2 = "No data"
    Else
        statusCell.value2 = _
            Format$(startIndex, "#,##0") & _
            " - " & _
            Format$(endIndex, "#,##0") & _
            " / " & _
            Format$(totalPoints, "#,##0")
    End If

End Sub

Private Sub UpdateGridPeriodButtonStyles( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByVal totalPoints As Long, _
    ByVal windowPoints As Long)

    SetGridPeriodButtonActive wsGrid, gridId, "1W", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_1W, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "1M", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_1M, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "3M", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_3M, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "6M", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_6M, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "1Y", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_1Y, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "3Y", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_3Y, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "5Y", _
        totalPoints > 0 And windowPoints = WorksheetFunction.Min(POINTS_5Y, totalPoints)

    SetGridPeriodButtonActive wsGrid, gridId, "MAX", _
        totalPoints > 0 And windowPoints = totalPoints

End Sub

Private Sub SetGridPeriodButtonActive( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByVal actionName As String, _
    ByVal isActive As Boolean)

    Dim buttonShape As Shape

    On Error Resume Next
    Set buttonShape = wsGrid.Shapes(GridPeriodButtonName(gridId, actionName))
    On Error GoTo 0

    If buttonShape Is Nothing Then Exit Sub

    If isActive Then

        buttonShape.Fill.ForeColor.RGB = RGB(205, 205, 205)
        buttonShape.Line.ForeColor.RGB = RGB(110, 110, 110)
        buttonShape.TextFrame2.TextRange.Font.Bold = msoTrue

    Else

        buttonShape.Fill.ForeColor.RGB = RGB(242, 242, 242)
        buttonShape.Line.ForeColor.RGB = RGB(170, 170, 170)
        buttonShape.TextFrame2.TextRange.Font.Bold = msoFalse

    End If

End Sub

'===============================================================================
' STATE / GRID HELPERS
'===============================================================================

Private Sub SyncGridSelectorsToState( _
    ByVal wsGrid As Worksheet, _
    ByVal wsState As Worksheet, _
    ByVal gridId As Long)

    Dim rowNumber As Long
    Dim slot As Long

    rowNumber = StateRow(gridId)

    For slot = 1 To SERIES_SLOT_COUNT

        wsState.Cells(rowNumber, DataSheetStateColumn(slot)).value2 = _
            SafeCellText(GridDataSheetCell(wsGrid, gridId, slot))

        wsState.Cells(rowNumber, SeriesStateColumn(slot)).value2 = _
            SafeCellText(GridSeriesCell(wsGrid, gridId, slot))

    Next slot

End Sub

Private Function StateRow(ByVal gridId As Long) As Long
    StateRow = gridId + 1
End Function

Private Function GridBaseRow(ByVal gridId As Long) As Long

    Dim wsState As Worksheet
    Dim gridColumns As Long
    Dim gridRowIndex As Long

    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    gridColumns = GetLongOrDefault(wsState.Range(STATE_META_GRID_COLUMNS).value2, 1)

    gridRowIndex = ((gridId - 1) \ gridColumns) + 1
    GridBaseRow = 1 + (gridRowIndex - 1) * GRID_ROW_STRIDE

End Function

Private Function GridBaseColumn(ByVal gridId As Long) As Long

    Dim wsState As Worksheet
    Dim gridColumns As Long
    Dim gridColumnIndex As Long

    Set wsState = GetWorksheet(STATE_SHEET_NAME)
    gridColumns = GetLongOrDefault(wsState.Range(STATE_META_GRID_COLUMNS).value2, 1)

    gridColumnIndex = ((gridId - 1) Mod gridColumns) + 1
    GridBaseColumn = 1 + (gridColumnIndex - 1) * GRID_COLUMN_STRIDE

End Function

Private Function GridChartRange(ByVal wsGrid As Worksheet, ByVal gridId As Long) As Range

    Dim baseRow As Long
    Dim baseColumn As Long

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    Set GridChartRange = wsGrid.Range( _
        wsGrid.Cells(baseRow + GRID_CHART_FIRST_REL_ROW - 1, baseColumn + GRID_CHART_FIRST_REL_COL - 1), _
        wsGrid.Cells(baseRow + GRID_CHART_LAST_REL_ROW - 1, baseColumn + GRID_CHART_LAST_REL_COL - 1))

End Function

Private Function GridDataSheetCell( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByVal slot As Long) As Range

    Dim baseRow As Long
    Dim baseColumn As Long

    ValidateSlot slot

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    Set GridDataSheetCell = wsGrid.Cells(baseRow + slot * 2 - 1, baseColumn + 1)

End Function

Private Function GridSeriesCell( _
    ByVal wsGrid As Worksheet, _
    ByVal gridId As Long, _
    ByVal slot As Long) As Range

    Dim baseRow As Long
    Dim baseColumn As Long

    ValidateSlot slot

    baseRow = GridBaseRow(gridId)
    baseColumn = GridBaseColumn(gridId)

    Set GridSeriesCell = wsGrid.Cells(baseRow + slot * 2, baseColumn + 1)

End Function

Private Function DataSheetStateColumn(ByVal slot As Long) As Long

    ValidateSlot slot

    Select Case slot
        Case 1
            DataSheetStateColumn = ST_DATA_SHEET_1
        Case 2
            DataSheetStateColumn = ST_DATA_SHEET_2
        Case 3
            DataSheetStateColumn = ST_DATA_SHEET_3
    End Select

End Function

Private Function SeriesStateColumn(ByVal slot As Long) As Long

    ValidateSlot slot

    Select Case slot
        Case 1
            SeriesStateColumn = ST_SERIES_1
        Case 2
            SeriesStateColumn = ST_SERIES_2
        Case 3
            SeriesStateColumn = ST_SERIES_3
    End Select

End Function

Private Function DataBaseColumn(ByVal gridId As Long) As Long
    DataBaseColumn = 1 + (gridId - 1) * CACHE_COLUMNS_PER_GRID
End Function

Private Function SeriesListColumn(ByVal gridId As Long, ByVal slot As Long) As Long

    ValidateSlot slot
    SeriesListColumn = STATE_SERIES_LIST_FIRST_COLUMN + (gridId - 1) * SERIES_SLOT_COUNT + slot - 1

End Function

Private Function GridSeriesListName(ByVal gridId As Long, ByVal slot As Long) As String

    ValidateSlot slot
    GridSeriesListName = "_CG_G" & CStr(gridId) & "_S" & CStr(slot)

End Function

Private Function GridChartName(ByVal gridId As Long) As String
    GridChartName = SHAPE_PREFIX & CStr(gridId) & CHART_NAME_SUFFIX
End Function

Private Function GridShapeName(ByVal gridId As Long, ByVal suffixText As String) As String
    GridShapeName = SHAPE_PREFIX & CStr(gridId) & "_" & suffixText
End Function

Private Function GridReadoutName(ByVal gridId As Long, ByVal controlKey As String) As String
    GridReadoutName = GridShapeName(gridId, LCase$(controlKey) & "_readout")
End Function

Private Function GridPeriodButtonName(ByVal gridId As Long, ByVal actionName As String) As String
    GridPeriodButtonName = GridShapeName(gridId, "period_" & LCase$(Replace$(actionName, "_", "")))
End Function

Private Function BuildCallerMetadata(ByVal gridId As Long, ByVal actionName As String) As String
    BuildCallerMetadata = "CG|" & CStr(gridId) & "|" & actionName
End Function

Private Function ParseCallerMetadata( _
    ByVal metadataText As String, _
    ByRef gridId As Long, _
    ByRef actionName As String) As Boolean

    Dim parts As Variant

    parts = Split(metadataText, "|")

    If UBound(parts) <> 2 Then Exit Function
    If StrComp(CStr(parts(0)), "CG", vbTextCompare) <> 0 Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    gridId = CLng(parts(1))
    actionName = CStr(parts(2))

    If gridId <= 0 Or Len(actionName) = 0 Then Exit Function

    ParseCallerMetadata = True

End Function

Private Function GetGridSelectorLocation( _
    ByVal targetCell As Range, _
    ByVal gridColumns As Long, _
    ByVal gridRows As Long, _
    ByRef gridId As Long, _
    ByRef slot As Long, _
    ByRef isDataSheetCell As Boolean) As Boolean

    Dim gridRowIndex As Long
    Dim gridColumnIndex As Long
    Dim baseRow As Long
    Dim baseColumn As Long
    Dim relativeRow As Long
    Dim relativeColumn As Long

    gridColumnIndex = ((targetCell.Column - 1) \ GRID_COLUMN_STRIDE) + 1
    gridRowIndex = ((targetCell.Row - 1) \ GRID_ROW_STRIDE) + 1

    If gridColumnIndex < 1 Or gridColumnIndex > gridColumns Then Exit Function
    If gridRowIndex < 1 Or gridRowIndex > gridRows Then Exit Function

    baseColumn = 1 + (gridColumnIndex - 1) * GRID_COLUMN_STRIDE
    baseRow = 1 + (gridRowIndex - 1) * GRID_ROW_STRIDE

    relativeColumn = targetCell.Column - baseColumn + 1
    relativeRow = targetCell.Row - baseRow + 1

    If relativeColumn <> 2 Then Exit Function
    If relativeRow < 2 Or relativeRow > 7 Then Exit Function

    gridId = (gridRowIndex - 1) * gridColumns + gridColumnIndex
    slot = relativeRow \ 2
    isDataSheetCell = (relativeRow Mod 2 = 0)

    GetGridSelectorLocation = True

End Function

Private Sub ValidateSlot(ByVal slot As Long)

    If slot < 1 Or slot > SERIES_SLOT_COUNT Then
        Err.Raise vbObjectError + 1000, , "Invalid series slot: " & CStr(slot)
    End If

End Sub

'===============================================================================
' DEFAULT SELECTIONS
'===============================================================================

Private Function GetDefaultDataSheet() As String

    Dim ws As Worksheet

    If IsSelectableDataSheet("JGBData") Then
        GetDefaultDataSheet = "JGBData"
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets

        If IsSelectableDataSheet(ws.Name) Then
            GetDefaultDataSheet = ws.Name
            Exit Function
        End If

    Next ws

End Function

Private Function GetDefaultSeriesForGrid( _
    ByVal dataSheetName As String, _
    ByVal gridId As Long) As String

    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim headerCount As Long
    Dim targetColumn As Long
    Dim seriesName As String

    Set ws = GetWorksheet(dataSheetName)

    If ws Is Nothing Then
        GetDefaultSeriesForGrid = NONE_SERIES_TEXT
        Exit Function
    End If

    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    If lastColumn < 2 Then
        GetDefaultSeriesForGrid = NONE_SERIES_TEXT
        Exit Function
    End If

    headerCount = lastColumn - 1
    targetColumn = 2 + ((gridId - 1) Mod headerCount)
    seriesName = Trim$(CStr(ws.Cells(1, targetColumn).value2))

    If Len(seriesName) = 0 Then seriesName = FirstSeriesName(dataSheetName)

    GetDefaultSeriesForGrid = seriesName

End Function

Private Function FirstSeriesName(ByVal dataSheetName As String) As String

    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim columnIndex As Long
    Dim headerText As String

    Set ws = GetWorksheet(dataSheetName)

    If ws Is Nothing Then
        FirstSeriesName = NONE_SERIES_TEXT
        Exit Function
    End If

    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For columnIndex = 2 To lastColumn

        If Not IsError(ws.Cells(1, columnIndex).value2) Then

            headerText = Trim$(CStr(ws.Cells(1, columnIndex).value2))

            If Len(headerText) > 0 Then
                FirstSeriesName = headerText
                Exit Function
            End If

        End If

    Next columnIndex

    FirstSeriesName = NONE_SERIES_TEXT

End Function

Private Function IsSeriesAvailable( _
    ByVal dataSheetName As String, _
    ByVal seriesName As String) As Boolean

    Dim ws As Worksheet

    If Len(seriesName) = 0 Then Exit Function
    If StrComp(seriesName, NONE_SERIES_TEXT, vbTextCompare) = 0 Then Exit Function

    Set ws = GetWorksheet(dataSheetName)
    If ws Is Nothing Then Exit Function

    IsSeriesAvailable = (FindSeriesColumn(ws, seriesName) > 0)

End Function

Private Function IsSelectableDataSheet(ByVal sheetName As String) As Boolean

    If Len(sheetName) = 0 Then Exit Function

    If StrComp(sheetName, GRID_SHEET_NAME, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, STATE_SHEET_NAME, vbTextCompare) = 0 Then Exit Function
    If StrComp(sheetName, DATA_SHEET_NAME, vbTextCompare) = 0 Then Exit Function

    IsSelectableDataSheet = WorksheetExists(sheetName)

End Function

'===============================================================================
' WORKSHEET PREPARATION
'===============================================================================

Private Sub PrepareGridWorksheet(ByVal wsGrid As Worksheet)

    Dim shapeIndex As Long
    Dim chartIndex As Long
    Dim usedRange As Range

    wsGrid.Activate
    wsGrid.Range("A1").Select

    For chartIndex = wsGrid.ChartObjects.count To 1 Step -1
        wsGrid.ChartObjects(chartIndex).Delete
    Next chartIndex

    For shapeIndex = wsGrid.Shapes.count To 1 Step -1
        wsGrid.Shapes(shapeIndex).Delete
    Next shapeIndex

    Set usedRange = wsGrid.usedRange

    On Error Resume Next
    usedRange.UnMerge
    On Error GoTo 0

    usedRange.Clear

End Sub

Private Sub PrepareStateWorksheet( _
    ByVal wsState As Worksheet, _
    ByVal gridColumns As Long, _
    ByVal gridRows As Long, _
    ByVal gridCount As Long)

    Dim usedRange As Range
    Dim nameIndex As Long
    Dim nameItem As Name

    Set usedRange = wsState.usedRange
    usedRange.Clear

    For nameIndex = ThisWorkbook.names.count To 1 Step -1

        Set nameItem = ThisWorkbook.names(nameIndex)

        If LCase$(Left$(nameItem.Name, 4)) = "_cg_" _
            Or InStr(1, LCase$(nameItem.Name), "!_cg_", vbTextCompare) > 0 Then

            On Error Resume Next
            nameItem.Delete
            On Error GoTo 0

        End If

    Next nameIndex

    wsState.Cells(1, ST_GRID_ID).value2 = "GridId"
    wsState.Cells(1, ST_GRID_ROW).value2 = "GridRow"
    wsState.Cells(1, ST_GRID_COL).value2 = "GridCol"
    wsState.Cells(1, ST_DATA_SHEET_1).value2 = "DataSheet1"
    wsState.Cells(1, ST_SERIES_1).value2 = "Series1"
    wsState.Cells(1, ST_DATA_SHEET_2).value2 = "DataSheet2"
    wsState.Cells(1, ST_SERIES_2).value2 = "Series2"
    wsState.Cells(1, ST_DATA_SHEET_3).value2 = "DataSheet3"
    wsState.Cells(1, ST_SERIES_3).value2 = "Series3"
    wsState.Cells(1, ST_START_INDEX).value2 = "StartIndex"
    wsState.Cells(1, ST_WINDOW_POINTS).value2 = "WindowPoints"
    wsState.Cells(1, ST_TOTAL_POINTS).value2 = "TotalPoints"
    wsState.Cells(1, ST_Y_ZOOM).value2 = "YZoom"
    wsState.Cells(1, ST_MARKER_SIZE).value2 = "MarkerSize"
    wsState.Cells(1, ST_AXIS_MIN).value2 = "AxisMin"
    wsState.Cells(1, ST_AXIS_MAX).value2 = "AxisMax"
    wsState.Cells(1, ST_SELECTION_SIGNATURE).value2 = "Signature"
    wsState.Cells(1, ST_SELECTED_COUNT).value2 = "SelectedCount"
    wsState.Cells(1, ST_LAST_ACTION).value2 = "LastAction"

    wsState.Range(STATE_META_GRID_COLUMNS).value2 = gridColumns
    wsState.Range(STATE_META_GRID_ROWS).value2 = gridRows
    wsState.Range(STATE_META_GRID_COUNT).value2 = gridCount

End Sub

Private Sub PrepareDataWorksheet(ByVal wsData As Worksheet, ByVal gridCount As Long)

    Dim requiredColumns As Long
    Dim usedRows As Long
    Dim usedColumns As Long

    requiredColumns = gridCount * CACHE_COLUMNS_PER_GRID
    usedRows = wsData.usedRange.rows.count
    usedColumns = MaxLong(wsData.usedRange.Columns.count, requiredColumns)

    wsData.Cells(1, 1).Resize(MaxLong(1, usedRows), MaxLong(1, usedColumns)).ClearContents

End Sub

'===============================================================================
' DATA UTILITIES
'===============================================================================

Private Function FindSeriesColumn( _
    ByVal ws As Worksheet, _
    ByVal seriesName As String) As Long

    Dim lastColumn As Long
    Dim headers As Variant
    Dim columnIndex As Long

    If Len(Trim$(seriesName)) = 0 Then Exit Function

    lastColumn = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If lastColumn < 2 Then Exit Function

    headers = ws.Range(ws.Cells(1, 2), ws.Cells(1, lastColumn)).value2

    For columnIndex = 1 To UBound(headers, 2)

        If Not IsError(headers(1, columnIndex)) Then

            If StrComp( _
                Trim$(CStr(headers(1, columnIndex))), _
                Trim$(seriesName), _
                vbTextCompare) = 0 Then

                FindSeriesColumn = columnIndex + 1
                Exit Function

            End If

        End If

    Next columnIndex

End Function

Private Function TryGetDateSerial( _
    ByVal sourceValue As Variant, _
    ByRef dateSerial As Double) As Boolean

    If IsError(sourceValue) Then Exit Function
    If IsEmpty(sourceValue) Then Exit Function

    If IsNumeric(sourceValue) Then

        dateSerial = Int(CDbl(sourceValue))
        If dateSerial > 0 Then TryGetDateSerial = True

    ElseIf IsDate(sourceValue) Then

        dateSerial = Int(CDbl(CDate(sourceValue)))
        TryGetDateSerial = True

    End If

End Function

Private Function TryGetDouble( _
    ByVal sourceValue As Variant, _
    ByRef numericValue As Double) As Boolean

    If IsError(sourceValue) Then Exit Function
    If IsEmpty(sourceValue) Then Exit Function

    If VarType(sourceValue) = vbString Then
        If Len(Trim$(CStr(sourceValue))) = 0 Then Exit Function
    End If

    If IsNumeric(sourceValue) Then
        numericValue = CDbl(sourceValue)
        TryGetDouble = True
    End If

End Function

Private Function DateKey(ByVal dateSerial As Variant) As String
    DateKey = CStr(CLng(Int(CDbl(dateSerial))))
End Function

Private Sub QuickSortDoubles( _
    ByRef values() As Variant, _
    ByVal firstIndex As Long, _
    ByVal lastIndex As Long)

    Dim lowIndex As Long
    Dim highIndex As Long
    Dim pivotValue As Double
    Dim temporaryValue As Variant

    lowIndex = firstIndex
    highIndex = lastIndex
    pivotValue = CDbl(values((firstIndex + lastIndex) \ 2))

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

    If firstIndex < highIndex Then QuickSortDoubles values, firstIndex, highIndex
    If lowIndex < lastIndex Then QuickSortDoubles values, lowIndex, lastIndex

End Sub

Private Function BuildSelectionSignature( _
    ByRef selectedSheets() As String, _
    ByRef selectedSeries() As String, _
    ByVal selectedCount As Long) As String

    Dim seriesIndex As Long
    Dim resultText As String

    For seriesIndex = 1 To selectedCount

        If Len(resultText) > 0 Then resultText = resultText & "|"

        resultText = resultText & _
            selectedSheets(seriesIndex) & _
            "::" & _
            selectedSeries(seriesIndex)

    Next seriesIndex

    BuildSelectionSignature = resultText

End Function

'===============================================================================
' COLORS
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
' PARSING
'===============================================================================

Private Function ParseGridSpecification( _
    ByVal specification As String, _
    ByRef gridColumns As Long, _
    ByRef gridRows As Long) As Boolean

    Dim normalizedText As String
    Dim parts As Variant

    normalizedText = Trim$(specification)
    normalizedText = Replace$(normalizedText, " ", vbNullString)
    normalizedText = Replace$(normalizedText, vbTab, vbNullString)

    normalizedText = Replace$(normalizedText, ChrW$(215), "*")
    normalizedText = Replace$(normalizedText, ChrW$(-246), "*")
    normalizedText = Replace$(normalizedText, "x", "*", 1, -1, vbTextCompare)

    parts = Split(normalizedText, "*")

    If UBound(parts) <> 1 Then Exit Function
    If Len(CStr(parts(0))) = 0 Or Len(CStr(parts(1))) = 0 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function

    gridColumns = CLng(parts(0))
    gridRows = CLng(parts(1))

    If gridColumns <= 0 Or gridRows <= 0 Then Exit Function
    If gridColumns > 20 Or gridRows > 20 Then Exit Function

    ParseGridSpecification = True

End Function

'===============================================================================
' APPLICATION STATE AND ERROR HANDLING
'===============================================================================

Private Sub EnterBusyMode()
    mBusyDepth = mBusyDepth + 1
End Sub

Private Sub LeaveBusyMode()
    If mBusyDepth > 0 Then mBusyDepth = mBusyDepth - 1
End Sub

Private Sub SaveAndApplyFastMode( _
    ByRef oldScreenUpdating As Boolean, _
    ByRef oldEnableEvents As Boolean, _
    ByRef oldDisplayAlerts As Boolean, _
    ByRef oldCalculation As XlCalculation, _
    ByRef oldStatusBar As Variant)

    With Application

        oldScreenUpdating = .ScreenUpdating
        oldEnableEvents = .EnableEvents
        oldDisplayAlerts = .DisplayAlerts
        oldCalculation = .Calculation
        oldStatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual

    End With

End Sub

Private Sub RestoreApplicationState( _
    ByVal oldScreenUpdating As Boolean, _
    ByVal oldEnableEvents As Boolean, _
    ByVal oldDisplayAlerts As Boolean, _
    ByVal oldCalculation As XlCalculation, _
    ByVal oldStatusBar As Variant)

    On Error Resume Next

    With Application
        .Calculation = oldCalculation
        .DisplayAlerts = oldDisplayAlerts
        .EnableEvents = oldEnableEvents
        .ScreenUpdating = oldScreenUpdating
        .StatusBar = oldStatusBar
    End With

    On Error GoTo 0

End Sub

Private Sub ShowDetailedError( _
    ByVal procedureName As String, _
    ByVal stageName As String, _
    ByVal errorNumber As Long, _
    ByVal errorDescription As String)

    MsgBox _
        "ChartGrid operation failed." & vbCrLf & vbCrLf & _
        "Procedure: " & procedureName & vbCrLf & _
        "Stage: " & stageName & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Description: " & errorDescription, _
        vbExclamation

End Sub

'===============================================================================
' GENERAL UTILITIES
'===============================================================================

Private Function GetWorksheet(ByVal sheetName As String) As Worksheet

    If Len(Trim$(sheetName)) = 0 Then Exit Function

    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

End Function

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet

    Set GetOrCreateWorksheet = GetWorksheet(sheetName)

    If GetOrCreateWorksheet Is Nothing Then

        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))

        GetOrCreateWorksheet.Name = sheetName

    End If

End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    Set ws = GetWorksheet(sheetName)
    WorksheetExists = Not ws Is Nothing

End Function

Private Function SafeCellText(ByVal target As Range) As String

    If IsError(target.value2) Then Exit Function
    If IsEmpty(target.value2) Then Exit Function

    SafeCellText = Trim$(CStr(target.value2))

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

Private Function ClampLong( _
    ByVal valueItem As Long, _
    ByVal minimumValue As Long, _
    ByVal maximumValue As Long) As Long

    If maximumValue < minimumValue Then maximumValue = minimumValue

    If valueItem < minimumValue Then
        ClampLong = minimumValue
    ElseIf valueItem > maximumValue Then
        ClampLong = maximumValue
    Else
        ClampLong = valueItem
    End If

End Function

Private Function MaxLong(ByVal value1 As Long, ByVal value2 As Long) As Long

    If value1 >= value2 Then
        MaxLong = value1
    Else
        MaxLong = value2
    End If

End Function

Private Function MaxDouble(ByVal value1 As Double, ByVal value2 As Double) As Double

    If value1 >= value2 Then
        MaxDouble = value1
    Else
        MaxDouble = value2
    End If

End Function

Private Function ScrollLimit(ByVal valueItem As Long) As Long

    Select Case valueItem
        Case Is < 1
            ScrollLimit = 1
        Case Is > 30000
            ScrollLimit = 30000
        Case Else
            ScrollLimit = valueItem
    End Select

End Function

Private Function SheetCellReference( _
    ByVal ws As Worksheet, _
    ByVal target As Range) As String

    SheetCellReference = _
        "'" & Replace(ws.Name, "'", "''") & "'!" & target.Address

End Function

Private Function SheetRangeFormula( _
    ByVal ws As Worksheet, _
    ByVal target As Range) As String

    SheetRangeFormula = _
        "'" & Replace(ws.Name, "'", "''") & "'!" & target.Address

End Function

Private Sub DeleteWorkbookName(ByVal nameText As String)

    On Error Resume Next
    ThisWorkbook.names(nameText).Delete
    On Error GoTo 0

End Sub

Private Sub ClearUsedColumn(ByVal ws As Worksheet, ByVal columnNumber As Long)

    Dim lastRow As Long

    lastRow = ws.Cells(ws.rows.count, columnNumber).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    ws.Cells(1, columnNumber).Resize(lastRow, 1).ClearContents

End Sub

Private Sub ApplyThinBorders(ByVal target As Range, ByVal borderColor As Long)

    With target.Borders
        .LineStyle = xlContinuous
        .Color = borderColor
        .Weight = xlThin
    End With

End Sub
