Option Explicit

'========================================================
' ユーザー設定
'========================================================

'元データを何倍してから日次変化・共分散を計算するか
'
'例:
'  元データが 1.250 = 1.250% の形式
'    0.01の変化が1bpなので 100#
'
'  元データが 0.01250 = 1.250% の小数形式
'    0.0001の変化が1bpなので 10000#
'
'  元データがすでにbp単位
'    1#
Private Const SOURCE_DATA_MULTIPLIER As Double = 100#

'出力シート名
Private Const OUTPUT_SHEET_NAME As String = "Cov"

'共分散行列の出力開始行
Private Const OUTPUT_START_ROW As Long = 11

'共分散行列ブロック間の空白列数
Private Const GAP_COLUMNS As Long = 2

'========================================================
' 30日・60日・120日の共分散行列を1枚のシートへ出力
'
' Source sheet:
'   Row 1    : headers
'   Column A : dates
'   Column B onward: time-series levels
'
' Output sheet:
'   Cov
'
' Named ranges:
'   CovTable30
'   CovMatrix30
'   CovHeaders30
'   CovRowHeaders30
'
'   CovTable60
'   CovMatrix60
'   CovHeaders60
'   CovRowHeaders60
'
'   CovTable120
'   CovMatrix120
'   CovHeaders120
'   CovRowHeaders120
'
' Missing-data rule:
'   各Lookbackの計算に必要な水準範囲内に、
'   空白、Excelエラー、空文字、非数値が1つでもある系列は、
'   当該Lookbackでは無効系列とする。
'
'   無効系列に対応する共分散行列の行・列は #N/A とする。
'   有効系列同士の共分散は通常どおり計算する。
'========================================================
Public Sub BuildCovMatrices()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsCov As Worksheet

    Dim srcName As String
    Dim multiplier As Double

    Dim lookbacks As Variant
    Dim lookbackDays As Long
    Dim idx As Long

    Dim lastRow As Long
    Dim lastCol As Long
    Dim nSeries As Long

    Dim outputStartCol As Long
    Dim matrixBlockWidth As Long

    Dim validCount As Long
    Dim invalidCount As Long
    Dim invalidSeriesText As String

    Dim summaryText As String
    Dim detailText As String

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim stateCaptured As Boolean

    Dim currentStep As String
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo ErrHandler

    '====================================================
    ' Workbook取得
    '====================================================
    currentStep = "Workbook取得"

    Set wb = ActiveWorkbook

    If wb Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "アクティブなブックが見つかりません。"
    End If

    '====================================================
    ' 元データシート名入力
    '====================================================
    currentStep = "入力: 元データシート名"

    srcName = InputBox( _
        "共分散行列を計算する元データシート名を入力してください。" & vbCrLf & _
        "前提：" & vbCrLf & _
        "A列 = 日付" & vbCrLf & _
        "B列以降 = 時系列データ" & vbCrLf & _
        "1行目 = ヘッダー", _
        "30日・60日・120日 Cov Matrix作成" _
    )

    If Len(Trim$(srcName)) = 0 Then
        Err.Raise vbObjectError + 100, , _
            "シート名が入力されていません。"
    End If

    currentStep = "元データシート取得"

    Set wsSrc = GetSheetOrError_Local(wb, srcName)

    If StrComp(wsSrc.Name, OUTPUT_SHEET_NAME, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 101, , _
            "計算元シートに " & OUTPUT_SHEET_NAME & _
            " シートは指定できません。"
    End If

    '====================================================
    ' VBA上部で指定した倍率を取得
    '====================================================
    currentStep = "元データ倍率設定"

    multiplier = SOURCE_DATA_MULTIPLIER

    If multiplier <= 0# Then
        Err.Raise vbObjectError + 102, , _
            "SOURCE_DATA_MULTIPLIERは0より大きい数値を指定してください。"
    End If

    '====================================================
    ' Excel状態退避
    '====================================================
    currentStep = "Excel設定退避"

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    stateCaptured = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    '====================================================
    ' 最終行・最終列取得
    '====================================================
    currentStep = "最終行・最終列取得"

    lastRow = wsSrc.Cells(wsSrc.rows.count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column

    If lastCol < 2 Then
        Err.Raise vbObjectError + 200, , _
            "B列以降に時系列データがありません。"
    End If

    If lastRow < 3 Then
        Err.Raise vbObjectError + 201, , _
            "水準データが不足しています。"
    End If

    nSeries = lastCol - 1

    '120日分の変化幅には121個の水準データが必要
    If lastRow - 1 < 121 Then
        Err.Raise vbObjectError + 202, , _
            "120日共分散の計算に必要なデータ数が不足しています。" & vbCrLf & _
            "必要な水準データ数: 121" & vbCrLf & _
            "実際の水準データ数: " & (lastRow - 1)
    End If

    '====================================================
    ' ヘッダー検証
    '====================================================
    currentStep = "ヘッダー検証"

    ValidateHeaders_Local _
        wsSrc:=wsSrc, _
        lastCol:=lastCol

    '====================================================
    ' Covシート作成・クリア
    '====================================================
    currentStep = "Covシート作成・クリア"

    Set wsCov = GetOrCreateSheet_Local(wb, OUTPUT_SHEET_NAME)

    wsCov.Cells.Clear

    DeleteAllCovNamedRanges_Local wb

    '====================================================
    ' 配置設定
    '====================================================
    lookbacks = Array(30, 60, 120)

    '行ヘッダー1列 + データ列
    matrixBlockWidth = nSeries + 1

    '====================================================
    ' 3種類の共分散行列作成
    '====================================================
    summaryText = _
        "共分散行列の作成が完了しました。" & vbCrLf & _
        "計算元シート: " & wsSrc.Name & vbCrLf & _
        "系列数: " & nSeries & vbCrLf & _
        "元データ倍率: " & multiplier & vbCrLf & vbCrLf

    detailText = vbNullString

    For idx = LBound(lookbacks) To UBound(lookbacks)

        lookbackDays = CLng(lookbacks(idx))

        'A列から開始し、行列幅+空白列数ずつ右へ移動
        outputStartCol = _
            1 + idx * (matrixBlockWidth + GAP_COLUMNS)

        validCount = 0
        invalidCount = 0
        invalidSeriesText = vbNullString

        currentStep = _
            CStr(lookbackDays) & "日共分散行列作成"

        BuildSingleCovMatrix_Local _
            wb:=wb, _
            wsSrc:=wsSrc, _
            wsCov:=wsCov, _
            lastRow:=lastRow, _
            lastCol:=lastCol, _
            lookbackDays:=lookbackDays, _
            multiplier:=multiplier, _
            outputStartRow:=OUTPUT_START_ROW, _
            outputStartCol:=outputStartCol, _
            validCount:=validCount, _
            invalidCount:=invalidCount, _
            invalidSeriesText:=invalidSeriesText

        summaryText = summaryText & _
            "【" & lookbackDays & "日】" & vbCrLf & _
            "名前付き範囲: CovMatrix" & lookbackDays & vbCrLf & _
            "有効系列: " & validCount & vbCrLf & _
            "無効系列: " & invalidCount & vbCrLf & vbCrLf

        If invalidCount > 0 Then

            detailText = detailText & _
                "【" & lookbackDays & "日で無効】" & vbCrLf & _
                invalidSeriesText & vbCrLf & vbCrLf

        End If

    Next idx

    '====================================================
    ' シート共通フォーマット
    '====================================================
    currentStep = "Covシート共通フォーマット"

    FormatCovWorksheet_Local _
        ws:=wsCov, _
        matrixBlockWidth:=matrixBlockWidth, _
        gapColumns:=GAP_COLUMNS, _
        outputStartRow:=OUTPUT_START_ROW

    '====================================================
    ' Excel状態復元
    '====================================================
    If stateCaptured Then

        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts

    End If

    If Len(detailText) > 0 Then

        summaryText = summaryText & _
            "欠損または非数値を含む系列:" & vbCrLf & _
            detailText

    End If

    MsgBox summaryText, vbInformation

    Exit Sub

ErrHandler:

    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next

    If stateCaptured Then

        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
        Application.DisplayAlerts = oldDisplayAlerts

    End If

    On Error GoTo 0

    If Len(errDesc) = 0 Then
        errDesc = _
            "(Excelから詳細なエラー内容が返されていません)"
    End If

    MsgBox _
        "共分散行列の作成中にエラーが発生しました。" & vbCrLf & _
        "Step: " & currentStep & vbCrLf & _
        "Err.Number: " & errNum & vbCrLf & _
        "内容: " & errDesc, _
        vbCritical

End Sub

'========================================================
' 1つのLookbackについて共分散行列を作成
'========================================================
Private Sub BuildSingleCovMatrix_Local( _
    ByVal wb As Workbook, _
    ByVal wsSrc As Worksheet, _
    ByVal wsCov As Worksheet, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long, _
    ByVal lookbackDays As Long, _
    ByVal multiplier As Double, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long, _
    ByRef validCount As Long, _
    ByRef invalidCount As Long, _
    ByRef invalidSeriesText As String _
)

    Dim nSeries As Long
    Dim obsCount As Long

    Dim firstChangeRow As Long
    Dim firstLevelRow As Long

    Dim headers() As String
    Dim changes() As Double
    Dim means() As Double
    Dim covMat() As Double
    Dim outArr() As Variant

    Dim validSeries() As Boolean
    Dim invalidAddress() As String

    Dim cellValue As Variant

    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    nSeries = lastCol - 1
    obsCount = lookbackDays

    firstChangeRow = lastRow - lookbackDays + 1
    firstLevelRow = firstChangeRow - 1

    If firstLevelRow < 2 Then
        Err.Raise vbObjectError + 300 + lookbackDays, , _
            lookbackDays & _
            "日共分散の計算に必要なデータ数が不足しています。"
    End If

    '====================================================
    ' 配列確保
    '====================================================
    ReDim headers(1 To nSeries)
    ReDim changes(1 To obsCount, 1 To nSeries)
    ReDim means(1 To nSeries)
    ReDim covMat(1 To nSeries, 1 To nSeries)
    ReDim outArr(1 To nSeries + 1, 1 To nSeries + 1)

    ReDim validSeries(1 To nSeries)
    ReDim invalidAddress(1 To nSeries)

    '====================================================
    ' ヘッダー格納
    '====================================================
    For c = 2 To lastCol

        headers(c - 1) = _
            Trim$(CStr(wsSrc.Cells(1, c).value))

    Next c

    '====================================================
    ' 系列別欠損判定
    '====================================================
    validCount = 0
    invalidCount = 0
    invalidSeriesText = vbNullString

    For c = 2 To lastCol

        j = c - 1
        validSeries(j) = True

        For r = firstLevelRow To lastRow

            cellValue = wsSrc.Cells(r, c).Value2

            If Not IsValidNumericCell_Local(cellValue) Then

                validSeries(j) = False

                invalidAddress(j) = _
                    wsSrc.Cells(r, c).Address(False, False)

                Exit For

            End If

        Next r

        If validSeries(j) Then

            validCount = validCount + 1

        Else

            invalidCount = invalidCount + 1

            If Len(invalidSeriesText) > 0 Then
                invalidSeriesText = _
                    invalidSeriesText & ", "
            End If

            invalidSeriesText = _
                invalidSeriesText & _
                headers(j) & _
                " [" & invalidAddress(j) & "]"

        End If

    Next c

    '====================================================
    ' 1日変化幅作成
    '====================================================
    k = 1

    For r = firstChangeRow To lastRow

        For c = 2 To lastCol

            j = c - 1

            If validSeries(j) Then

                changes(k, j) = _
                    (CDbl(wsSrc.Cells(r, c).Value2) - _
                     CDbl(wsSrc.Cells(r - 1, c).Value2)) * _
                    multiplier

            End If

        Next c

        k = k + 1

    Next r

    '====================================================
    ' 平均変化幅計算
    '====================================================
    For j = 1 To nSeries

        If validSeries(j) Then

            means(j) = 0#

            For k = 1 To obsCount

                means(j) = _
                    means(j) + changes(k, j)

            Next k

            means(j) = _
                means(j) / obsCount

        End If

    Next j

    '====================================================
    ' 標本共分散行列計算
    '====================================================
    For i = 1 To nSeries

        For j = i To nSeries

            If validSeries(i) And validSeries(j) Then

                covMat(i, j) = 0#

                For k = 1 To obsCount

                    covMat(i, j) = _
                        covMat(i, j) + _
                        (changes(k, i) - means(i)) * _
                        (changes(k, j) - means(j))

                Next k

                covMat(i, j) = _
                    covMat(i, j) / (obsCount - 1)

                covMat(j, i) = covMat(i, j)

            End If

        Next j

    Next i

    '====================================================
    ' 出力配列作成
    '====================================================
    outArr(1, 1) = _
        "Cov " & lookbackDays & "D"

    For j = 1 To nSeries
        outArr(1, j + 1) = headers(j)
    Next j

    For i = 1 To nSeries

        outArr(i + 1, 1) = headers(i)

        For j = 1 To nSeries

            If validSeries(i) And validSeries(j) Then

                outArr(i + 1, j + 1) = _
                    covMat(i, j)

            Else

                outArr(i + 1, j + 1) = _
                    CVErr(xlErrNA)

            End If

        Next j

    Next i

    '====================================================
    ' メタ情報出力
    '====================================================
    WriteCovMetaBlock_Local _
        wsCov:=wsCov, _
        wsSrc:=wsSrc, _
        lookbackDays:=lookbackDays, _
        obsCount:=obsCount, _
        multiplier:=multiplier, _
        firstChangeRow:=firstChangeRow, _
        lastRow:=lastRow, _
        validCount:=validCount, _
        invalidCount:=invalidCount, _
        metaStartCol:=outputStartCol

    '====================================================
    ' 共分散行列出力
    '====================================================
    wsCov.Cells(outputStartRow, outputStartCol) _
         .Resize(nSeries + 1, nSeries + 1) _
         .value = outArr

    '====================================================
    ' 名前付き範囲作成
    '====================================================
    CreateCovNamedRanges_Local _
        wb:=wb, _
        wsCov:=wsCov, _
        nSeries:=nSeries, _
        outputStartRow:=outputStartRow, _
        outputStartCol:=outputStartCol, _
        lookbackDays:=lookbackDays

    '====================================================
    ' 個別ブロックフォーマット
    '====================================================
    FormatCovBlock_Local _
        ws:=wsCov, _
        nSeries:=nSeries, _
        outputStartRow:=outputStartRow, _
        outputStartCol:=outputStartCol

    '====================================================
    ' 無効系列ヘッダー強調
    '====================================================
    MarkInvalidSeries_Local _
        ws:=wsCov, _
        validSeries:=validSeries, _
        nSeries:=nSeries, _
        outputStartRow:=outputStartRow, _
        outputStartCol:=outputStartCol

End Sub

'========================================================
' ヘッダー検証
'========================================================
Private Sub ValidateHeaders_Local( _
    ByVal wsSrc As Worksheet, _
    ByVal lastCol As Long _
)

    Dim i As Long
    Dim j As Long

    Dim headerI As String
    Dim headerJ As String

    For i = 2 To lastCol

        headerI = _
            Trim$(CStr(wsSrc.Cells(1, i).value))

        If Len(headerI) = 0 Then

            Err.Raise vbObjectError + 400, , _
                "空白のヘッダーがあります。" & vbCrLf & _
                "列番号: " & i

        End If

    Next i

    For i = 2 To lastCol - 1

        headerI = _
            Trim$(CStr(wsSrc.Cells(1, i).value))

        For j = i + 1 To lastCol

            headerJ = _
                Trim$(CStr(wsSrc.Cells(1, j).value))

            If StrComp( _
                headerI, _
                headerJ, _
                vbTextCompare _
            ) = 0 Then

                Err.Raise vbObjectError + 401, , _
                    "ヘッダー名が重複しています。" & vbCrLf & _
                    "ヘッダー: " & headerI & vbCrLf & _
                    "列番号: " & i & " と " & j

            End If

        Next j

    Next i

End Sub

'========================================================
' 数値セルとして利用可能か判定
'========================================================
Private Function IsValidNumericCell_Local( _
    ByVal cellValue As Variant _
) As Boolean

    IsValidNumericCell_Local = False

    If IsError(cellValue) Then Exit Function
    If IsNull(cellValue) Then Exit Function
    If IsEmpty(cellValue) Then Exit Function

    If VarType(cellValue) = vbString Then

        If Len(Trim$(CStr(cellValue))) = 0 Then
            Exit Function
        End If

    End If

    If Not IsNumeric(cellValue) Then Exit Function

    IsValidNumericCell_Local = True

End Function

'========================================================
' 各共分散行列ブロック上部のメタ情報出力
'========================================================
Private Sub WriteCovMetaBlock_Local( _
    ByVal wsCov As Worksheet, _
    ByVal wsSrc As Worksheet, _
    ByVal lookbackDays As Long, _
    ByVal obsCount As Long, _
    ByVal multiplier As Double, _
    ByVal firstChangeRow As Long, _
    ByVal lastRow As Long, _
    ByVal validCount As Long, _
    ByVal invalidCount As Long, _
    ByVal metaStartCol As Long _
)

    With wsCov

        .Cells(1, metaStartCol).value = _
            "Covariance Matrix"

        .Cells(1, metaStartCol + 1).value = _
            lookbackDays & " Days"

        .Cells(2, metaStartCol).value = _
            "Source Sheet"

        .Cells(2, metaStartCol + 1).value = _
            wsSrc.Name

        .Cells(3, metaStartCol).value = _
            "As Of Date"

        .Cells(3, metaStartCol + 1).value = _
            wsSrc.Cells(lastRow, 1).value

        .Cells(4, metaStartCol).value = _
            "Lookback Days"

        .Cells(4, metaStartCol + 1).value = _
            lookbackDays

        .Cells(5, metaStartCol).value = _
            "Observation Count"

        .Cells(5, metaStartCol + 1).value = _
            obsCount

        .Cells(6, metaStartCol).value = _
            "Source Multiplier"

        .Cells(6, metaStartCol + 1).value = _
            multiplier

        .Cells(7, metaStartCol).value = _
            "Valid Series"

        .Cells(7, metaStartCol + 1).value = _
            validCount

        .Cells(8, metaStartCol).value = _
            "Invalid Series"

        .Cells(8, metaStartCol + 1).value = _
            invalidCount

        .Cells(9, metaStartCol).value = _
            "Change Row Range"

        .Cells(9, metaStartCol + 1).value = _
            firstChangeRow & ":" & lastRow

        .Cells(10, metaStartCol).value = _
            "Named Range"

        .Cells(10, metaStartCol + 1).value = _
            "CovMatrix" & lookbackDays

    End With

End Sub

'========================================================
' 名前付き範囲作成
'========================================================
Private Sub CreateCovNamedRanges_Local( _
    ByVal wb As Workbook, _
    ByVal wsCov As Worksheet, _
    ByVal nSeries As Long, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long, _
    ByVal lookbackDays As Long _
)

    Dim suffix As String

    Dim rngTable As Range
    Dim rngMatrix As Range
    Dim rngHeaders As Range
    Dim rngRowHeaders As Range

    suffix = CStr(lookbackDays)

    Set rngTable = _
        wsCov.Cells( _
            outputStartRow, _
            outputStartCol _
        ).Resize( _
            nSeries + 1, _
            nSeries + 1 _
        )

    Set rngMatrix = _
        wsCov.Cells( _
            outputStartRow + 1, _
            outputStartCol + 1 _
        ).Resize( _
            nSeries, _
            nSeries _
        )

    Set rngHeaders = _
        wsCov.Cells( _
            outputStartRow, _
            outputStartCol + 1 _
        ).Resize( _
            1, _
            nSeries _
        )

    Set rngRowHeaders = _
        wsCov.Cells( _
            outputStartRow + 1, _
            outputStartCol _
        ).Resize( _
            nSeries, _
            1 _
        )

    DeleteWorkbookNameIfExists_Local _
        wb, "CovTable" & suffix

    DeleteWorkbookNameIfExists_Local _
        wb, "CovMatrix" & suffix

    DeleteWorkbookNameIfExists_Local _
        wb, "CovHeaders" & suffix

    DeleteWorkbookNameIfExists_Local _
        wb, "CovRowHeaders" & suffix

    AddWorkbookRangeName_Local _
        wb:=wb, _
        rangeName:="CovTable" & suffix, _
        targetRange:=rngTable

    AddWorkbookRangeName_Local _
        wb:=wb, _
        rangeName:="CovMatrix" & suffix, _
        targetRange:=rngMatrix

    AddWorkbookRangeName_Local _
        wb:=wb, _
        rangeName:="CovHeaders" & suffix, _
        targetRange:=rngHeaders

    AddWorkbookRangeName_Local _
        wb:=wb, _
        rangeName:="CovRowHeaders" & suffix, _
        targetRange:=rngRowHeaders

End Sub

'========================================================
' Cov関連の既存名前付き範囲を一括削除
'========================================================
Private Sub DeleteAllCovNamedRanges_Local( _
    ByVal wb As Workbook _
)

    Dim lookbacks As Variant
    Dim prefixes As Variant

    Dim i As Long
    Dim j As Long

    lookbacks = Array("30", "60", "120")

    prefixes = Array( _
        "CovTable", _
        "CovMatrix", _
        "CovHeaders", _
        "CovRowHeaders" _
    )

    For i = LBound(lookbacks) To UBound(lookbacks)

        For j = LBound(prefixes) To UBound(prefixes)

            DeleteWorkbookNameIfExists_Local _
                wb, _
                CStr(prefixes(j)) & CStr(lookbacks(i))

        Next j

    Next i

End Sub

'========================================================
' ブックレベルの名前付き範囲を追加
'========================================================
Private Sub AddWorkbookRangeName_Local( _
    ByVal wb As Workbook, _
    ByVal rangeName As String, _
    ByVal targetRange As Range _
)

    Dim sheetNameEscaped As String
    Dim refersToText As String

    sheetNameEscaped = _
        Replace( _
            targetRange.Worksheet.Name, _
            "'", _
            "''" _
        )

    refersToText = _
        "='" & sheetNameEscaped & "'!" & _
        targetRange.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1 _
        )

    wb.names.Add _
        Name:=rangeName, _
        RefersTo:=refersToText, _
        Visible:=True

End Sub

'========================================================
' 指定したブックレベル名があれば削除
'========================================================
Private Sub DeleteWorkbookNameIfExists_Local( _
    ByVal wb As Workbook, _
    ByVal rangeName As String _
)

    Dim nm As Name

    On Error Resume Next

    Set nm = wb.names(rangeName)

    If Not nm Is Nothing Then
        nm.Delete
    End If

    Set nm = Nothing

    On Error GoTo 0

End Sub

'========================================================
' 各共分散行列ブロックのフォーマット
'========================================================
Private Sub FormatCovBlock_Local( _
    ByVal ws As Worksheet, _
    ByVal nSeries As Long, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long _
)

    Dim lastOutputRow As Long
    Dim lastOutputCol As Long

    Dim matrixRange As Range
    Dim numberRange As Range
    Dim headerRowRange As Range
    Dim headerColRange As Range
    Dim metaLabelRange As Range
    Dim metaValueRange As Range
    Dim diagRange As Range

    Dim i As Long
    Dim c As Long

    lastOutputRow = outputStartRow + nSeries
    lastOutputCol = outputStartCol + nSeries

    With ws

        Set matrixRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(lastOutputRow, lastOutputCol) _
        )

        Set numberRange = .Range( _
            .Cells(outputStartRow + 1, outputStartCol + 1), _
            .Cells(lastOutputRow, lastOutputCol) _
        )

        Set headerRowRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(outputStartRow, lastOutputCol) _
        )

        Set headerColRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(lastOutputRow, outputStartCol) _
        )

        Set metaLabelRange = .Range( _
            .Cells(1, outputStartCol), _
            .Cells(10, outputStartCol) _
        )

        Set metaValueRange = .Range( _
            .Cells(1, outputStartCol + 1), _
            .Cells(10, outputStartCol + 1) _
        )

        'メタ情報
        metaLabelRange.Font.Bold = True
        metaLabelRange.Interior.Color = _
            RGB(220, 230, 241)

        .Range( _
            .Cells(1, outputStartCol), _
            .Cells(10, outputStartCol + 1) _
        ).Borders.LineStyle = xlContinuous

        .Range( _
            .Cells(1, outputStartCol), _
            .Cells(10, outputStartCol + 1) _
        ).Borders.Color = RGB(200, 200, 200)

        metaValueRange.WrapText = False

        .Cells(1, outputStartCol).Interior.Color = _
            RGB(68, 114, 196)

        .Cells(1, outputStartCol).Font.Color = _
            RGB(255, 255, 255)

        .Cells(1, outputStartCol + 1).Interior.Color = _
            RGB(68, 114, 196)

        .Cells(1, outputStartCol + 1).Font.Color = _
            RGB(255, 255, 255)

        .Cells(1, outputStartCol).Font.Bold = True
        .Cells(1, outputStartCol + 1).Font.Bold = True

        '行列
        matrixRange.Borders.LineStyle = xlContinuous
        matrixRange.Borders.Color = RGB(210, 210, 210)

        headerRowRange.Font.Bold = True
        headerColRange.Font.Bold = True

        headerRowRange.Interior.Color = _
            RGB(220, 230, 241)

        headerColRange.Interior.Color = _
            RGB(220, 230, 241)

        headerRowRange.HorizontalAlignment = xlCenter
        headerColRange.HorizontalAlignment = xlLeft

        numberRange.numberFormat = "0.000000"
        numberRange.HorizontalAlignment = xlRight

        '対角分散
        For i = 1 To nSeries

            If diagRange Is Nothing Then

                Set diagRange = _
                    .Cells( _
                        outputStartRow + i, _
                        outputStartCol + i _
                    )

            Else

                Set diagRange = Union( _
                    diagRange, _
                    .Cells( _
                        outputStartRow + i, _
                        outputStartCol + i _
                    ) _
                )

            End If

        Next i

        If Not diagRange Is Nothing Then

            diagRange.Interior.Color = _
                RGB(255, 242, 204)

            diagRange.Font.Bold = True

        End If

        '列幅
        .Columns(outputStartCol).ColumnWidth = 18

        For c = outputStartCol + 1 To lastOutputCol
            .Columns(c).ColumnWidth = 11
        Next c

    End With

End Sub

'========================================================
' Covシート全体の共通フォーマット
'========================================================
Private Sub FormatCovWorksheet_Local( _
    ByVal ws As Worksheet, _
    ByVal matrixBlockWidth As Long, _
    ByVal gapColumns As Long, _
    ByVal outputStartRow As Long _
)

    Dim idx As Long
    Dim gapStartCol As Long
    Dim c As Long

    With ws

        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10

        .rows("1:10").RowHeight = 18
        .rows(outputStartRow).RowHeight = 18

        '30日と60日、60日と120日の間の空白列
        For idx = 0 To 1

            gapStartCol = _
                1 + _
                idx * (matrixBlockWidth + gapColumns) + _
                matrixBlockWidth

            For c = gapStartCol To _
                    gapStartCol + gapColumns - 1

                .Columns(c).ColumnWidth = 3

            Next c

        Next idx

        'Freeze panesは30日ブロックを基準
        On Error Resume Next

        .Activate
        ActiveWindow.FreezePanes = False

        .Cells(outputStartRow + 1, 2).Select

        ActiveWindow.FreezePanes = True

        On Error GoTo 0

    End With

End Sub

'========================================================
' 欠損系列の行・列ヘッダーを強調
'========================================================
Private Sub MarkInvalidSeries_Local( _
    ByVal ws As Worksheet, _
    ByRef validSeries() As Boolean, _
    ByVal nSeries As Long, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long _
)

    Dim i As Long

    For i = 1 To nSeries

        If Not validSeries(i) Then

            With ws.Cells( _
                outputStartRow, _
                outputStartCol + i _
            )

                .Interior.Color = RGB(244, 204, 204)
                .Font.Color = RGB(156, 0, 6)
                .Font.Bold = True

            End With

            With ws.Cells( _
                outputStartRow + i, _
                outputStartCol _
            )

                .Interior.Color = RGB(244, 204, 204)
                .Font.Color = RGB(156, 0, 6)
                .Font.Bold = True

            End With

        End If

    Next i

End Sub

'========================================================
' 指定シート取得
'========================================================
Private Function GetSheetOrError_Local( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
) As Worksheet

    On Error GoTo NotFound

    Set GetSheetOrError_Local = _
        wb.Worksheets(sheetName)

    Exit Function

NotFound:

    Err.Raise vbObjectError + 600, , _
        "指定されたシートが見つかりません: " & _
        sheetName

End Function

'========================================================
' シート取得または新規作成
'========================================================
Private Function GetOrCreateSheet_Local( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.count) _
        )

        ws.Name = sheetName

    End If

    Set GetOrCreateSheet_Local = ws

End Function

