
Option Explicit

'========================================================
' ユーザー設定
'========================================================

' 元データを何倍してから日次変化・共分散を計算するか
'
' 例：
'   1.250 = 1.250%       → 100#
'   0.01250 = 1.250%     → 10000#
'   元データがbp単位    → 1#

Private Const SOURCE_DATA_MULTIPLIER As Double = 100#

Private Const OUTPUT_SHEET_NAME As String = "Cov"

Private Const OUTPUT_START_ROW As Long = 11

Private Const GAP_COLUMNS As Long = 2


'========================================================
' BuildCovMatrices
'
' 30日・60日・120日の共分散行列を作成する。
'
' 元データ：
'   1行目    ヘッダー
'   A列      日付
'   B列以降  時系列水準
'
' 欠損処理：
'
' 1. A列が日付でない行は計算対象から除外する。
'
' 2. 30日Covでは直近31個の有効日付行を使用する。
'    60日Covでは直近61個の有効日付行を使用する。
'    120日Covでは直近121個の有効日付行を使用する。
'
' 3. 各系列について、その期間の水準値を検証する。
'
' 4. 1つでも欠損値・Excelエラー・非数値があれば、
'    当該Lookbackでは系列全体を無効とする。
'
' 5. 無効系列に対応する共分散行列の行・列は
'    すべて #N/A とする。
'
' 6. 有効系列同士の共分散は、同じ日付区間の
'    全変化幅を使用して計算する。
'
'========================================================

Public Sub BuildCovMatrices()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsCov As Worksheet

    Dim srcName As String

    Dim lastRow As Long
    Dim lastCol As Long
    Dim nSeries As Long

    Dim dateRows() As Long
    Dim dateCount As Long
    Dim skippedRows As Long

    Dim headers() As String

    Dim lookbacks As Variant
    Dim matrixBlocks(0 To 2) As Variant
    Dim validityBlocks(0 To 2) As Variant

    Dim validCounts(0 To 2) As Long
    Dim invalidCounts(0 To 2) As Long

    Dim invalidTexts(0 To 2) As String

    Dim firstRows(0 To 2) As Long
    Dim finalRows(0 To 2) As Long

    Dim idx As Long
    Dim lookbackDays As Long

    Dim matrixBlockWidth As Long
    Dim outputStartCol As Long

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
    ' 1. Workbook取得
    '====================================================

    currentStep = "Workbook取得"

    Set wb = ActiveWorkbook

    If wb Is Nothing Then

        Err.Raise vbObjectError + 100, , _
            "アクティブなブックが見つかりません。"

    End If


    '====================================================
    ' 2. 元データシート名入力
    '====================================================

    currentStep = "元データシート名入力"

    srcName = InputBox( _
        "共分散行列を計算する元データシート名を入力してください。" & _
        vbCrLf & vbCrLf & _
        "A列：日付" & vbCrLf & _
        "B列以降：時系列水準" & vbCrLf & _
        "1行目：ヘッダー", _
        "Build Covariance Matrices")

    srcName = Trim$(srcName)

    If Len(srcName) = 0 Then
        Exit Sub
    End If


    '====================================================
    ' 3. 元データシート取得
    '====================================================

    currentStep = "元データシート取得"

    Set wsSrc = GetSheetOrError_Local(wb, srcName)

    If StrComp( _
        wsSrc.Name, _
        OUTPUT_SHEET_NAME, _
        vbTextCompare) = 0 Then

        Err.Raise vbObjectError + 101, , _
            "Covシートを計算元に指定できません。"

    End If


    '====================================================
    ' 4. 倍率検証
    '====================================================

    currentStep = "倍率検証"

    If SOURCE_DATA_MULTIPLIER <= 0# Then

        Err.Raise vbObjectError + 102, , _
            "SOURCE_DATA_MULTIPLIERは0より大きい数値を" & _
            "指定してください。"

    End If


    '====================================================
    ' 5. Excel設定退避
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
    ' 6. 最終行・最終列
    '====================================================

    currentStep = "最終行・最終列取得"

    lastRow = wsSrc.Cells( _
        wsSrc.rows.count, 1).End(xlUp).Row

    lastCol = wsSrc.Cells( _
        1, wsSrc.Columns.count).End(xlToLeft).Column


    If lastCol < 2 Then

        Err.Raise vbObjectError + 200, , _
            "B列以降に時系列データがありません。"

    End If


    If lastRow < 2 Then

        Err.Raise vbObjectError + 201, , _
            "日付データがありません。"

    End If


    nSeries = lastCol - 1


    '====================================================
    ' 7. ヘッダー検証
    '====================================================

    currentStep = "ヘッダー検証"

    ValidateHeaders_Local _
        wsSrc, _
        lastCol, _
        headers


    '====================================================
    ' 8. 有効日付行の抽出
    '
    ' ここではB列以降の欠損は確認しない。
    '
    ' 日付として認識できる行だけを抽出する。
    '====================================================

    currentStep = "日付行抽出"

    GetValidDateRows_Local _
        wsSrc, _
        lastRow, _
        dateRows, _
        dateCount, _
        skippedRows


    '====================================================
    ' 9. 行列配置の検証
    '====================================================

    currentStep = "行列配置検証"

    lookbacks = Array(30, 60, 120)

    matrixBlockWidth = nSeries + 1


    If 3 * matrixBlockWidth + _
       2 * GAP_COLUMNS > wsSrc.Columns.count Then

        Err.Raise vbObjectError + 202, , _
            "系列数が多すぎるため、3つの共分散行列を" & _
            "横方向に配置できません。"

    End If


    If OUTPUT_START_ROW + nSeries > wsSrc.rows.count Then

        Err.Raise vbObjectError + 203, , _
            "共分散行列の行数がExcelの上限を超えます。"

    End If


    '====================================================
    ' 10. 共分散行列をメモリ上で計算
    '
    ' Covシートをクリアする前に、
    ' すべての計算処理を完了する。
    '====================================================

    For idx = 0 To 2

        lookbackDays = CLng(lookbacks(idx))

        currentStep = _
            CStr(lookbackDays) & "日共分散行列計算"


        matrixBlocks(idx) = _
            BuildSingleCovMatrix_Local( _
                wsSrc, _
                lastCol, _
                headers, _
                dateRows, _
                dateCount, _
                lookbackDays, _
                SOURCE_DATA_MULTIPLIER, _
                validityBlocks(idx), _
                validCounts(idx), _
                invalidCounts(idx), _
                invalidTexts(idx), _
                firstRows(idx), _
                finalRows(idx))

    Next idx


    '====================================================
    ' 11. Covシート取得
    '====================================================

    currentStep = "Covシート取得"

    Set wsCov = GetOrCreateSheet_Local( _
        wb, OUTPUT_SHEET_NAME)


    '====================================================
    ' 12. 既存出力のクリア
    '====================================================

    currentStep = "既存出力のクリア"

    wsCov.Cells.Clear

    DeleteAllCovNamedRanges_Local wb


    '====================================================
    ' 13. 共分散行列出力
    '====================================================

    For idx = 0 To 2

        lookbackDays = CLng(lookbacks(idx))

        outputStartCol = _
            1 + idx * (matrixBlockWidth + GAP_COLUMNS)

        currentStep = _
            CStr(lookbackDays) & "日共分散行列出力"


        '--------------------------------------------
        ' メタ情報
        '--------------------------------------------

        WriteCovMetaBlock_Local _
            wsCov, _
            wsSrc, _
            lookbackDays, _
            SOURCE_DATA_MULTIPLIER, _
            dateRows, _
            dateCount, _
            skippedRows, _
            firstRows(idx), _
            finalRows(idx), _
            validCounts(idx), _
            invalidCounts(idx), _
            outputStartCol


        '--------------------------------------------
        ' 共分散行列本体
        '--------------------------------------------

        wsCov.Cells( _
            OUTPUT_START_ROW, _
            outputStartCol) _
            .Resize( _
                nSeries + 1, _
                nSeries + 1).value2 = matrixBlocks(idx)


        '--------------------------------------------
        ' 名前付き範囲
        '--------------------------------------------

        CreateCovNamedRanges_Local _
            wb, _
            wsCov, _
            nSeries, _
            OUTPUT_START_ROW, _
            outputStartCol, _
            lookbackDays


        '--------------------------------------------
        ' 個別ブロック書式
        '--------------------------------------------

        FormatCovBlock_Local _
            wsCov, _
            nSeries, _
            OUTPUT_START_ROW, _
            outputStartCol


        '--------------------------------------------
        ' 無効系列ヘッダーを強調
        '--------------------------------------------

        MarkInvalidSeries_Local _
            wsCov, _
            validityBlocks(idx), _
            nSeries, _
            OUTPUT_START_ROW, _
            outputStartCol

    Next idx


    '====================================================
    ' 14. シート共通書式
    '====================================================

    currentStep = "Covシート共通書式"

    FormatCovWorksheet_Local _
        wsCov, _
        matrixBlockWidth, _
        GAP_COLUMNS, _
        OUTPUT_START_ROW


    '====================================================
    ' 15. 完了メッセージ作成
    '====================================================

    summaryText = _
        "共分散行列の作成が完了しました。" & _
        vbCrLf & vbCrLf & _
        "計算元シート: " & wsSrc.Name & vbCrLf & _
        "系列数: " & nSeries & vbCrLf & _
        "有効日付行: " & dateCount & vbCrLf & _
        "日付以外の除外行: " & skippedRows & _
        vbCrLf & vbCrLf


    detailText = vbNullString


    For idx = 0 To 2

        lookbackDays = CLng(lookbacks(idx))

        summaryText = summaryText & _
            "【" & lookbackDays & "日】" & vbCrLf & _
            "有効系列: " & validCounts(idx) & vbCrLf & _
            "無効系列: " & invalidCounts(idx) & _
            vbCrLf & vbCrLf


        If invalidCounts(idx) > 0 Then

            detailText = detailText & _
                "【" & lookbackDays & "日で無効】" & _
                vbCrLf & _
                invalidTexts(idx) & _
                vbCrLf & vbCrLf

        End If

    Next idx


    If Len(detailText) > 0 Then

        summaryText = summaryText & _
            "欠損・非数値・データ不足による無効系列:" & _
            vbCrLf & _
            detailText

    End If


    '====================================================
    ' 16. Excel設定復元
    '====================================================

    RestoreExcelState_Local _
        oldCalc, _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        stateCaptured

    stateCaptured = False


    MsgBox summaryText, vbInformation

    Exit Sub


ErrHandler:

    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next

    RestoreExcelState_Local _
        oldCalc, _
        oldScreenUpdating, _
        oldEnableEvents, _
        oldDisplayAlerts, _
        stateCaptured

    On Error GoTo 0


    MsgBox _
        "共分散行列の作成中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        "Step: " & currentStep & vbCrLf & _
        "Err.Number: " & errNum & vbCrLf & _
        "内容: " & errDesc, _
        vbCritical

End Sub


'========================================================
' 有効日付行の抽出
'
' A列が日付の行だけを抽出する。
'
' 空白行、EOD_Chg、1w Chg、1m Chg等は除外する。
'
' 日付行のB列以降に欠損があっても、
' 日付行そのものは除外しない。
'
'========================================================

Private Sub GetValidDateRows_Local( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByRef dateRows() As Long, _
    ByRef dateCount As Long, _
    ByRef skippedRows As Long _
)

    Dim r As Long

    Dim v As Variant

    Dim dateKey As Long
    Dim previousDate As Long

    Dim hasPreviousDate As Boolean


    ReDim dateRows(1 To lastRow - 1)

    dateCount = 0
    skippedRows = 0

    hasPreviousDate = False


    For r = 2 To lastRow

        v = ws.Cells(r, 1).value


        '--------------------------------------------
        ' Excelエラーの日付行はスキップ
        '--------------------------------------------

        If IsError(v) Then

            skippedRows = skippedRows + 1


        '--------------------------------------------
        ' 有効日付
        '--------------------------------------------

        ElseIf IsValidDateValue_Local(v) Then

            dateKey = CLng(Fix(CDbl(CDate(v))))


            '----------------------------------------
            ' 日付の重複・逆転を検出
            '----------------------------------------

            If hasPreviousDate Then

                If dateKey <= previousDate Then

                    Err.Raise vbObjectError + 300, , _
                        "日付が昇順ではないか、重複しています。" & _
                        vbCrLf & _
                        "シート: " & ws.Name & vbCrLf & _
                        "行: " & r

                End If

            End If


            dateCount = dateCount + 1

            dateRows(dateCount) = r

            previousDate = dateKey

            hasPreviousDate = True


        '--------------------------------------------
        ' 空白行・集計行・その他の非日付行
        '--------------------------------------------

        Else

            skippedRows = skippedRows + 1

        End If

    Next r


    If dateCount = 0 Then

        Err.Raise vbObjectError + 301, , _
            "有効な日付行が存在しません。" & _
            vbCrLf & _
            "シート: " & ws.Name

    End If


    ReDim Preserve dateRows(1 To dateCount)

End Sub


'========================================================
' 日付の有効性判定
'
' Excel日付型・日付シリアル値・日付文字列に対応。
'
'========================================================

Private Function IsValidDateValue_Local( _
    ByVal v As Variant _
) As Boolean

    Dim d As Date

    IsValidDateValue_Local = False


    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbBoolean Then
        Exit Function
    End If

    If Not IsDate(v) Then
        Exit Function
    End If


    d = CDate(v)


    If Year(d) < 1900 Then
        Exit Function
    End If


    IsValidDateValue_Local = True

End Function


'========================================================
' 1つのLookbackの共分散行列を計算
'
' 【今回の修正の中心】
'
' 直近N+1個の日付行について、
' 系列ごとに全水準値を検証する。
'
' 1つでも欠損値がある系列は、
' 当該Lookbackの計算対象から系列全体を除外する。
'
' 有効系列同士は直近N本の変化幅をすべて使用する。
'
'========================================================

Private Function BuildSingleCovMatrix_Local( _
    ByVal wsSrc As Worksheet, _
    ByVal lastCol As Long, _
    ByRef headers() As String, _
    ByRef dateRows() As Long, _
    ByVal dateCount As Long, _
    ByVal lookbackDays As Long, _
    ByVal multiplier As Double, _
    ByRef validityResult As Variant, _
    ByRef validCount As Long, _
    ByRef invalidCount As Long, _
    ByRef invalidSeriesText As String, _
    ByRef firstLevelRow As Long, _
    ByRef lastLevelRow As Long _
) As Variant

    Dim nSeries As Long

    Dim outArr() As Variant

    Dim validSeries() As Boolean

    Dim changes() As Double
    Dim means() As Double

    Dim levelsBlock As Variant

    Dim firstDateIndex As Long

    Dim firstPhysicalRow As Long
    Dim lastPhysicalRow As Long

    Dim currentRow As Long
    Dim previousRow As Long

    Dim currentOffset As Long
    Dim previousOffset As Long

    Dim currentValue As Variant
    Dim previousValue As Variant

    Dim invalidAddress As String

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim sumValue As Double
    Dim covValue As Double


    nSeries = lastCol - 1

    validCount = 0
    invalidCount = 0

    invalidSeriesText = vbNullString

    firstLevelRow = 0
    lastLevelRow = dateRows(dateCount)


    '====================================================
    ' 出力配列
    '====================================================

    ReDim outArr( _
        1 To nSeries + 1, _
        1 To nSeries + 1)

    ReDim validSeries(1 To nSeries)


    '====================================================
    ' ヘッダー
    '====================================================

    outArr(1, 1) = _
        "Cov " & CStr(lookbackDays) & "D"


    For j = 1 To nSeries

        outArr(1, j + 1) = headers(j)

        outArr(j + 1, 1) = headers(j)

    Next j


    '====================================================
    ' 必要な日付行数が不足
    '
    ' 30日Cov：31個の水準
    ' 60日Cov：61個の水準
    ' 120日Cov：121個の水準
    '
    ' 不足したLookbackだけをNAにする。
    '====================================================

    If dateCount <= lookbackDays Then

        For i = 1 To nSeries

            validSeries(i) = False

            For j = 1 To nSeries

                outArr(i + 1, j + 1) = _
                    CVErr(xlErrNA)

            Next j

        Next i


        invalidCount = nSeries

        invalidSeriesText = _
            "日付データ数不足: " & _
            CStr(dateCount) & _
            "水準 / 必要数 " & _
            CStr(lookbackDays + 1)


        validityResult = validSeries

        BuildSingleCovMatrix_Local = outArr

        Exit Function

    End If


    '====================================================
    ' 対象期間
    '====================================================

    firstDateIndex = dateCount - lookbackDays

    firstPhysicalRow = dateRows(firstDateIndex)

    lastPhysicalRow = dateRows(dateCount)

    firstLevelRow = firstPhysicalRow

    lastLevelRow = lastPhysicalRow


    '====================================================
    ' 対象期間のB列以降を一括取得
    '
    ' 途中に空白行・集計行が存在しても、
    ' dateRowsで指定した日付行だけを参照する。
    '
    ' Excelセルへのアクセスを減らし、
    ' 系列数が多い場合の処理速度を改善する。
    '====================================================

    levelsBlock = wsSrc.Range( _
        wsSrc.Cells(firstPhysicalRow, 2), _
        wsSrc.Cells(lastPhysicalRow, lastCol) _
    ).value2


    '====================================================
    ' 配列確保
    '====================================================

    ReDim changes( _
        1 To lookbackDays, _
        1 To nSeries)

    ReDim means(1 To nSeries)


    '====================================================
    ' 1. 各系列の有効性を判定
    '
    ' 直近N+1個の水準のうち、
    ' 1つでも欠損・エラー・非数値があれば、
    ' その系列全体を無効とする。
    '====================================================

    For j = 1 To nSeries

        validSeries(j) = True

        invalidAddress = vbNullString


        For k = firstDateIndex To dateCount

            currentRow = dateRows(k)

            currentOffset = _
                currentRow - firstPhysicalRow + 1


            currentValue = _
                levelsBlock(currentOffset, j)


            If Not IsValidNumericCell_Local( _
                currentValue) Then

                validSeries(j) = False

                invalidAddress = _
                    wsSrc.Cells( _
                        currentRow, j + 1).Address( _
                            False, False)

                Exit For

            End If

        Next k


        '--------------------------------------------
        ' 有効系列
        '--------------------------------------------

        If validSeries(j) Then

            validCount = validCount + 1


        '--------------------------------------------
        ' 無効系列
        '--------------------------------------------

        Else

            invalidCount = invalidCount + 1


            If Len(invalidSeriesText) > 0 Then

                invalidSeriesText = _
                    invalidSeriesText & ", "

            End If


            invalidSeriesText = _
                invalidSeriesText & _
                headers(j) & _
                " [" & invalidAddress & "]"

        End If

    Next j


    '====================================================
    ' 2. 有効系列の1日変化幅を作成
    '
    ' 欠損を含む系列はここでは計算しない。
    '
    ' 有効系列については全N本の変化幅を作成する。
    '====================================================

    For j = 1 To nSeries

        If validSeries(j) Then

            For k = 1 To lookbackDays

                previousRow = _
                    dateRows(firstDateIndex + k - 1)

                currentRow = _
                    dateRows(firstDateIndex + k)


                previousOffset = _
                    previousRow - firstPhysicalRow + 1

                currentOffset = _
                    currentRow - firstPhysicalRow + 1


                previousValue = _
                    levelsBlock(previousOffset, j)

                currentValue = _
                    levelsBlock(currentOffset, j)


                changes(k, j) = _
                    (CDbl(currentValue) - _
                     CDbl(previousValue)) * multiplier

            Next k

        End If

    Next j


    '====================================================
    ' 3. 有効系列の平均変化幅
    '
    ' N本すべての変化幅を使用する。
    '====================================================

    For j = 1 To nSeries

        If validSeries(j) Then

            sumValue = 0#


            For k = 1 To lookbackDays

                sumValue = _
                    sumValue + changes(k, j)

            Next k


            means(j) = sumValue / lookbackDays

        End If

    Next j


    '====================================================
    ' 4. 標本共分散行列
    '
    ' Cov(i,j)
    '
    ' = SUM[
    '     (dYi - Mean_i) *
    '     (dYj - Mean_j)
    '   ] / (N - 1)
    '
    ' すべての有効系列で同一のN日を使用する。
    '
    ' 無効系列に対応する行・列はすべてNA。
    '====================================================

    For i = 1 To nSeries

        For j = i To nSeries


            If validSeries(i) And validSeries(j) Then

                covValue = 0#


                For k = 1 To lookbackDays

                    covValue = _
                        covValue + _
                        (changes(k, i) - means(i)) * _
                        (changes(k, j) - means(j))

                Next k


                covValue = _
                    covValue / (lookbackDays - 1)


                outArr(i + 1, j + 1) = covValue

                outArr(j + 1, i + 1) = covValue


            Else

                outArr(i + 1, j + 1) = _
                    CVErr(xlErrNA)

                outArr(j + 1, i + 1) = _
                    CVErr(xlErrNA)

            End If

        Next j

    Next i


    '====================================================
    ' 戻り値
    '====================================================

    validityResult = validSeries

    BuildSingleCovMatrix_Local = outArr

End Function


'========================================================
' 数値として利用可能か判定
'
' False:
'   Empty
'   Null
'   Excelエラー
'   数式による空文字
'   Boolean
'   非数値文字列
'
'========================================================

Private Function IsValidNumericCell_Local( _
    ByVal cellValue As Variant _
) As Boolean

    IsValidNumericCell_Local = False


    If IsError(cellValue) Then Exit Function

    If IsNull(cellValue) Then Exit Function

    If IsEmpty(cellValue) Then Exit Function

    If VarType(cellValue) = vbBoolean Then
        Exit Function
    End If


    If Len(Trim$(CStr(cellValue))) = 0 Then
        Exit Function
    End If


    If Not IsNumeric(cellValue) Then
        Exit Function
    End If


    IsValidNumericCell_Local = True

End Function


'========================================================
' ヘッダー検証
'
' 空白・エラー・重複を検出する。
'========================================================

Private Sub ValidateHeaders_Local( _
    ByVal wsSrc As Worksheet, _
    ByVal lastCol As Long, _
    ByRef headers() As String _
)

    Dim seen As Object

    Dim i As Long

    Dim v As Variant
    Dim headerName As String


    Set seen = CreateObject("Scripting.Dictionary")

    seen.CompareMode = vbTextCompare


    ReDim headers(1 To lastCol - 1)


    For i = 2 To lastCol

        v = wsSrc.Cells(1, i).value2


        If IsError(v) Then

            Err.Raise vbObjectError + 400, , _
                "ヘッダーにExcelエラーがあります。" & _
                vbCrLf & _
                "列番号: " & i

        End If


        If IsEmpty(v) Then

            Err.Raise vbObjectError + 401, , _
                "空白ヘッダーがあります。" & _
                vbCrLf & _
                "列番号: " & i

        End If


        headerName = Trim$(CStr(v))


        If Len(headerName) = 0 Then

            Err.Raise vbObjectError + 402, , _
                "空白ヘッダーがあります。" & _
                vbCrLf & _
                "列番号: " & i

        End If


        If seen.Exists(headerName) Then

            Err.Raise vbObjectError + 403, , _
                "ヘッダー名が重複しています。" & _
                vbCrLf & _
                "ヘッダー: " & headerName

        End If


        seen.Add headerName, True

        headers(i - 1) = headerName

    Next i

End Sub


'========================================================
' メタ情報出力
'========================================================

Private Sub WriteCovMetaBlock_Local( _
    ByVal wsCov As Worksheet, _
    ByVal wsSrc As Worksheet, _
    ByVal lookbackDays As Long, _
    ByVal multiplier As Double, _
    ByRef dateRows() As Long, _
    ByVal dateCount As Long, _
    ByVal skippedRows As Long, _
    ByVal firstLevelRow As Long, _
    ByVal lastLevelRow As Long, _
    ByVal validCount As Long, _
    ByVal invalidCount As Long, _
    ByVal metaStartCol As Long _
)

    With wsCov

        '--------------------------------------------
        ' 1: Title
        '--------------------------------------------

        .Cells(1, metaStartCol).value2 = _
            "Covariance Matrix"

        .Cells(1, metaStartCol + 1).value2 = _
            lookbackDays & " Days"


        '--------------------------------------------
        ' 2: Source Sheet
        '--------------------------------------------

        .Cells(2, metaStartCol).value2 = _
            "Source Sheet"

        .Cells(2, metaStartCol + 1).value2 = _
            wsSrc.Name


        '--------------------------------------------
        ' 3: As Of Date
        '--------------------------------------------

        .Cells(3, metaStartCol).value2 = _
            "As Of Date"

        .Cells(3, metaStartCol + 1).value = _
            wsSrc.Cells( _
                dateRows(dateCount), 1).value

        .Cells(3, metaStartCol + 1).numberFormat = _
            "yyyy/mm/dd"


        '--------------------------------------------
        ' 4: Lookback Days
        '--------------------------------------------

        .Cells(4, metaStartCol).value2 = _
            "Lookback Days"

        .Cells(4, metaStartCol + 1).value2 = _
            lookbackDays


        '--------------------------------------------
        ' 5: Observation Count
        '--------------------------------------------

        .Cells(5, metaStartCol).value2 = _
            "Observation Count"

        If firstLevelRow > 0 Then

            .Cells(5, metaStartCol + 1).value2 = _
                lookbackDays

        Else

            .Cells(5, metaStartCol + 1).value2 = 0

        End If


        '--------------------------------------------
        ' 6: Source Multiplier
        '--------------------------------------------

        .Cells(6, metaStartCol).value2 = _
            "Source Multiplier"

        .Cells(6, metaStartCol + 1).value2 = _
            multiplier


        '--------------------------------------------
        ' 7: Valid Series
        '--------------------------------------------

        .Cells(7, metaStartCol).value2 = _
            "Valid Series"

        .Cells(7, metaStartCol + 1).value2 = _
            validCount


        '--------------------------------------------
        ' 8: Invalid Series
        '--------------------------------------------

        .Cells(8, metaStartCol).value2 = _
            "Invalid Series"

        .Cells(8, metaStartCol + 1).value2 = _
            invalidCount


        '--------------------------------------------
        ' 9: Change Row Range
        '--------------------------------------------

        .Cells(9, metaStartCol).value2 = _
            "Change Row Range"

        If firstLevelRow > 0 Then

            .Cells(9, metaStartCol + 1).value2 = _
                firstLevelRow & ":" & lastLevelRow

        Else

            .Cells(9, metaStartCol + 1).value2 = _
                "Insufficient data"

        End If


        '--------------------------------------------
        ' 10: Named Range
        '--------------------------------------------

        .Cells(10, metaStartCol).value2 = _
            "Named Range"

        .Cells(10, metaStartCol + 1).value2 = _
            "CovMatrix" & lookbackDays

    End With

End Sub


'========================================================
' 名前付き範囲作成
'
' CovTable30 / 60 / 120
' CovMatrix30 / 60 / 120
' CovHeaders30 / 60 / 120
' CovRowHeaders30 / 60 / 120
'
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
            outputStartCol).Resize( _
                nSeries + 1, nSeries + 1)


    Set rngMatrix = _
        wsCov.Cells( _
            outputStartRow + 1, _
            outputStartCol + 1).Resize( _
                nSeries, nSeries)


    Set rngHeaders = _
        wsCov.Cells( _
            outputStartRow, _
            outputStartCol + 1).Resize( _
                1, nSeries)


    Set rngRowHeaders = _
        wsCov.Cells( _
            outputStartRow + 1, _
            outputStartCol).Resize( _
                nSeries, 1)


    AddWorkbookRangeName_Local _
        wb, "CovTable" & suffix, rngTable

    AddWorkbookRangeName_Local _
        wb, "CovMatrix" & suffix, rngMatrix

    AddWorkbookRangeName_Local _
        wb, "CovHeaders" & suffix, rngHeaders

    AddWorkbookRangeName_Local _
        wb, "CovRowHeaders" & suffix, rngRowHeaders

End Sub


'========================================================
' Cov関連名前付き範囲を削除
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
        "CovRowHeaders")


    For i = LBound(lookbacks) To UBound(lookbacks)

        For j = LBound(prefixes) To UBound(prefixes)

            DeleteWorkbookNameIfExists_Local _
                wb, _
                CStr(prefixes(j)) & _
                CStr(lookbacks(i))

        Next j

    Next i

End Sub


'========================================================
' ブックレベルの名前付き範囲を作成
'========================================================

Private Sub AddWorkbookRangeName_Local( _
    ByVal wb As Workbook, _
    ByVal rangeName As String, _
    ByVal targetRange As Range _
)

    Dim escapedSheetName As String
    Dim refersToText As String


    escapedSheetName = Replace( _
        targetRange.Worksheet.Name, _
        "'", "''")


    refersToText = _
        "='" & escapedSheetName & "'!" & _
        targetRange.Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1)


    wb.names.Add _
        Name:=rangeName, _
        RefersTo:=refersToText, _
        Visible:=True

End Sub


'========================================================
' ブックレベルの名前付き範囲を削除
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
' 各共分散行列ブロックの書式設定
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

    Dim i As Long
    Dim c As Long


    lastOutputRow = outputStartRow + nSeries

    lastOutputCol = outputStartCol + nSeries


    With ws

        Set matrixRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(lastOutputRow, lastOutputCol))


        Set numberRange = .Range( _
            .Cells(outputStartRow + 1, outputStartCol + 1), _
            .Cells(lastOutputRow, lastOutputCol))


        Set headerRowRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(outputStartRow, lastOutputCol))


        Set headerColRange = .Range( _
            .Cells(outputStartRow, outputStartCol), _
            .Cells(lastOutputRow, outputStartCol))


        Set metaLabelRange = .Range( _
            .Cells(1, outputStartCol), _
            .Cells(10, outputStartCol))


        Set metaValueRange = .Range( _
            .Cells(1, outputStartCol + 1), _
            .Cells(10, outputStartCol + 1))


        '--------------------------------------------
        ' メタ情報
        '--------------------------------------------

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

        .Cells(1, outputStartCol + 1).Interior.Color = _
            RGB(68, 114, 196)


        .Cells(1, outputStartCol).Font.Color = _
            RGB(255, 255, 255)

        .Cells(1, outputStartCol + 1).Font.Color = _
            RGB(255, 255, 255)


        '--------------------------------------------
        ' 共分散行列
        '--------------------------------------------

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


        '--------------------------------------------
        ' 対角分散の強調
        '--------------------------------------------

        For i = 1 To nSeries

            With .Cells( _
                outputStartRow + i, _
                outputStartCol + i)

                .Interior.Color = RGB(255, 242, 204)

                .Font.Bold = True

            End With

        Next i


        '--------------------------------------------
        ' 列幅
        '--------------------------------------------

        .Columns(outputStartCol).ColumnWidth = 18


        For c = outputStartCol + 1 To lastOutputCol

            .Columns(c).ColumnWidth = 11

        Next c

    End With

End Sub


'========================================================
' Covシート共通書式
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


        '--------------------------------------------
        ' ブロック間の空白列
        '--------------------------------------------

        For idx = 0 To 1

            gapStartCol = _
                1 + idx * _
                (matrixBlockWidth + gapColumns) + _
                matrixBlockWidth


            For c = gapStartCol To _
                gapStartCol + gapColumns - 1

                .Columns(c).ColumnWidth = 3

            Next c

        Next idx


        '--------------------------------------------
        ' ウィンドウ枠固定
        '--------------------------------------------

        On Error Resume Next

        .Activate

        ActiveWindow.FreezePanes = False

        .Cells(outputStartRow + 1, 2).Select

        ActiveWindow.FreezePanes = True

        On Error GoTo 0

    End With

End Sub


'========================================================
' 無効系列の行・列ヘッダーを強調
'
' validityResult：
'   True  = 有効系列
'   False = 無効系列
'
'========================================================

Private Sub MarkInvalidSeries_Local( _
    ByVal ws As Worksheet, _
    ByRef validityResult As Variant, _
    ByVal nSeries As Long, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long _
)

    Dim i As Long


    For i = 1 To nSeries

        If Not CBool(validityResult(i)) Then


            '----------------------------------------
            ' 列ヘッダー
            '----------------------------------------

            With ws.Cells( _
                outputStartRow, _
                outputStartCol + i)

                .Interior.Color = RGB(244, 204, 204)

                .Font.Color = RGB(156, 0, 6)

                .Font.Bold = True

            End With


            '----------------------------------------
            ' 行ヘッダー
            '----------------------------------------

            With ws.Cells( _
                outputStartRow + i, _
                outputStartCol)

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
' シート取得または作成
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
            After:=wb.Worksheets(wb.Worksheets.count))

        ws.Name = sheetName

    End If


    Set GetOrCreateSheet_Local = ws

End Function


'========================================================
' Excel状態復元
'========================================================

Private Sub RestoreExcelState_Local( _
    ByVal oldCalc As XlCalculation, _
    ByVal oldScreenUpdating As Boolean, _
    ByVal oldEnableEvents As Boolean, _
    ByVal oldDisplayAlerts As Boolean, _
    ByVal stateCaptured As Boolean _
)

    If Not stateCaptured Then
        Exit Sub
    End If


    Application.Calculation = oldCalc

    Application.EnableEvents = oldEnableEvents

    Application.DisplayAlerts = oldDisplayAlerts

    Application.ScreenUpdating = oldScreenUpdating

End Sub


