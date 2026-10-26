
Option Explicit

'========================================================
' RISKSTATS - USER SETTINGS
'========================================================

Private Const OUTPUT_SHEET As String = "RiskStats"

' 変化幅
Private Const CHANGE_1_N As Long = 1
Private Const CHANGE_5_N As Long = 5
Private Const CHANGE_25_N As Long = 25

' Z-score
Private Const ZSCORE_SHORT_N As Long = 5
Private Const ZSCORE_LONG_N As Long = 25

' High / Low
Private Const HIGHLOW_SHORT_N As Long = 5
Private Const HIGHLOW_LONG_N As Long = 25

' Volatility
Private Const VOLATILITY_N As Long = 25

' Regression
Private Const REGRESSION_N As Long = 25

' InputBox初期値
' Sourceが空欄ならアクティブシート名を使用
' Explanatory Sheetが空欄なら最初のSourceを使用
Private Const DEFAULT_SOURCE_SHEETS As String = "JGBData, OISData"
Private Const DEFAULT_OUTPUT_MULTIPLIER As String = "100"
Private Const DEFAULT_EXPLANATORY_SHEET As String = "JGBData"
Private Const DEFAULT_EXPLANATORY_SERIES As String = "yield_10y"

' A:R
Private Const OUTPUT_COLS As Long = 18


'========================================================
' MAIN
'========================================================

Public Sub BuildRiskStats()

    Dim sourceInput As String
    Dim normalizedInput As String

    Dim sourceDefault As String
    Dim explanatoryDefault As String

    Dim sheetItems As Variant
    Dim item As Variant

    Dim multiplierInput As String
    Dim outputMultiplier As Double

    Dim explanatorySheetName As String
    Dim explanatoryName As String
    Dim explanatoryCol As Long

    Dim wsSrc As Worksheet
    Dim wsExp As Worksheet
    Dim wsOut As Worksheet

    Dim sourceSheets As Collection
    Dim seenSheets As Object

    Dim sheetName As String

    Dim expBlock As Variant
    Dim expDates() As Long
    Dim expSeries As Variant

    Dim srcBlock As Variant
    Dim srcDates() As Long
    Dim srcSeries As Variant

    Dim expLastCol As Long
    Dim lastCol As Long

    Dim expCount As Long
    Dim seriesCount As Long

    Dim expSkipped As Long
    Dim skipped As Long

    Dim xChangeMap As Object
    Dim results() As Variant

    Dim totalSeries As Long
    Dim outputIndex As Long

    Dim sheetIndex As Long
    Dim c As Long

    Dim betaValue As Variant
    Dim r2Value As Variant
    Dim tValue As Variant

    Dim tradeName As String
    Dim sourceSummary As String

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldAutoFill As Boolean

    Dim stateChanged As Boolean
    Dim autoFillCaptured As Boolean

    Dim currentStep As String
    Dim errNumber As Long
    Dim errText As String

    On Error GoTo ErrHandler

    '====================================================
    ' 1. Settings validation
    '====================================================

    currentStep = "設定値検証"

    ValidateRiskStatsSettings

    '====================================================
    ' 2. Source sheets - InputBox
    '====================================================

    currentStep = "計算元シート入力"

    sourceDefault = Trim$(DEFAULT_SOURCE_SHEETS)

    If Len(sourceDefault) = 0 Then
        sourceDefault = ActiveSheet.Name
    End If

    sourceInput = InputBox( _
        "計算元シートを指定してください。" & vbCrLf & _
        "複数指定はカンマ区切りで入力してください。" & _
        vbCrLf & vbCrLf & _
        "例：JGBData,OISData", _
        "RiskStats - 計算元シート", _
        sourceDefault)

    If Len(Trim$(sourceInput)) = 0 Then
        Exit Sub
    End If

    normalizedInput = sourceInput

    normalizedInput = Replace(normalizedInput, "，", ",")
    normalizedInput = Replace(normalizedInput, "、", ",")
    normalizedInput = Replace(normalizedInput, ";", ",")
    normalizedInput = Replace(normalizedInput, "；", ",")

    sheetItems = Split(normalizedInput, ",")

    '====================================================
    ' 3. Output multiplier - InputBox
    '====================================================

    currentStep = "出力倍率入力"

    multiplierInput = InputBox( _
        "出力倍率を入力してください。" & vbCrLf & _
        "1 = 元データの単位" & vbCrLf & _
        "100 = %表記からbpに変換", _
        "RiskStats - 出力倍率", _
        DEFAULT_OUTPUT_MULTIPLIER)

    If Len(Trim$(multiplierInput)) = 0 Then
        Exit Sub
    End If

    If Not IsNumeric(multiplierInput) Then

        Err.Raise vbObjectError + 101, , _
            "出力倍率は数値で入力してください。"

    End If

    outputMultiplier = CDbl(multiplierInput)

    If outputMultiplier <= 0# Then

        Err.Raise vbObjectError + 102, , _
            "出力倍率は0より大きい値を指定してください。"

    End If

    '====================================================
    ' 4. Explanatory sheet - InputBox
    '====================================================

    currentStep = "説明変数シート入力"

    explanatoryDefault = Trim$(DEFAULT_EXPLANATORY_SHEET)

    If Len(explanatoryDefault) = 0 Then

        explanatoryDefault = _
            Trim$(CStr(sheetItems(LBound(sheetItems))))

    End If

    explanatorySheetName = InputBox( _
        "回帰説明変数が存在するシート名を入力してください。", _
        "RiskStats - 説明変数シート", _
        explanatoryDefault)

    If Len(Trim$(explanatorySheetName)) = 0 Then
        Exit Sub
    End If

    explanatorySheetName = Trim$(explanatorySheetName)

    If StrComp( _
        explanatorySheetName, OUTPUT_SHEET, _
        vbTextCompare) = 0 Then

        Err.Raise vbObjectError + 103, , _
            "RiskStatsを説明変数シートには指定できません。"

    End If

    '====================================================
    ' 5. Explanatory series - InputBox
    '====================================================

    currentStep = "説明変数系列入力"

    explanatoryName = InputBox( _
        "回帰説明変数のヘッダー名を入力してください。", _
        "RiskStats - 説明変数系列", _
        DEFAULT_EXPLANATORY_SERIES)

    If Len(Trim$(explanatoryName)) = 0 Then
        Exit Sub
    End If

    explanatoryName = Trim$(explanatoryName)

    '====================================================
    ' 6. Validate source sheets
    '====================================================

    currentStep = "計算元シート検証"

    Set sourceSheets = New Collection

    Set seenSheets = CreateObject("Scripting.Dictionary")
    seenSheets.CompareMode = vbTextCompare

    totalSeries = 0

    For Each item In sheetItems

        sheetName = Trim$(CStr(item))

        If Len(sheetName) = 0 Then

            Err.Raise vbObjectError + 104, , _
                "計算元シート名に空欄があります。"

        End If

        If StrComp( _
            sheetName, OUTPUT_SHEET, _
            vbTextCompare) = 0 Then

            Err.Raise vbObjectError + 105, , _
                "RiskStatsを計算元シートには指定できません。"

        End If

        If seenSheets.Exists(sheetName) Then

            Err.Raise vbObjectError + 106, , _
                "計算元シートが重複しています: " & sheetName

        End If

        seenSheets.Add sheetName, True

        Set wsSrc = GetWorksheetOrError(sheetName)

        lastCol = wsSrc.Cells( _
            1, wsSrc.Columns.Count).End(xlToLeft).Column

        If lastCol < 2 Then

            Err.Raise vbObjectError + 107, , _
                "B列以降に系列がありません: " & sheetName

        End If

        For c = 2 To lastCol
            ValidateHeader wsSrc, c
        Next c

        sourceSheets.Add wsSrc

        totalSeries = totalSeries + lastCol - 1

        If Len(sourceSummary) > 0 Then
            sourceSummary = sourceSummary & ", "
        End If

        sourceSummary = sourceSummary & wsSrc.Name

    Next item

    If totalSeries < 1 Then

        Err.Raise vbObjectError + 108, , _
            "計算対象系列が存在しません。"

    End If

    '====================================================
    ' 7. Read explanatory data
    '====================================================

    currentStep = "説明変数データ取得"

    Set wsExp = GetWorksheetOrError(explanatorySheetName)

    expLastCol = wsExp.Cells( _
        1, wsExp.Columns.Count).End(xlToLeft).Column

    explanatoryCol = FindHeaderColumnOrError( _
        wsExp, explanatoryName, expLastCol)

    ReadSheetData _
        wsExp, expBlock, expDates, expLastCol

    BuildSeriesData _
        expBlock, expDates, explanatoryCol, _
        expSeries, expCount, expSkipped

    Set xChangeMap = CreateObject("Scripting.Dictionary")

    BuildChangeMap _
        expSeries, expCount, xChangeMap

    '====================================================
    ' 8. Calculate all series in memory
    '
    ' 出力シートにはまだ触れない。
    '====================================================

    currentStep = "RiskStats計算"

    ReDim results(1 To totalSeries, 1 To OUTPUT_COLS)

    outputIndex = 0

    For sheetIndex = 1 To sourceSheets.Count

        Set wsSrc = sourceSheets(sheetIndex)

        ReadSheetData _
            wsSrc, srcBlock, srcDates, lastCol

        For c = 2 To lastCol

            outputIndex = outputIndex + 1

            BuildSeriesData _
                srcBlock, srcDates, c, _
                srcSeries, seriesCount, skipped

            tradeName = CStr(wsSrc.Cells(1, c).value2)

            ' A: Sheet
            results(outputIndex, 1) = wsSrc.Name

            ' B: Trade
            results(outputIndex, 2) = tradeName

            ' C: Level
            results(outputIndex, 3) = _
                LatestLevel( _
                    srcSeries, seriesCount, outputMultiplier)

            ' D: 1d-Chg
            results(outputIndex, 4) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    CHANGE_1_N, outputMultiplier)

            ' E: 5d-Chg
            results(outputIndex, 5) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    CHANGE_5_N, outputMultiplier)

            ' F: 25d-Chg
            results(outputIndex, 6) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    CHANGE_25_N, outputMultiplier)

            ' G: 5d-Zscore
            results(outputIndex, 7) = _
                ZScoreLevel( _
                    srcSeries, seriesCount, ZSCORE_SHORT_N)

            ' H: 5d-High
            results(outputIndex, 8) = _
                HighLowLevel( _
                    srcSeries, seriesCount, _
                    HIGHLOW_SHORT_N, outputMultiplier, True)

            ' I: 5d-Low
            results(outputIndex, 9) = _
                HighLowLevel( _
                    srcSeries, seriesCount, _
                    HIGHLOW_SHORT_N, outputMultiplier, False)

            ' J: 25d-Zscore
            results(outputIndex, 10) = _
                ZScoreLevel( _
                    srcSeries, seriesCount, ZSCORE_LONG_N)

            ' K: 25d-High
            results(outputIndex, 11) = _
                HighLowLevel( _
                    srcSeries, seriesCount, _
                    HIGHLOW_LONG_N, outputMultiplier, True)

            ' L: 25d-Low
            results(outputIndex, 12) = _
                HighLowLevel( _
                    srcSeries, seriesCount, _
                    HIGHLOW_LONG_N, outputMultiplier, False)

            ' M: 25d-vol
            results(outputIndex, 13) = _
                VolDailyChange( _
                    srcSeries, seriesCount, _
                    VOLATILITY_N, outputMultiplier)

            ' N:P Regression
            Regression25DailyChangeCrossSheet _
                srcSeries, seriesCount, _
                xChangeMap, _
                betaValue, r2Value, tValue

            results(outputIndex, 14) = betaValue
            results(outputIndex, 15) = r2Value
            results(outputIndex, 16) = tValue

            ' Q: Latest
            If seriesCount > 0 Then

                results(outputIndex, 17) = _
                    CDbl(srcSeries(1, seriesCount))

            Else

                results(outputIndex, 17) = "NA"

            End If

            ' R: Type
            results(outputIndex, 18) = _
                ClassifyRiskStatsTradeType(tradeName)

        Next c

    Next sheetIndex

    If outputIndex <> totalSeries Then

        Err.Raise vbObjectError + 110, , _
            "計算結果の行数が一致しません。"

    End If

    '====================================================
    ' 9. Preserve Excel application settings
    '====================================================

    currentStep = "Excel設定退避"

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    stateChanged = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Excelテーブルの計算列自動補完を一時停止
    On Error Resume Next

    Err.Clear

    oldAutoFill = _
        Application.AutoCorrect.AutoFillFormulasInLists

    autoFillCaptured = (Err.Number = 0)

    If autoFillCaptured Then

        Application.AutoCorrect.AutoFillFormulasInLists = False

    End If

    On Error GoTo ErrHandler

    '====================================================
    ' 10. Output
    '
    ' フィルター条件をクリアしてから更新。
    '====================================================

    currentStep = "RiskStats出力"

    Set wsOut = GetOrCreateWorksheet(OUTPUT_SHEET)

    WriteRiskStatsSafely _
        wsOut, results, totalSeries

    '====================================================
    ' 11. Restore Excel settings
    '====================================================

    If autoFillCaptured Then

        Application.AutoCorrect.AutoFillFormulasInLists = _
            oldAutoFill

    End If

    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    stateChanged = False

    MsgBox _
        "RiskStatsの更新が完了しました。" & vbCrLf & _
        "計算元: " & sourceSummary & vbCrLf & _
        "系列数: " & totalSeries & vbCrLf & _
        "説明変数: " & _
        wsExp.Name & "!" & explanatoryName & vbCrLf & _
        "フィルター条件: クリア" & vbCrLf & _
        "データ照合: PASS", _
        vbInformation

    Exit Sub

ErrHandler:

    errNumber = Err.Number
    errText = Err.Description

    On Error Resume Next

    If stateChanged Then

        If autoFillCaptured Then

            Application.AutoCorrect.AutoFillFormulasInLists = _
                oldAutoFill

        End If

        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating

    End If

    On Error GoTo 0

    MsgBox _
        "RiskStats更新エラー" & vbCrLf & _
        "処理: " & currentStep & vbCrLf & _
        "Err.Number: " & CStr(errNumber) & vbCrLf & _
        errText, _
        vbCritical

End Sub


'========================================================
' CLEAR FILTER CRITERIA
'
' 今回の主要修正
'
' ・Excelテーブルのフィルターをクリア
' ・通常のAutoFilterもクリア
' ・フィルターの▼は残す
' ・抽出条件は復元しない
' ・並べ替え条件は変更しない
'========================================================

Private Sub ClearRiskStatsFilters( _
    ByVal ws As Worksheet _
)

    Dim tbl As ListObject

    '====================================================
    ' 1. Excelテーブル
    '====================================================

    For Each tbl In ws.ListObjects

        If tbl.ShowAutoFilter Then

            If tbl.AutoFilter.FilterMode Then

                tbl.AutoFilter.ShowAllData

            End If

        End If

    Next tbl

    '====================================================
    ' 2. 通常のシートAutoFilter
    '====================================================

    If ws.FilterMode Then

        ws.ShowAllData

    End If

    '====================================================
    ' 3. フィルター解除を確認
    '====================================================

    For Each tbl In ws.ListObjects

        If tbl.ShowAutoFilter Then

            If tbl.AutoFilter.FilterMode Then

                Err.Raise vbObjectError + 700, , _
                    "テーブルのフィルター条件を" & _
                    "クリアできませんでした: " & tbl.Name

            End If

        End If

    Next tbl

    If ws.FilterMode Then

        Err.Raise vbObjectError + 701, , _
            "ワークシートのフィルター条件を" & _
            "クリアできませんでした。"

    End If

End Sub


'========================================================
' SAFE OUTPUT
'
' 通常範囲とExcelテーブルの両方に対応。
'
' 重要：
'   フィルターを先にクリアする。
'   データ行全体を更新する。
'   NumberFormatを変更しない。
'   AutoFilterModeをFalseにしない。
'   フィルター条件を復元しない。
'========================================================

Private Sub WriteRiskStatsSafely( _
    ByVal ws As Worksheet, _
    ByRef results() As Variant, _
    ByVal resultCount As Long _
)

    Dim tbl As ListObject
    Dim candidate As ListObject

    Dim oldLastRow As Long
    Dim clearLastRow As Long
    Dim targetLastRow As Long

    Dim outputRange As Range
    Dim conflictingRange As Range

    If resultCount < 1 Then

        Err.Raise vbObjectError + 702, , _
            "出力対象となる系列がありません。"

    End If

    If ws.ProtectContents Then

        Err.Raise vbObjectError + 703, , _
            "RiskStatsシートが保護されています。"

    End If

    targetLastRow = resultCount + 1

    '====================================================
    ' 1. Filter criteriaをクリア
    '
    ' これにより全行が表示される。
    '====================================================

    ClearRiskStatsFilters ws

    '====================================================
    ' 2. Last used row
    '====================================================

    oldLastRow = RiskStatsLastUsedRow(ws)

    clearLastRow = Application.Max( _
        oldLastRow, targetLastRow)

    '====================================================
    ' 3. Excelテーブルの有無を判定
    '====================================================

    Set tbl = Nothing

    For Each candidate In ws.ListObjects

        If candidate.Range.Row = 1 And _
           candidate.Range.Column = 1 Then

            Set tbl = candidate

            Exit For

        End If

    Next candidate

    '====================================================
    ' 4. 他のテーブルとの重複を確認
    '
    ' RiskStats出力範囲に別のテーブルがある場合、
    ' 他データを破壊しないよう中止する。
    '====================================================

    Set outputRange = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(clearLastRow, OUTPUT_COLS))

    For Each candidate In ws.ListObjects

        If tbl Is Nothing Then

            Set conflictingRange = _
                Application.Intersect( _
                    candidate.Range, outputRange)

            If Not conflictingRange Is Nothing Then

                Err.Raise vbObjectError + 704, , _
                    "RiskStats出力領域に別のExcelテーブルがあります: " & _
                    candidate.Name

            End If

        ElseIf Not candidate Is tbl Then

            Set conflictingRange = _
                Application.Intersect( _
                    candidate.Range, outputRange)

            If Not conflictingRange Is Nothing Then

                Err.Raise vbObjectError + 705, , _
                    "RiskStats出力領域に別のExcelテーブルがあります: " & _
                    candidate.Name

            End If

        End If

        Set conflictingRange = Nothing

    Next candidate

    '====================================================
    ' 5A. 通常のセル範囲
    '====================================================

    If tbl Is Nothing Then

        '--------------------------------------------
        ' ヘッダーを残し、2行目以降を削除
        '--------------------------------------------

        ws.Range( _
            ws.Cells(2, 1), _
            ws.Cells(clearLastRow, OUTPUT_COLS) _
        ).ClearContents

        '--------------------------------------------
        ' ヘッダーを更新
        '--------------------------------------------

        WriteHeaders ws

        '--------------------------------------------
        ' 計算結果を全行へ書き込み
        '--------------------------------------------

        ws.Cells(2, 1).Resize( _
            resultCount, OUTPUT_COLS).value2 = results

    Else

        '================================================
        ' 5B. Excelテーブル
        '================================================

        If Not tbl.ShowHeaders Then

            Err.Raise vbObjectError + 706, , _
                "RiskStatsテーブルのヘッダーが非表示です。"

        End If

        If tbl.ShowTotals Then

            Err.Raise vbObjectError + 707, , _
                "RiskStatsテーブルに集計行があります。" & _
                "集計行を保護するため更新を中止しました。"

        End If

        If tbl.ListColumns.Count > OUTPUT_COLS Then

            Err.Raise vbObjectError + 708, , _
                "RiskStatsテーブルが18列を超えています。" & _
                "追加列を保護するため更新を中止しました。"

        End If

        '--------------------------------------------
        ' 既存テーブルの全データを消去
        '
        ' フィルターはすでにクリア済み。
        '--------------------------------------------

        If Not tbl.DataBodyRange Is Nothing Then

            tbl.DataBodyRange.ClearContents

        End If

        '--------------------------------------------
        ' テーブルのサイズを計算結果に合わせる
        '
        ' A1:R(resultCount+1)
        '
        ' フィルター機能は維持する。
        '--------------------------------------------

        If tbl.Range.rows.Count <> resultCount + 1 Or _
           tbl.Range.Columns.Count <> OUTPUT_COLS Then

            tbl.Resize ws.Range( _
                ws.Cells(1, 1), _
                ws.Cells(targetLastRow, OUTPUT_COLS))

        End If

        '--------------------------------------------
        ' Resize後のデータ領域を確認
        '--------------------------------------------

        If tbl.DataBodyRange Is Nothing Then

            Err.Raise vbObjectError + 709, , _
                "テーブルのデータ領域を取得できません。"

        End If

        If tbl.DataBodyRange.rows.Count <> resultCount Or _
           tbl.DataBodyRange.Columns.Count <> OUTPUT_COLS Then

            Err.Raise vbObjectError + 710, , _
                "テーブルのサイズが計算結果と一致しません。"

        End If

        '--------------------------------------------
        ' Resizeで生成された既定値や数式も削除
        '--------------------------------------------

        tbl.DataBodyRange.ClearContents

        '--------------------------------------------
        ' ヘッダー更新
        '--------------------------------------------

        WriteHeaders ws

        '--------------------------------------------
        ' 計算結果をテーブル全体に一括書き込み
        '--------------------------------------------

        tbl.DataBodyRange.value2 = results

    End If

    '====================================================
    ' 6. 古い結果が残った行を消去
    '
    ' 新しい系列数が減った場合に必要。
    '====================================================

    If clearLastRow > targetLastRow Then

        ws.Range( _
            ws.Cells(targetLastRow + 1, 1), _
            ws.Cells(clearLastRow, OUTPUT_COLS) _
        ).ClearContents

    End If

    '====================================================
    ' 7. 書き込み結果を検証
    '
    ' フィルター条件は復元しない。
    '====================================================

    VerifyRiskStatsOutput _
        ws, results, resultCount

End Sub


'========================================================
' LAST USED ROW
'
' A列だけでなくA:R全列から取得する。
'========================================================

Private Function RiskStatsLastUsedRow( _
    ByVal ws As Worksheet _
) As Long

    Dim searchRange As Range
    Dim foundCell As Range

    Set searchRange = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(ws.rows.Count, OUTPUT_COLS))

    Set foundCell = searchRange.Find( _
        What:="*", _
        After:=searchRange.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False)

    If foundCell Is Nothing Then

        RiskStatsLastUsedRow = 1

    Else

        RiskStatsLastUsedRow = foundCell.Row

    End If

End Function


'========================================================
' VERIFY OUTPUT
'
' 1. まず各セルを元配列と比較
' 2. 不一致の場合は行順に依存しない比較
' 3. レコードの内容が異なればエラー
'
' Excelテーブルの並べ替えで行順が変わっても、
' 全レコードが一致すればPASSとする。
'========================================================

Private Sub VerifyRiskStatsOutput( _
    ByVal ws As Worksheet, _
    ByRef expected() As Variant, _
    ByVal n As Long _
)

    Dim actual As Variant

    Dim i As Long
    Dim j As Long

    Dim mismatchRow As Long
    Dim mismatchCol As Long

    Dim positionMatches As Boolean
    Dim recordMatches As Boolean

    Dim expectedCounts As Object
    Dim rowKey As String
    Dim key As Variant

    Dim actualText As String
    Dim expectedText As String

    Dim formulaCells As Range
    Dim sourceRange As Range

    Set sourceRange = ws.Cells(2, 1).Resize( _
        n, OUTPUT_COLS)

    '====================================================
    ' 数式が残っていないか確認
    '====================================================

    Set formulaCells = Nothing

    On Error Resume Next

    Set formulaCells = sourceRange.SpecialCells( _
        xlCellTypeFormulas)

    On Error GoTo 0

    If Not formulaCells Is Nothing Then

        Err.Raise vbObjectError + 720, , _
            "RiskStats出力領域に数式が残っています。" & _
            vbCrLf & _
            "Cell: " & _
            formulaCells.Cells(1, 1).Address(False, False)

    End If

    actual = sourceRange.value2

    '====================================================
    ' 1. 行・列位置で比較
    '====================================================

    positionMatches = True

    For i = 1 To n

        For j = 1 To OUTPUT_COLS

            If Not RiskStatsValuesEqual( _
                expected(i, j), actual(i, j)) Then

                positionMatches = False

                mismatchRow = i
                mismatchCol = j

                Exit For

            End If

        Next j

        If Not positionMatches Then Exit For

    Next i

    If positionMatches Then Exit Sub

    '====================================================
    ' 2. レコード集合で比較
    '
    ' 同じレコードが複数存在しても正しく照合する。
    '====================================================

    Set expectedCounts = _
        CreateObject("Scripting.Dictionary")

    expectedCounts.CompareMode = vbBinaryCompare

    For i = 1 To n

        rowKey = RiskStatsRowSignature(expected, i)

        If expectedCounts.Exists(rowKey) Then

            expectedCounts(rowKey) = _
                CLng(expectedCounts(rowKey)) + 1

        Else

            expectedCounts.Add rowKey, 1

        End If

    Next i

    recordMatches = True

    For i = 1 To n

        rowKey = RiskStatsRowSignature(actual, i)

        If Not expectedCounts.Exists(rowKey) Then

            recordMatches = False
            Exit For

        End If

        If CLng(expectedCounts(rowKey)) <= 0 Then

            recordMatches = False
            Exit For

        End If

        expectedCounts(rowKey) = _
            CLng(expectedCounts(rowKey)) - 1

    Next i

    If recordMatches Then

        For Each key In expectedCounts.keys

            If CLng(expectedCounts(key)) <> 0 Then

                recordMatches = False
                Exit For

            End If

        Next key

    End If

    If recordMatches Then Exit Sub

    '====================================================
    ' 3. 本当のデータ不一致
    '====================================================

    expectedText = RiskStatsDisplayValue( _
        expected(mismatchRow, mismatchCol))

    actualText = RiskStatsDisplayValue( _
        actual(mismatchRow, mismatchCol))

    Err.Raise vbObjectError + 721, , _
        "出力内容が計算結果と一致しません。" & _
        vbCrLf & _
        "最初の位置不一致: " & _
        ws.Cells( _
            mismatchRow + 1, _
            mismatchCol).Address(False, False) & _
        vbCrLf & _
        "期待値: " & expectedText & _
        vbCrLf & _
        "実際値: " & actualText & _
        vbCrLf & _
        "全レコード照合も不一致です。"

End Sub


'========================================================
' VALUE COMPARISON
'========================================================

Private Function RiskStatsValuesEqual( _
    ByVal expected As Variant, _
    ByVal actual As Variant _
) As Boolean

    Dim x As Double
    Dim y As Double
    Dim tolerance As Double

    RiskStatsValuesEqual = False

    If IsError(expected) Or IsError(actual) Then
        Exit Function
    End If

    If VarType(expected) = vbString Then

        If VarType(actual) <> vbString Then
            Exit Function
        End If

        RiskStatsValuesEqual = _
            (StrComp( _
                CStr(expected), _
                CStr(actual), _
                vbBinaryCompare) = 0)

        Exit Function

    End If

    If Not IsNumeric(actual) Then Exit Function

    x = CDbl(expected)
    y = CDbl(actual)

    tolerance = 0.0000000001 * _
        Application.Max(1#, Abs(x))

    RiskStatsValuesEqual = _
        (Abs(x - y) <= tolerance)

End Function


'========================================================
' ROW SIGNATURE
'========================================================

Private Function RiskStatsRowSignature( _
    ByRef arr As Variant, _
    ByVal rowIndex As Long _
) As String

    Dim j As Long
    Dim token As String
    Dim signature As String

    signature = vbNullString

    For j = 1 To OUTPUT_COLS

        token = RiskStatsValueSignature( _
            arr(rowIndex, j))

        signature = signature & _
            CStr(Len(token)) & ":" & token & "|"

    Next j

    RiskStatsRowSignature = signature

End Function


'========================================================
' VALUE SIGNATURE
'========================================================

Private Function RiskStatsValueSignature( _
    ByVal v As Variant _
) As String

    If IsError(v) Then

        RiskStatsValueSignature = "ERROR"
        Exit Function

    End If

    If IsNull(v) Then

        RiskStatsValueSignature = "NULL"
        Exit Function

    End If

    If IsEmpty(v) Then

        RiskStatsValueSignature = "EMPTY"
        Exit Function

    End If

    If VarType(v) = vbString Then

        RiskStatsValueSignature = "S:" & CStr(v)

        Exit Function

    End If

    If IsNumeric(v) Then

        RiskStatsValueSignature = _
            "N:" & Format$( _
                CDbl(v), _
                "0.00000000000000E+000")

        Exit Function

    End If

    RiskStatsValueSignature = "OTHER:" & CStr(v)

End Function


'========================================================
' DISPLAY VALUE FOR ERROR
'========================================================

Private Function RiskStatsDisplayValue( _
    ByVal v As Variant _
) As String

    If IsError(v) Then

        RiskStatsDisplayValue = "[Excel Error]"

    ElseIf IsNull(v) Then

        RiskStatsDisplayValue = "[Null]"

    ElseIf IsEmpty(v) Then

        RiskStatsDisplayValue = "[Empty]"

    Else

        RiskStatsDisplayValue = CStr(v)

    End If

    If Len(RiskStatsDisplayValue) > 100 Then

        RiskStatsDisplayValue = _
            Left$(RiskStatsDisplayValue, 100) & "..."

    End If

End Function


'========================================================
' TRADE TYPE CLASSIFICATION
'
' Outright:
'   10y, 1m, yield_10y, JN123
'
' Curve:
'   1s2s, 3m6m, 3m1s
'
' Fly:
'   1m2m3m, 5s7s10s
'========================================================

Private Function ClassifyRiskStatsTradeType( _
    ByVal tradeName As String _
) As String

    Dim s As String
    Dim re As Object
    Dim tickerRe As Object
    Dim matches As Object

    On Error GoTo Fail

    ClassifyRiskStatsTradeType = "NA"

    s = LCase$(Trim$(tradeName))

    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    If Len(s) = 0 Then Exit Function

    ' Prefix
    If Left$(s, 6) = "yield_" Then

        s = mid$(s, 7)

    ElseIf Left$(s, 5) = "rate_" Then

        s = mid$(s, 6)

    ElseIf Left$(s, 4) = "jgb_" Then

        s = mid$(s, 5)

    ElseIf Left$(s, 3) = "jgb" Then

        s = mid$(s, 4)

    End If

    If Len(s) = 0 Then Exit Function

    ' JGB銘柄識別子
    Set tickerRe = CreateObject("VBScript.RegExp")

    tickerRe.Global = False
    tickerRe.IgnoreCase = True

    tickerRe.pattern = _
        "^(jn|js|jb|jl|jx|ju)[0-9]+$"

    If tickerRe.Test(s) Then

        ClassifyRiskStatsTradeType = "Outright"
        Exit Function

    End If

    ' Tenor pattern
    Set re = CreateObject("VBScript.RegExp")

    re.Global = False
    re.IgnoreCase = True

    re.pattern = _
        "^[0-9]+(\.[0-9]+)?[mys]" & _
        "([/_-]?[0-9]+(\.[0-9]+)?[mys]){0,2}$"

    If Not re.Test(s) Then Exit Function

    re.Global = True
    re.pattern = "[0-9]+(\.[0-9]+)?[mys]"

    Set matches = re.Execute(s)

    Select Case matches.Count

        Case 1
            ClassifyRiskStatsTradeType = "Outright"

        Case 2
            ClassifyRiskStatsTradeType = "Curve"

        Case 3
            ClassifyRiskStatsTradeType = "Fly"

    End Select

    Exit Function

Fail:

    ClassifyRiskStatsTradeType = "NA"

End Function


'========================================================
' SETTINGS VALIDATION
'========================================================

Private Sub ValidateRiskStatsSettings()

    If CHANGE_1_N < 1 Or _
       CHANGE_5_N < 1 Or _
       CHANGE_25_N < 1 Then

        Err.Raise vbObjectError + 200, , _
            "Change期間は1以上にしてください。"

    End If

    If ZSCORE_SHORT_N < 2 Or _
       ZSCORE_LONG_N < 2 Then

        Err.Raise vbObjectError + 201, , _
            "Z-score期間は2以上にしてください。"

    End If

    If HIGHLOW_SHORT_N < 1 Or _
       HIGHLOW_LONG_N < 1 Then

        Err.Raise vbObjectError + 202, , _
            "High/Low期間は1以上にしてください。"

    End If

    If VOLATILITY_N < 2 Then

        Err.Raise vbObjectError + 203, , _
            "Volatility期間は2以上にしてください。"

    End If

    If REGRESSION_N < 3 Then

        Err.Raise vbObjectError + 204, , _
            "Regression期間は3以上にしてください。"

    End If

End Sub


'========================================================
' READ SOURCE SHEET
'
' dates(r):
'   >0 : 有効日付
'   =0 : 空白行・集計行
'
' 末尾集計行への対応を維持。
'========================================================

Private Sub ReadSheetData( _
    ByVal ws As Worksheet, _
    ByRef dataBlock As Variant, _
    ByRef dates() As Long, _
    ByRef lastCol As Long _
)

    Dim lastRow As Long
    Dim r As Long

    Dim dayKey As Long
    Dim previousDay As Long

    Dim hasPreviousDate As Boolean
    Dim hasSummaryStarted As Boolean

    Dim dateValue As Variant

    lastCol = ws.Cells( _
        1, ws.Columns.Count).End(xlToLeft).Column

    lastRow = ws.Cells( _
        ws.rows.Count, 1).End(xlUp).Row

    If lastCol < 2 Then

        Err.Raise vbObjectError + 300, , _
            "データ系列がありません: " & ws.Name

    End If

    If lastRow < 2 Then

        Err.Raise vbObjectError + 301, , _
            "時系列データがありません: " & ws.Name

    End If

    If IsError(ws.Cells(1, 1).value2) Then

        Err.Raise vbObjectError + 302, , _
            "日付ヘッダーがエラーです: " & ws.Name

    End If

    dataBlock = ws.Range( _
        ws.Cells(2, 1), _
        ws.Cells(lastRow, lastCol) _
    ).value2

    ReDim dates(1 To UBound(dataBlock, 1))

    previousDay = 0
    hasPreviousDate = False
    hasSummaryStarted = False

    For r = 1 To UBound(dataBlock, 1)

        dateValue = dataBlock(r, 1)

        If IsError(dateValue) Then

            Err.Raise vbObjectError + 303, , _
                "日付列にエラーがあります: " & _
                ws.Name & " row " & CStr(r + 1)

        End If

        dayKey = ParseDateKey(dateValue)

        If dayKey > 0 Then

            If hasSummaryStarted Then

                Err.Raise vbObjectError + 304, , _
                    "集計行の後に日付があります: " & _
                    ws.Name & " row " & CStr(r + 1)

            End If

            If hasPreviousDate Then

                If dayKey <= previousDay Then

                    Err.Raise vbObjectError + 305, , _
                        "日付が昇順ではありません: " & _
                        ws.Name & " row " & CStr(r + 1)

                End If

            End If

            dates(r) = dayKey

            previousDay = dayKey
            hasPreviousDate = True

        Else

            dates(r) = 0

            If Not IsBlankVariant(dateValue) Then

                If Not hasPreviousDate Then

                    Err.Raise vbObjectError + 306, , _
                        "日付開始前に文字列行があります: " & _
                        ws.Name

                End If

                hasSummaryStarted = True

            End If

        End If

    Next r

    If Not hasPreviousDate Then

        Err.Raise vbObjectError + 307, , _
            "有効な日付データがありません: " & ws.Name

    End If

End Sub


'========================================================
' PARSE DATE
'========================================================

Private Function ParseDateKey( _
    ByVal v As Variant _
) As Long

    Dim numericDate As Double

    ParseDateKey = 0

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function

    If IsNumeric(v) Then

        numericDate = CDbl(v)

        If numericDate < 1# Then Exit Function
        If numericDate > 2958465# Then Exit Function

        ParseDateKey = CLng(Fix(numericDate))

        Exit Function

    End If

    If IsDate(v) Then

        numericDate = CDbl(CDate(v))

        If ThisWorkbook.Date1904 Then

            numericDate = numericDate - 1462#

        End If

        If numericDate < 1# Then Exit Function
        If numericDate > 2958465# Then Exit Function

        ParseDateKey = CLng(Fix(numericDate))

    End If

End Function


'========================================================
' BUILD SERIES DATA
'
' series(1,i): 日付
' series(2,i): 水準
'
' 欠損は系列ごとに処理。
'========================================================

Private Sub BuildSeriesData( _
    ByRef dataBlock As Variant, _
    ByRef dates() As Long, _
    ByVal colNum As Long, _
    ByRef series As Variant, _
    ByRef seriesCount As Long, _
    ByRef skippedCount As Long _
)

    Dim tempSeries() As Double

    Dim r As Long
    Dim v As Variant

    ReDim tempSeries(1 To 2, 1 To UBound(dates))

    seriesCount = 0
    skippedCount = 0

    For r = LBound(dates) To UBound(dates)

        If dates(r) > 0 Then

            v = dataBlock(r, colNum)

            If IsUsableNumericValue(v) Then

                seriesCount = seriesCount + 1

                tempSeries(1, seriesCount) = _
                    CDbl(dates(r))

                tempSeries(2, seriesCount) = _
                    CDbl(v)

            Else

                skippedCount = skippedCount + 1

            End If

        End If

    Next r

    If seriesCount = 0 Then

        series = Empty
        Exit Sub

    End If

    ReDim Preserve tempSeries( _
        1 To 2, 1 To seriesCount)

    series = tempSeries

End Sub


'========================================================
' VALID NUMERIC
'========================================================

Private Function IsUsableNumericValue( _
    ByVal v As Variant _
) As Boolean

    IsUsableNumericValue = False

    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function

    If Len(Trim$(CStr(v))) = 0 Then
        Exit Function
    End If

    IsUsableNumericValue = IsNumeric(v)

End Function


'========================================================
' BLANK VARIANT
'========================================================

Private Function IsBlankVariant( _
    ByVal v As Variant _
) As Boolean

    IsBlankVariant = False

    If IsError(v) Then Exit Function

    If IsEmpty(v) Then

        IsBlankVariant = True
        Exit Function

    End If

    If IsNull(v) Then

        IsBlankVariant = True
        Exit Function

    End If

    IsBlankVariant = _
        (Len(Trim$(CStr(v))) = 0)

End Function


'========================================================
' LATEST LEVEL
'========================================================

Private Function LatestLevel( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    If seriesCount = 0 Then

        LatestLevel = "NA"
        Exit Function

    End If

    LatestLevel = _
        series(2, seriesCount) * outputMultiplier

End Function


'========================================================
' N-OBSERVATION CHANGE
'========================================================

Private Function ChangeN( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal n As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    If seriesCount <= n Then

        ChangeN = "NA"
        Exit Function

    End If

    ChangeN = _
        (series(2, seriesCount) - _
         series(2, seriesCount - n)) * outputMultiplier

End Function


'========================================================
' Z-SCORE
'
' 最新N個の有効水準
'========================================================

Private Function ZScoreLevel( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal n As Long _
) As Variant

    Dim i As Long
    Dim firstIndex As Long

    Dim avgVal As Double
    Dim sumSq As Double
    Dim sdVal As Double

    If seriesCount < n Then

        ZScoreLevel = "NA"
        Exit Function

    End If

    firstIndex = seriesCount - n + 1

    For i = firstIndex To seriesCount

        avgVal = avgVal + series(2, i)

    Next i

    avgVal = avgVal / n

    For i = firstIndex To seriesCount

        sumSq = sumSq + _
            (series(2, i) - avgVal) ^ 2

    Next i

    sdVal = Sqr(sumSq / (n - 1))

    If sdVal = 0# Then

        ZScoreLevel = "NA"
        Exit Function

    End If

    ZScoreLevel = _
        (series(2, seriesCount) - avgVal) / sdVal

End Function


'========================================================
' HIGH / LOW
'
' 最新N個の有効水準の最大・最小
'========================================================

Private Function HighLowLevel( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal n As Long, _
    ByVal outputMultiplier As Double, _
    ByVal isHigh As Boolean _
) As Variant

    Dim firstIndex As Long
    Dim i As Long

    Dim extremeValue As Double
    Dim currentValue As Double

    If seriesCount < n Then

        HighLowLevel = "NA"
        Exit Function

    End If

    firstIndex = seriesCount - n + 1

    extremeValue = series(2, firstIndex)

    For i = firstIndex + 1 To seriesCount

        currentValue = series(2, i)

        If isHigh Then

            If currentValue > extremeValue Then
                extremeValue = currentValue
            End If

        Else

            If currentValue < extremeValue Then
                extremeValue = currentValue
            End If

        End If

    Next i

    HighLowLevel = extremeValue * outputMultiplier

End Function


'========================================================
' DAILY CHANGE VOLATILITY
'========================================================

Private Function VolDailyChange( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal n As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    Dim changes() As Double

    Dim i As Long
    Dim j As Long

    If seriesCount <= n Then

        VolDailyChange = "NA"
        Exit Function

    End If

    ReDim changes(1 To n)

    j = 1

    For i = seriesCount - n + 1 To seriesCount

        changes(j) = _
            series(2, i) - series(2, i - 1)

        j = j + 1

    Next i

    VolDailyChange = _
        StDevSampleArray(changes) * outputMultiplier

End Function


'========================================================
' SAMPLE STANDARD DEVIATION
'========================================================

Private Function StDevSampleArray( _
    ByRef arr() As Double _
) As Double

    Dim i As Long
    Dim n As Long

    Dim avgVal As Double
    Dim sumVal As Double
    Dim sumSq As Double

    n = UBound(arr) - LBound(arr) + 1

    If n < 2 Then

        StDevSampleArray = 0#
        Exit Function

    End If

    For i = LBound(arr) To UBound(arr)

        sumVal = sumVal + arr(i)

    Next i

    avgVal = sumVal / n

    For i = LBound(arr) To UBound(arr)

        sumSq = sumSq + (arr(i) - avgVal) ^ 2

    Next i

    StDevSampleArray = Sqr(sumSq / (n - 1))

End Function


'========================================================
' EXPLANATORY CHANGE MAP
'
' Key: 開始日|終了日
'========================================================

Private Sub BuildChangeMap( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal changeMap As Object _
)

    Dim i As Long
    Dim key As String

    If seriesCount < 2 Then Exit Sub

    For i = 2 To seriesCount

        key = DatePairKey( _
            CLng(series(1, i - 1)), _
            CLng(series(1, i)))

        If Not changeMap.Exists(key) Then

            changeMap.Add key, _
                series(2, i) - series(2, i - 1)

        End If

    Next i

End Sub


'========================================================
' DATE PAIR KEY
'========================================================

Private Function DatePairKey( _
    ByVal startDate As Long, _
    ByVal endDate As Long _
) As String

    DatePairKey = _
        CStr(startDate) & "|" & CStr(endDate)

End Function


'========================================================
' CROSS-SHEET REGRESSION
'
' y = alpha + beta*x + epsilon
'
' 共通の日付区間の変化幅だけを使用。
'========================================================

Private Sub Regression25DailyChangeCrossSheet( _
    ByRef targetSeries As Variant, _
    ByVal targetCount As Long, _
    ByVal xChangeMap As Object, _
    ByRef betaOut As Variant, _
    ByRef r2Out As Variant, _
    ByRef tValueOut As Variant _
)

    Dim x() As Double
    Dim y() As Double

    Dim i As Long
    Dim j As Long

    Dim countObs As Long
    Dim key As String

    Dim xMean As Double
    Dim yMean As Double

    Dim sxx As Double
    Dim syy As Double
    Dim sxy As Double

    Dim beta As Double
    Dim r2 As Double

    Dim sse As Double
    Dim sigma2 As Double
    Dim seBeta As Double

    betaOut = "NA"
    r2Out = "NA"
    tValueOut = "NA"

    If targetCount <= REGRESSION_N Then Exit Sub
    If xChangeMap.Count < REGRESSION_N Then Exit Sub

    ReDim x(1 To REGRESSION_N)
    ReDim y(1 To REGRESSION_N)

    countObs = 0

    For i = targetCount To 2 Step -1

        key = DatePairKey( _
            CLng(targetSeries(1, i - 1)), _
            CLng(targetSeries(1, i)))

        If xChangeMap.Exists(key) Then

            countObs = countObs + 1

            x(countObs) = CDbl(xChangeMap(key))

            y(countObs) = _
                targetSeries(2, i) - _
                targetSeries(2, i - 1)

            If countObs = REGRESSION_N Then
                Exit For
            End If

        End If

    Next i

    If countObs < REGRESSION_N Then Exit Sub

    ' Means
    For j = 1 To REGRESSION_N

        xMean = xMean + x(j)
        yMean = yMean + y(j)

    Next j

    xMean = xMean / REGRESSION_N
    yMean = yMean / REGRESSION_N

    ' Centered sums
    For j = 1 To REGRESSION_N

        sxx = sxx + (x(j) - xMean) ^ 2
        syy = syy + (y(j) - yMean) ^ 2

        sxy = sxy + _
            (x(j) - xMean) * (y(j) - yMean)

    Next j

    If sxx <= 0# Or syy <= 0# Then Exit Sub

    beta = sxy / sxx

    r2 = (sxy ^ 2) / (sxx * syy)

    If r2 < 0# Then r2 = 0#
    If r2 > 1# Then r2 = 1#

    betaOut = beta
    r2Out = r2

    sse = syy - beta * sxy

    If sse < 0# Then

        If Abs(sse) < 0.000000000001 Then
            sse = 0#
        End If

    End If

    If sse <= 0# Then Exit Sub

    sigma2 = sse / (REGRESSION_N - 2)
    seBeta = Sqr(sigma2 / sxx)

    If seBeta = 0# Then Exit Sub

    tValueOut = beta / seBeta

End Sub


'========================================================
' GET WORKSHEET
'========================================================

Private Function GetWorksheetOrError( _
    ByVal sheetName As String _
) As Worksheet

    On Error GoTo NotFound

    Set GetWorksheetOrError = _
        ThisWorkbook.Worksheets(sheetName)

    Exit Function

NotFound:

    Err.Raise vbObjectError + 600, , _
        "シートが見つかりません: " & sheetName

End Function


'========================================================
' GET OR CREATE WORKSHEET
'========================================================

Private Function GetOrCreateWorksheet( _
    ByVal sheetName As String _
) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = sheetName

    End If

    Set GetOrCreateWorksheet = ws

End Function


'========================================================
' HEADER VALIDATION
'========================================================

Private Sub ValidateHeader( _
    ByVal ws As Worksheet, _
    ByVal colNum As Long _
)

    Dim v As Variant

    v = ws.Cells(1, colNum).value2

    If IsError(v) Then

        Err.Raise vbObjectError + 610, , _
            "ヘッダーがエラーです: " & _
            ws.Name & "!" & _
            ws.Cells(1, colNum).Address(False, False)

    End If

    If IsBlankVariant(v) Then

        Err.Raise vbObjectError + 611, , _
            "ヘッダーが空白です: " & _
            ws.Name & "!" & _
            ws.Cells(1, colNum).Address(False, False)

    End If

End Sub


'========================================================
' FIND EXPLANATORY HEADER
'========================================================

Private Function FindHeaderColumnOrError( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal lastCol As Long _
) As Long

    Dim c As Long
    Dim foundCol As Long
    Dim foundCount As Long

    Dim h As String
    Dim v As Variant

    For c = 2 To lastCol

        v = ws.Cells(1, c).value2

        If IsError(v) Then

            Err.Raise vbObjectError + 620, , _
                "ヘッダーがエラーです: " & ws.Name

        End If

        If IsEmpty(v) Then

            h = ""

        Else

            h = Trim$(CStr(v))

        End If

        If StrComp( _
            h, Trim$(headerName), _
            vbTextCompare) = 0 Then

            foundCol = c
            foundCount = foundCount + 1

        End If

    Next c

    If foundCount = 0 Then

        Err.Raise vbObjectError + 621, , _
            "説明変数が見つかりません: " & _
            ws.Name & "!" & headerName

    End If

    If foundCount > 1 Then

        Err.Raise vbObjectError + 622, , _
            "説明変数ヘッダーが重複しています: " & _
            ws.Name & "!" & headerName

    End If

    FindHeaderColumnOrError = foundCol

End Function


'========================================================
' OUTPUT HEADERS
'
' A:R
'
' ヘッダー文字列だけを書き込む。
' セル書式・フィルター機能は変更しない。
'========================================================

Private Sub WriteHeaders( _
    ByVal ws As Worksheet _
)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Sheet", _
        "Trade", _
        "Level", _
        CStr(CHANGE_1_N) & "d-Chg", _
        CStr(CHANGE_5_N) & "d-Chg", _
        CStr(CHANGE_25_N) & "d-Chg", _
        CStr(ZSCORE_SHORT_N) & "d-Zscore", _
        CStr(HIGHLOW_SHORT_N) & "d-High", _
        CStr(HIGHLOW_SHORT_N) & "d-Low", _
        CStr(ZSCORE_LONG_N) & "d-Zscore", _
        CStr(HIGHLOW_LONG_N) & "d-High", _
        CStr(HIGHLOW_LONG_N) & "d-Low", _
        CStr(VOLATILITY_N) & "d-vol", _
        CStr(REGRESSION_N) & "d-beta", _
        "R2", _
        "t-value", _
        "Latest", _
        "Type")

    For i = LBound(headers) To UBound(headers)

        ws.Cells(1, i + 1).value2 = headers(i)

    Next i

End Sub


