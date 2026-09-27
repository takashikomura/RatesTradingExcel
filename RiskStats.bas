
Option Explicit

'========================================================
' RiskStats
'
' 複数シート対応
' 系列単位の欠損処理
' 共通説明変数による25日回帰
' 日付照合によるクロスシート回帰
' Latest列（日付型）
' 空白行・末尾集計行への対応
'
' 出力：
'
' A  Sheet
' B  Trade
' C  Level
' D  1d-Chg
' E  5d-Chg
' F  25d-Chg
' G  5d-Zscore
' H  25d-Zscore
' I  60d-Zscore
' J  5d-vol
' K  25d-vol
' L  60d-vol
' M  25d-beta
' N  R2
' O  t-value
' P  Latest
'
' 出力シートの既存書式は維持する。
' P列の表示形式のみ yyyy/mm/dd に設定する。
'
'========================================================

Private Const OUTPUT_SHEET As String = "RiskStats"

Private Const REGRESSION_N As Long = 25

Private Const OUTPUT_COLS As Long = 16


'========================================================
' Main
'========================================================

Public Sub BuildRiskStats()

    Dim sourceInput As String
    Dim normalizedInput As String

    Dim sheetItems As Variant
    Dim Item As Variant

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
    Dim i As Long

    Dim beta25 As Variant
    Dim r2_25 As Variant
    Dim tValue25 As Variant

    Dim sourceSummary As String

    Dim oldScreenUpdating As Boolean
    Dim screenUpdatingChanged As Boolean

    Dim oldOutputLastRow As Long
    Dim clearLastRow As Long

    On Error GoTo ErrHandler

    oldScreenUpdating = Application.ScreenUpdating


    '====================================================
    ' 1. 計算元シート指定
    '====================================================

    sourceInput = InputBox( _
        "計算元シートを指定してください。" & vbCrLf & _
        "複数指定はカンマ区切りで入力してください。" & _
        vbCrLf & vbCrLf & _
        "例：JGB,Swap,Curve", _
        "RiskStats - 計算元シート")

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
    ' 2. 出力倍率
    '====================================================

    multiplierInput = InputBox( _
        "出力倍率を入力してください。" & vbCrLf & _
        "1 = 元データの単位" & vbCrLf & _
        "100 = %表記の金利変化をbp表示に変換", _
        "RiskStats - 出力倍率", _
        "1")

    If Len(Trim$(multiplierInput)) = 0 Then
        Exit Sub
    End If

    If Not IsNumeric(multiplierInput) Then

        Err.Raise vbObjectError + 101, , _
            "出力倍率は数値で入力してください。"

    End If

    outputMultiplier = CDbl(multiplierInput)

    If outputMultiplier <= 0 Then

        Err.Raise vbObjectError + 102, , _
            "出力倍率は0より大きい数値を指定してください。"

    End If


    '====================================================
    ' 3. 説明変数シート
    '====================================================

    explanatorySheetName = InputBox( _
        "回帰説明変数が存在するシート名を入力してください。", _
        "RiskStats - 説明変数シート")

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
    ' 4. 説明変数系列
    '====================================================

    explanatoryName = InputBox( _
        "回帰説明変数とする系列のヘッダー名を入力してください。" & _
        vbCrLf & vbCrLf & _
        "例：10Y", _
        "RiskStats - 説明変数系列")

    If Len(Trim$(explanatoryName)) = 0 Then
        Exit Sub
    End If

    explanatoryName = Trim$(explanatoryName)


    '====================================================
    ' 5. 計算元シートの事前確認
    '====================================================

    Set sourceSheets = New Collection

    Set seenSheets = CreateObject("Scripting.Dictionary")

    seenSheets.CompareMode = vbTextCompare

    totalSeries = 0


    For Each Item In sheetItems

        sheetName = Trim$(CStr(Item))

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
                "B列以降にデータ系列がありません: " & sheetName

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

    Next Item


    If totalSeries = 0 Then

        Err.Raise vbObjectError + 108, , _
            "計算対象となる系列が存在しません。"

    End If


    '====================================================
    ' 6. 説明変数のデータ取得
    '
    ' 指定した1系列だけを抽出する。
    '
    ' 他の系列に欠損があっても影響しない。
    '====================================================

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


    '====================================================
    ' 回帰説明変数の変化幅を辞書に格納
    '
    ' Key:
    '   開始日|終了日
    '
    ' Value:
    '   説明変数の変化幅
    '
    ' 日付区間が一致する変化幅だけを回帰に使用する。
    '====================================================

    Set xChangeMap = CreateObject("Scripting.Dictionary")

    BuildChangeMap _
        expSeries, expCount, xChangeMap


    '====================================================
    ' 7. 出力配列
    '
    ' 全系列を計算してからRiskStatsに書き込む。
    '====================================================

    ReDim results(1 To totalSeries, 1 To OUTPUT_COLS)

    outputIndex = 0


    '====================================================
    ' 8. 各計算元シート
    '====================================================

    For sheetIndex = 1 To sourceSheets.Count

        Set wsSrc = sourceSheets(sheetIndex)


        '--------------------------------------------
        ' シート全体を取得
        '
        ' 日付列は共通だが、各系列の有効観測は
        ' それぞれ独立に構築する。
        '--------------------------------------------

        ReadSheetData _
            wsSrc, srcBlock, srcDates, lastCol


        '================================================
        ' 各系列
        '================================================

        For c = 2 To lastCol

            outputIndex = outputIndex + 1


            '--------------------------------------------
            ' この系列だけの有効観測を作成
            '
            ' 他の列の欠損は判定しない。
            '--------------------------------------------

            BuildSeriesData _
                srcBlock, srcDates, c, _
                srcSeries, seriesCount, skipped


            '--------------------------------------------
            ' A: Sheet
            '--------------------------------------------

            results(outputIndex, 1) = wsSrc.Name


            '--------------------------------------------
            ' B: Trade
            '--------------------------------------------

            results(outputIndex, 2) = _
                wsSrc.Cells(1, c).value2


            '--------------------------------------------
            ' C: Level
            '--------------------------------------------

            results(outputIndex, 3) = _
                LatestLevel( _
                    srcSeries, seriesCount, outputMultiplier)


            '--------------------------------------------
            ' D:F Changes
            '--------------------------------------------

            results(outputIndex, 4) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    1, outputMultiplier)

            results(outputIndex, 5) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    5, outputMultiplier)

            results(outputIndex, 6) = _
                ChangeN( _
                    srcSeries, seriesCount, _
                    25, outputMultiplier)


            '--------------------------------------------
            ' G:I Z-score
            '--------------------------------------------

            results(outputIndex, 7) = _
                ZScoreLevel(srcSeries, seriesCount, 5)

            results(outputIndex, 8) = _
                ZScoreLevel(srcSeries, seriesCount, 25)

            results(outputIndex, 9) = _
                ZScoreLevel(srcSeries, seriesCount, 60)


            '--------------------------------------------
            ' J:L Vol
            '--------------------------------------------

            results(outputIndex, 10) = _
                VolDailyChange( _
                    srcSeries, seriesCount, _
                    5, outputMultiplier)

            results(outputIndex, 11) = _
                VolDailyChange( _
                    srcSeries, seriesCount, _
                    25, outputMultiplier)

            results(outputIndex, 12) = _
                VolDailyChange( _
                    srcSeries, seriesCount, _
                    60, outputMultiplier)


            '--------------------------------------------
            ' M:O Cross-sheet regression
            '
            ' この系列と共通説明変数について、
            ' 日付区間が一致する変化幅を使用。
            '--------------------------------------------

            Regression25DailyChangeCrossSheet _
                srcSeries, seriesCount, _
                xChangeMap, _
                beta25, r2_25, tValue25


            results(outputIndex, 13) = beta25
            results(outputIndex, 14) = r2_25
            results(outputIndex, 15) = tValue25


            '--------------------------------------------
            ' P: Latest
            '
            ' この系列における最新の有効日付。
            '
            ' 全観測が欠損している系列はNA。
            '--------------------------------------------

            If seriesCount > 0 Then

                results(outputIndex, 16) = _
                    CDbl(srcSeries(1, seriesCount))

            Else

                results(outputIndex, 16) = "NA"

            End If

        Next c

    Next sheetIndex


    '====================================================
    ' 9. RiskStatsへ一括出力
    '
    ' すべての計算が完了してから実行する。
    '====================================================

    Set wsOut = GetOrCreateWorksheet(OUTPUT_SHEET)

    oldOutputLastRow = wsOut.Cells( _
        wsOut.rows.Count, 1).End(xlUp).Row

    clearLastRow = Application.WorksheetFunction.Max( _
        oldOutputLastRow, totalSeries + 1)


    Application.ScreenUpdating = False

    screenUpdatingChanged = True


    '--------------------------------------------
    ' 既存書式を維持したまま内容だけ削除
    '--------------------------------------------

    wsOut.Range( _
        wsOut.Cells(1, 1), _
        wsOut.Cells(clearLastRow, OUTPUT_COLS) _
    ).ClearContents


    '--------------------------------------------
    ' ヘッダー
    '--------------------------------------------

    WriteHeaders wsOut


    '--------------------------------------------
    ' 計算結果
    '--------------------------------------------

    wsOut.Range("A2").Resize( _
        totalSeries, OUTPUT_COLS).value2 = results


    '====================================================
    ' 10. Latest列の日付表示形式
    '
    ' P列のみ設定
    '====================================================

    wsOut.Range( _
        "P2:P" & CStr(totalSeries + 1) _
    ).numberFormat = "yyyy/mm/dd"


    Application.ScreenUpdating = oldScreenUpdating

    screenUpdatingChanged = False


    MsgBox _
        "RiskStatsの作成が完了しました。" & _
        vbCrLf & vbCrLf & _
        "計算元シート: " & sourceSummary & vbCrLf & _
        "対象系列数: " & totalSeries & vbCrLf & _
        "出力倍率: " & outputMultiplier & vbCrLf & _
        "回帰説明変数: " & _
        wsExp.Name & "!" & explanatoryName & vbCrLf & vbCrLf & _
        "欠損値は系列ごとに除外して計算しました。", _
        vbInformation

    Exit Sub


ErrHandler:

    If screenUpdatingChanged Then

        Application.ScreenUpdating = oldScreenUpdating

    End If

    MsgBox _
        "RiskStats作成中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'========================================================
' シートデータ取得
'
' 日付列の有効性を判定する。
'
' 数値系列の欠損判定はここでは行わない。
' 各系列のBuildSeriesDataで独立に処理する。
'
' dates(r)
'   > 0 : 有効な日付
'   = 0 : 空白行・集計行等
'
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
            "B列以降にデータ系列がありません。" & _
            vbCrLf & _
            "シート: " & ws.Name

    End If


    If lastRow < 2 Then

        Err.Raise vbObjectError + 301, , _
            "時系列データがありません。" & _
            vbCrLf & _
            "シート: " & ws.Name

    End If


    If IsError(ws.Cells(1, 1).value2) Then

        Err.Raise vbObjectError + 302, , _
            "日付ヘッダーがエラー値です。" & _
            vbCrLf & _
            "シート: " & ws.Name

    End If


    '--------------------------------------------
    ' データを配列へ読み込む
    '
    ' dataBlock(1, 1) = シートA2
    ' dataBlock(2, 1) = シートA3
    '--------------------------------------------

    dataBlock = ws.Range( _
        ws.Cells(2, 1), _
        ws.Cells(lastRow, lastCol) _
    ).value2


    ReDim dates(1 To UBound(dataBlock, 1))


    previousDay = 0

    hasPreviousDate = False
    hasSummaryStarted = False


    '====================================================
    ' 日付列を走査
    '====================================================

    For r = 1 To UBound(dataBlock, 1)

        dateValue = dataBlock(r, 1)


        If IsError(dateValue) Then

            Err.Raise vbObjectError + 303, , _
                "日付列にエラー値があります。" & _
                vbCrLf & _
                "シート: " & ws.Name & vbCrLf & _
                "行: " & CStr(r + 1)

        End If


        dayKey = ParseDateKey(dateValue)


        If dayKey > 0 Then


            '----------------------------------------
            ' 集計領域の後に日付が再登場した場合
            '----------------------------------------

            If hasSummaryStarted Then

                Err.Raise vbObjectError + 304, , _
                    "集計行の後に日付が存在します。" & _
                    vbCrLf & _
                    "シート: " & ws.Name & vbCrLf & _
                    "行: " & CStr(r + 1)

            End If


            '----------------------------------------
            ' 日付の昇順を確認
            '
            ' 数値系列の欠損とは独立に検証する。
            '----------------------------------------

            If hasPreviousDate Then

                If dayKey <= previousDay Then

                    Err.Raise vbObjectError + 305, , _
                        "日付が昇順ではありません。" & _
                        vbCrLf & _
                        "シート: " & ws.Name & vbCrLf & _
                        "行: " & CStr(r + 1)

                End If

            End If


            dates(r) = dayKey

            previousDay = dayKey

            hasPreviousDate = True


        Else


            '----------------------------------------
            ' 日付ではない行
            '
            ' 空白行は除外する。
            '
            ' EOD_Chg等の文字列が存在する場合、
            ' 以降を集計領域として扱う。
            '----------------------------------------

            dates(r) = 0


            If Not IsBlankVariant(dateValue) Then

                If Not hasPreviousDate Then

                    Err.Raise vbObjectError + 306, , _
                        "日付データ開始前に文字列行があります。" & _
                        vbCrLf & _
                        "シート: " & ws.Name & vbCrLf & _
                        "行: " & CStr(r + 1)

                End If

                hasSummaryStarted = True

            End If

        End If

    Next r


    If Not hasPreviousDate Then

        Err.Raise vbObjectError + 307, , _
            "有効な日付データがありません。" & _
            vbCrLf & _
            "シート: " & ws.Name

    End If

End Sub


'========================================================
' 日付判定
'
' Excel日付シリアル値をキーとして使用する。
'
' 日付でなければ0を返す。
'
' 時刻部分は切り捨てる。
'
'========================================================

Private Function ParseDateKey( _
    ByVal v As Variant _
) As Long

    Dim numericDate As Double
    Dim dayKey As Long


    ParseDateKey = 0


    If IsError(v) Then Exit Function

    If IsNull(v) Then Exit Function

    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbBoolean Then Exit Function


    '--------------------------------------------
    ' 数値の日付シリアル値
    '--------------------------------------------

    If IsNumeric(v) Then

        numericDate = CDbl(v)

        If numericDate < 1 Then Exit Function

        If numericDate > 2958465# Then Exit Function

        dayKey = CLng(Fix(numericDate))

        ParseDateKey = dayKey

        Exit Function

    End If


    '--------------------------------------------
    ' 文字列の日付
    '--------------------------------------------

    If IsDate(v) Then

        numericDate = CDbl(CDate(v))

        ' 1904日付システムのブックでは、
        ' VBA日付からExcel日付へ変換する。

        If ThisWorkbook.Date1904 Then

            numericDate = numericDate - 1462#

        End If

        If numericDate < 1 Then Exit Function

        If numericDate > 2958465# Then Exit Function

        ParseDateKey = CLng(Fix(numericDate))

    End If

End Function


'========================================================
' 系列単位の有効観測抽出
'
' dataBlock:
'   シート全体のデータ
'
' dates:
'   各行の日付キー
'
' colNum:
'   対象系列の列番号
'
' series(1, i):
'   日付キー
'
' series(2, i):
'   数値
'
' 欠損値があれば、その系列の当該行だけ除外する。
'
' 他の系列の有効観測には影響しない。
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

    Dim capacity As Long


    capacity = UBound(dates)

    ReDim tempSeries(1 To 2, 1 To capacity)

    seriesCount = 0
    skippedCount = 0


    For r = LBound(dates) To UBound(dates)


        '--------------------------------------------
        ' 日付ではない行は除外
        '--------------------------------------------

        If dates(r) > 0 Then

            v = dataBlock(r, colNum)


            '----------------------------------------
            ' この系列の値だけを検証
            '----------------------------------------

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


    '====================================================
    ' 有効データが存在しない系列
    '
    ' エラーにはしない。
    '
    ' 計算関数がNAを返す。
    '====================================================

    If seriesCount = 0 Then

        series = Empty

        Exit Sub

    End If


    '====================================================
    ' 有効観測だけに縮小
    '====================================================

    ReDim Preserve tempSeries( _
        1 To 2, 1 To seriesCount)


    series = tempSeries

End Sub


'========================================================
' 数値として利用できる値か判定
'
' False:
'   Empty
'   Null
'   空文字列
'   Excelエラー値
'   非数値
'
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


    If Not IsNumeric(v) Then
        Exit Function
    End If


    IsUsableNumericValue = True

End Function


'========================================================
' 空白値判定
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
' Latest Level
'
' 最新の有効観測を使用
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
' N-observation Change
'
' 欠損値を除外した後の有効観測で計算する。
'
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
        ( _
            series(2, seriesCount) _
            - _
            series(2, seriesCount - n) _
        ) * outputMultiplier

End Function


'========================================================
' Z-score
'
' 直近n個の有効なLevelを使用する。
'
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

    Dim x As Double


    If seriesCount < n Then

        ZScoreLevel = "NA"
        Exit Function

    End If


    firstIndex = seriesCount - n + 1


    '--------------------------------------------
    ' Mean
    '--------------------------------------------

    avgVal = 0


    For i = firstIndex To seriesCount

        avgVal = avgVal + series(2, i)

    Next i


    avgVal = avgVal / n


    '--------------------------------------------
    ' Sample standard deviation
    '--------------------------------------------

    sumSq = 0


    For i = firstIndex To seriesCount

        x = series(2, i)

        sumSq = sumSq + (x - avgVal) ^ 2

    Next i


    sdVal = Sqr(sumSq / (n - 1))


    If sdVal = 0 Then

        ZScoreLevel = "NA"
        Exit Function

    End If


    ZScoreLevel = _
        (series(2, seriesCount) - avgVal) / sdVal

End Function


'========================================================
' Volatility of daily changes
'
' 直近n本の有効観測間の変化幅を使用する。
'
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
' Sample standard deviation
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

        StDevSampleArray = 0
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
' 回帰説明変数の変化幅辞書
'
' Key:
'   開始日|終了日
'
' Value:
'   説明変数の変化幅
'
' 欠損値を除外した説明変数の系列から構築する。
'========================================================

Private Sub BuildChangeMap( _
    ByRef series As Variant, _
    ByVal seriesCount As Long, _
    ByVal changeMap As Object _
)

    Dim i As Long

    Dim key As String
    Dim changeValue As Double


    If seriesCount < 2 Then
        Exit Sub
    End If


    For i = 2 To seriesCount


        key = DatePairKey( _
            CLng(series(1, i - 1)), _
            CLng(series(1, i)))


        changeValue = _
            series(2, i) - series(2, i - 1)


        If Not changeMap.Exists(key) Then

            changeMap.Add key, changeValue

        End If

    Next i

End Sub


'========================================================
' 日付区間キー
'
' 例：
' 46280|46281
'
'========================================================

Private Function DatePairKey( _
    ByVal startDate As Long, _
    ByVal endDate As Long _
) As String

    DatePairKey = _
        CStr(startDate) & "|" & CStr(endDate)

End Function


'========================================================
' Cross-sheet regression
'
' y_t = alpha + beta*x_t + epsilon_t
'
' 被説明変数と説明変数の日付区間を照合する。
'
' 欠損によって日付区間が異なる変化幅は
' 回帰に採用しない。
'
' 直近25個の対応する変化幅を使用する。
'
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


    '====================================================
    ' 被説明変数のデータが不足
    '====================================================

    If targetCount <= REGRESSION_N Then
        Exit Sub
    End If


    If xChangeMap.Count < REGRESSION_N Then
        Exit Sub
    End If


    ReDim x(1 To REGRESSION_N)
    ReDim y(1 To REGRESSION_N)


    countObs = 0


    '====================================================
    ' 被説明変数の最新観測から遡る
    '====================================================

    For i = targetCount To 2 Step -1


        '--------------------------------------------
        ' 被説明変数側の変化幅の日付区間
        '--------------------------------------------

        key = DatePairKey( _
            CLng(targetSeries(1, i - 1)), _
            CLng(targetSeries(1, i)))


        '--------------------------------------------
        ' 説明変数側に同じ区間が存在する場合
        '--------------------------------------------

        If xChangeMap.Exists(key) Then


            countObs = countObs + 1


            '----------------------------------------
            ' X: 説明変数の変化幅
            '----------------------------------------

            x(countObs) = CDbl(xChangeMap(key))


            '----------------------------------------
            ' Y: 被説明変数の変化幅
            '----------------------------------------

            y(countObs) = _
                targetSeries(2, i) _
                - _
                targetSeries(2, i - 1)


            If countObs = REGRESSION_N Then
                Exit For
            End If

        End If

    Next i


    '====================================================
    ' 25観測未満の場合
    '====================================================

    If countObs < REGRESSION_N Then
        Exit Sub
    End If


    '====================================================
    ' 平均
    '====================================================

    For j = 1 To REGRESSION_N

        xMean = xMean + x(j)
        yMean = yMean + y(j)

    Next j


    xMean = xMean / REGRESSION_N
    yMean = yMean / REGRESSION_N


    '====================================================
    ' 分散・共分散
    '====================================================

    For j = 1 To REGRESSION_N

        sxx = sxx + (x(j) - xMean) ^ 2

        syy = syy + (y(j) - yMean) ^ 2

        sxy = sxy + _
            (x(j) - xMean) * (y(j) - yMean)

    Next j


    If sxx <= 0 Or syy <= 0 Then
        Exit Sub
    End If


    '====================================================
    ' Beta
    '====================================================

    beta = sxy / sxx


    '====================================================
    ' R-squared
    '====================================================

    r2 = (sxy ^ 2) / (sxx * syy)

    If r2 < 0 Then r2 = 0
    If r2 > 1 Then r2 = 1


    betaOut = beta
    r2Out = r2


    '====================================================
    ' t-value
    '
    ' 自由度：25 - 2 = 23
    '====================================================

    sse = syy - beta * sxy


    ' 浮動小数点誤差への対応
    If sse < 0 Then

        If Abs(sse) < 0.000000000001 Then

            sse = 0

        End If

    End If


    If sse <= 0 Then

        tValueOut = "NA"
        Exit Sub

    End If


    sigma2 = sse / (REGRESSION_N - 2)

    seBeta = Sqr(sigma2 / sxx)


    If seBeta = 0 Then

        tValueOut = "NA"
        Exit Sub

    End If


    tValueOut = beta / seBeta

End Sub


'========================================================
' Worksheet取得
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
        "指定されたシートが見つかりません: " & _
        sheetName

End Function


'========================================================
' Worksheet作成・取得
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
' ヘッダー検証
'========================================================

Private Sub ValidateHeader( _
    ByVal ws As Worksheet, _
    ByVal colNum As Long _
)

    Dim v As Variant

    v = ws.Cells(1, colNum).value2


    If IsError(v) Then

        Err.Raise vbObjectError + 610, , _
            "ヘッダーにエラー値があります。" & vbCrLf & _
            ws.Name & "!" & _
            ws.Cells(1, colNum).Address(False, False)

    End If


    If IsBlankVariant(v) Then

        Err.Raise vbObjectError + 611, , _
            "ヘッダーが空白です。" & vbCrLf & _
            ws.Name & "!" & _
            ws.Cells(1, colNum).Address(False, False)

    End If

End Sub


'========================================================
' 説明変数のヘッダー検索
'========================================================

Private Function FindHeaderColumnOrError( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal lastCol As Long _
) As Long

    Dim c As Long

    Dim FoundCol As Long
    Dim FoundCount As Long

    Dim h As String
    Dim v As Variant


    For c = 2 To lastCol

        v = ws.Cells(1, c).value2


        If IsError(v) Then

            Err.Raise vbObjectError + 620, , _
                "ヘッダーにエラー値があります。" & vbCrLf & _
                ws.Name & "!" & _
                ws.Cells(1, c).Address(False, False)

        End If


        If IsEmpty(v) Then

            h = ""

        Else

            h = Trim$(CStr(v))

        End If


        If StrComp( _
            h, Trim$(headerName), vbTextCompare) = 0 Then

            FoundCol = c
            FoundCount = FoundCount + 1

        End If

    Next c


    If FoundCount = 0 Then

        Err.Raise vbObjectError + 621, , _
            "説明変数のヘッダーが見つかりません。" & _
            vbCrLf & _
            "シート: " & ws.Name & vbCrLf & _
            "系列名: " & headerName

    End If


    If FoundCount > 1 Then

        Err.Raise vbObjectError + 622, , _
            "説明変数のヘッダーが重複しています。" & _
            vbCrLf & _
            "シート: " & ws.Name & vbCrLf & _
            "系列名: " & headerName

    End If


    FindHeaderColumnOrError = FoundCol

End Function


'========================================================
' Output Headers
'
' セルの書式は変更しない。
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
        "1d-Chg", _
        "5d-Chg", _
        "25d-Chg", _
        "5d-Zscore", _
        "25d-Zscore", _
        "60d-Zscore", _
        "5d-vol", _
        "25d-vol", _
        "60d-vol", _
        "25d-beta", _
        "R2", _
        "t-value", _
        "Latest" _
    )


    For i = LBound(headers) To UBound(headers)

        ws.Cells(1, i + 1).value2 = headers(i)

    Next i

End Sub


