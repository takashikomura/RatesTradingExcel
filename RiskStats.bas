Option Explicit

'========================================================
' Main procedure
'========================================================
Public Sub BuildRiskStats()

    Dim srcNamesInput As String
    Dim normalizedInput As String
    Dim sheetItems As Variant
    Dim item As Variant

    Dim multiplierInput As String
    Dim outputMultiplier As Double

    Dim explanatorySheetName As String
    Dim explanatoryName As String
    Dim explanatoryCol As Long

    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim wsExp As Worksheet

    Dim sourceSheets As Collection
    Dim seenSheets As Object

    Dim sheetName As String

    Dim lastRow As Long
    Dim lastCol As Long

    Dim expLastRow As Long
    Dim expLastCol As Long

    Dim c As Long
    Dim outRow As Long

    Dim tradeName As String

    Dim beta25 As Variant
    Dim r2_25 As Variant
    Dim tValue25 As Variant

    Dim totalSeries As Long
    Dim sourceSummary As String

    On Error GoTo ErrHandler


    '====================================================
    ' 1. 計算元シートを複数指定
    '====================================================
    srcNamesInput = InputBox( _
        "計算元のシート名を入力してください。" & vbCrLf & _
        "複数指定する場合はカンマ区切りで入力してください。" & vbCrLf & vbCrLf & _
        "例：" & vbCrLf & _
        "JGB,Swap,Curve" & vbCrLf & vbCrLf & _
        "各シートの前提：" & vbCrLf & _
        "A列 = 日付" & vbCrLf & _
        "B列以降 = 時系列データ" & vbCrLf & _
        "1行目 = ヘッダー" & vbCrLf & _
        "日付 = 古いものから新しいものへの昇順", _
        "RiskStats 作成" _
    )

    If Len(Trim$(srcNamesInput)) = 0 Then
        Err.Raise vbObjectError + 100, , _
            "計算元シートが入力されていません。"
    End If


    ' 区切り文字を統一
    normalizedInput = srcNamesInput

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
        "例：" & vbCrLf & _
        "1   = 元データの単位のまま" & vbCrLf & _
        "100 = %表記の金利変化をbp表示に変換", _
        "出力倍率", _
        "1" _
    )

    If Len(Trim$(multiplierInput)) = 0 Then
        Err.Raise vbObjectError + 101, , _
            "出力倍率が入力されていません。"
    End If

    If Not IsNumeric(multiplierInput) Then
        Err.Raise vbObjectError + 102, , _
            "出力倍率が数値ではありません。"
    End If

    outputMultiplier = CDbl(multiplierInput)

    If outputMultiplier <= 0 Then
        Err.Raise vbObjectError + 103, , _
            "出力倍率は0より大きい数値を指定してください。"
    End If


    '====================================================
    ' 3. 回帰説明変数のシートを指定
    '====================================================
    explanatorySheetName = InputBox( _
        "25日回帰の説明変数が存在するシート名を入力してください。" & vbCrLf & vbCrLf & _
        "例：" & vbCrLf & _
        "JGB", _
        "回帰：説明変数シート" _
    )

    If Len(Trim$(explanatorySheetName)) = 0 Then
        Err.Raise vbObjectError + 104, , _
            "説明変数のシート名が入力されていません。"
    End If

    explanatorySheetName = Trim$(explanatorySheetName)

    If StrComp( _
        explanatorySheetName, _
        "RiskStats", _
        vbTextCompare) = 0 Then

        Err.Raise vbObjectError + 105, , _
            "RiskStats シートを回帰説明変数のシートには指定できません。"

    End If


    '====================================================
    ' 4. 回帰説明変数の系列を指定
    '====================================================
    explanatoryName = InputBox( _
        "25日回帰の説明変数とする系列のヘッダー名を入力してください。" & vbCrLf & vbCrLf & _
        "例：" & vbCrLf & _
        "10Y" & vbCrLf & _
        "5Y" & vbCrLf & _
        "2s5s", _
        "回帰：説明変数系列" _
    )

    If Len(Trim$(explanatoryName)) = 0 Then
        Err.Raise vbObjectError + 106, , _
            "説明変数の系列名が入力されていません。"
    End If

    explanatoryName = Trim$(explanatoryName)


    '====================================================
    ' 5. 回帰説明変数シートを取得・検証
    '====================================================
    Set wsExp = GetWorksheetOrError(explanatorySheetName)

    expLastRow = _
        wsExp.Cells( _
            wsExp.rows.count, _
            "A" _
        ).End(xlUp).Row

    expLastCol = _
        wsExp.Cells( _
            1, _
            wsExp.Columns.count _
        ).End(xlToLeft).Column

    explanatoryCol = _
        FindHeaderColumnOrError( _
            wsExp, _
            explanatoryName, _
            expLastCol _
        )

    ' 説明変数として実際に使う日付列・系列だけを検証
    ValidateRegressionSeries _
        wsExp, _
        expLastRow, _
        explanatoryCol


    '====================================================
    ' 6. 計算元シートを事前検証
    '
    ' RiskStatsを消去する前にすべて確認する
    '====================================================
    Set sourceSheets = New Collection

    Set seenSheets = CreateObject("Scripting.Dictionary")
    seenSheets.CompareMode = vbTextCompare


    For Each item In sheetItems

        sheetName = Trim$(CStr(item))

        If Len(sheetName) > 0 Then

            If StrComp( _
                sheetName, _
                "RiskStats", _
                vbTextCompare) = 0 Then

                Err.Raise vbObjectError + 107, , _
                    "計算元シートに RiskStats は指定できません。"

            End If


            ' 同一シートの二重指定を禁止
            If seenSheets.Exists(sheetName) Then

                Err.Raise vbObjectError + 108, , _
                    "同じシートが複数回指定されています: " & _
                    sheetName

            End If

            seenSheets.Add sheetName, True


            Set wsSrc = GetWorksheetOrError(sheetName)

            lastRow = _
                wsSrc.Cells( _
                    wsSrc.rows.count, _
                    "A" _
                ).End(xlUp).Row

            lastCol = _
                wsSrc.Cells( _
                    1, _
                    wsSrc.Columns.count _
                ).End(xlToLeft).Column


            ValidateSourceSheet _
                wsSrc, _
                lastRow, _
                lastCol


            sourceSheets.Add wsSrc

        End If

    Next item


    If sourceSheets.count = 0 Then

        Err.Raise vbObjectError + 109, , _
            "有効な計算元シートが指定されていません。"

    End If


    '====================================================
    ' 7. RiskStats取得
    '====================================================
    Set wsOut = GetOrCreateWorksheet("RiskStats")


    '====================================================
    ' 値だけ削除
    '
    ' A:O
    '
    ' 書式は一切変更しない
    '====================================================
    wsOut.Range("A:O").ClearContents


    '====================================================
    ' 8. ヘッダー
    '====================================================
    WriteHeaders wsOut


    '====================================================
    ' 9. 各シートを順番に計算
    '====================================================
    outRow = 2
    totalSeries = 0


    For Each wsSrc In sourceSheets

        lastRow = _
            wsSrc.Cells( _
                wsSrc.rows.count, _
                "A" _
            ).End(xlUp).Row

        lastCol = _
            wsSrc.Cells( _
                1, _
                wsSrc.Columns.count _
            ).End(xlToLeft).Column


        For c = 2 To lastCol

            tradeName = _
                Trim$(CStr(wsSrc.Cells(1, c).value))


            '================================================
            ' A列：取得元シート名
            '================================================
            wsOut.Cells(outRow, 1).value = wsSrc.Name


            '================================================
            ' B列：系列名
            '================================================
            wsOut.Cells(outRow, 2).value = tradeName


            '================================================
            ' C列：Level
            '================================================
            wsOut.Cells(outRow, 3).value = _
                LatestLevel( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    outputMultiplier _
                )


            '================================================
            ' D:F Changes
            '================================================
            wsOut.Cells(outRow, 4).value = _
                ChangeN( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    1, _
                    outputMultiplier _
                )

            wsOut.Cells(outRow, 5).value = _
                ChangeN( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    5, _
                    outputMultiplier _
                )

            wsOut.Cells(outRow, 6).value = _
                ChangeN( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    25, _
                    outputMultiplier _
                )


            '================================================
            ' G:I Z-score
            '================================================
            wsOut.Cells(outRow, 7).value = _
                ZScoreLevel( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    5 _
                )

            wsOut.Cells(outRow, 8).value = _
                ZScoreLevel( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    25 _
                )

            wsOut.Cells(outRow, 9).value = _
                ZScoreLevel( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    60 _
                )


            '================================================
            ' J:L Vol
            '================================================
            wsOut.Cells(outRow, 10).value = _
                VolDailyChange( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    5, _
                    outputMultiplier _
                )

            wsOut.Cells(outRow, 11).value = _
                VolDailyChange( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    25, _
                    outputMultiplier _
                )

            wsOut.Cells(outRow, 12).value = _
                VolDailyChange( _
                    wsSrc, _
                    lastRow, _
                    c, _
                    60, _
                    outputMultiplier _
                )


            '================================================
            ' M:O
            '
            ' 指定した1つの共通説明変数に対する25d回帰
            '
            ' target:
            '   wsSrc の各系列
            '
            ' explanatory:
            '   wsExp の explanatoryName
            '
            ' シート間で日付を照合して計算する
            '================================================
            Regression25DailyChangeCrossSheet _
                wsSrc, _
                lastRow, _
                c, _
                wsExp, _
                expLastRow, _
                explanatoryCol, _
                beta25, _
                r2_25, _
                tValue25


            wsOut.Cells(outRow, 13).value = beta25
            wsOut.Cells(outRow, 14).value = r2_25
            wsOut.Cells(outRow, 15).value = tValue25


            outRow = outRow + 1

        Next c


        totalSeries = _
            totalSeries + _
            (lastCol - 1)


        If Len(sourceSummary) > 0 Then
            sourceSummary = sourceSummary & ", "
        End If

        sourceSummary = _
            sourceSummary & wsSrc.Name

    Next wsSrc


    '====================================================
    ' 書式設定処理は一切実施しない
    '====================================================


    '====================================================
    ' 10. 完了
    '====================================================
    MsgBox _
        "RiskStats の作成が完了しました。" & vbCrLf & vbCrLf & _
        "計算元シート数: " & sourceSheets.count & vbCrLf & _
        "計算元シート: " & sourceSummary & vbCrLf & _
        "対象系列数: " & totalSeries & vbCrLf & _
        "出力倍率: " & outputMultiplier & vbCrLf & vbCrLf & _
        "回帰説明変数: " & _
        wsExp.Name & "!" & explanatoryName, _
        vbInformation

    Exit Sub


ErrHandler:

    MsgBox _
        "RiskStats 作成中にエラーが発生しました。" & vbCrLf & vbCrLf & _
        "内容: " & Err.Description, _
        vbCritical

End Sub


'========================================================
' Validation
'========================================================
Private Sub ValidateSourceSheet( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long _
)

    Dim r As Long
    Dim c As Long
    Dim prevDate As Date
    Dim currDate As Date

    If lastRow < 2 Then
        Err.Raise vbObjectError + 300, , _
            "データ行がありません。" & _
            "シート: " & ws.Name
    End If

    If lastCol < 2 Then
        Err.Raise vbObjectError + 301, , _
            "B列以降に時系列データがありません。" & _
            "シート: " & ws.Name
    End If

    If Trim$(CStr(ws.Cells(1, 1).value)) = "" Then
        Err.Raise vbObjectError + 302, , _
            "A1に日付列のヘッダーがありません。" & _
            "シート: " & ws.Name
    End If


    For r = 2 To lastRow

        If Not IsDate(ws.Cells(r, 1).value) Then

            Err.Raise vbObjectError + 303, , _
                "A列に日付として認識できない値があります。" & vbCrLf & _
                "シート: " & ws.Name & vbCrLf & _
                "行: " & r

        End If


        If r > 2 Then

            prevDate = CDate(ws.Cells(r - 1, 1).value)
            currDate = CDate(ws.Cells(r, 1).value)

            If currDate <= prevDate Then

                Err.Raise vbObjectError + 304, , _
                    "A列の日付が昇順になっていません。" & vbCrLf & _
                    "シート: " & ws.Name & vbCrLf & _
                    "行 " & (r - 1) & " と 行 " & r

            End If

        End If

    Next r


    For c = 2 To lastCol

        If Trim$(CStr(ws.Cells(1, c).value)) = "" Then

            Err.Raise vbObjectError + 305, , _
                "1行目に空白のヘッダーがあります。" & vbCrLf & _
                "シート: " & ws.Name & vbCrLf & _
                "列番号: " & c

        End If


        For r = 2 To lastRow

            If IsError(ws.Cells(r, c).value) Then

                Err.Raise vbObjectError + 306, , _
                    "セルがエラー値です。" & vbCrLf & _
                    ws.Name & "!" & _
                    ws.Cells(r, c).Address(False, False)

            End If


            If IsEmpty(ws.Cells(r, c).value) Or _
               Trim$(CStr(ws.Cells(r, c).value)) = "" Then

                Err.Raise vbObjectError + 307, , _
                    "空白セルがあります。" & vbCrLf & _
                    ws.Name & "!" & _
                    ws.Cells(r, c).Address(False, False)

            End If


            If Not IsNumeric(ws.Cells(r, c).value) Then

                Err.Raise vbObjectError + 308, , _
                    "数値として認識できないデータがあります。" & vbCrLf & _
                    ws.Name & "!" & _
                    ws.Cells(r, c).Address(False, False)

            End If

        Next r

    Next c

End Sub


'========================================================
' 回帰説明変数系列だけを検証
'========================================================
Private Sub ValidateRegressionSeries( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal colNum As Long _
)

    Dim r As Long
    Dim prevDate As Date
    Dim currDate As Date

    If lastRow < 3 Then

        Err.Raise vbObjectError + 320, , _
            "回帰説明変数のデータが不足しています。" & vbCrLf & _
            "シート: " & ws.Name

    End If


    For r = 2 To lastRow

        If Not IsDate(ws.Cells(r, 1).value) Then

            Err.Raise vbObjectError + 321, , _
                "回帰説明変数シートのA列に日付でない値があります。" & vbCrLf & _
                ws.Name & "!" & _
                ws.Cells(r, 1).Address(False, False)

        End If


        If r > 2 Then

            prevDate = CDate(ws.Cells(r - 1, 1).value)
            currDate = CDate(ws.Cells(r, 1).value)

            If currDate <= prevDate Then

                Err.Raise vbObjectError + 322, , _
                    "回帰説明変数シートの日付が昇順ではありません。" & vbCrLf & _
                    "シート: " & ws.Name

            End If

        End If


        If IsError(ws.Cells(r, colNum).value) Then

            Err.Raise vbObjectError + 323, , _
                "回帰説明変数にエラー値があります。" & vbCrLf & _
                ws.Name & "!" & _
                ws.Cells(r, colNum).Address(False, False)

        End If


        If IsEmpty(ws.Cells(r, colNum).value) Or _
           Trim$(CStr(ws.Cells(r, colNum).value)) = "" Then

            Err.Raise vbObjectError + 324, , _
                "回帰説明変数に空白があります。" & vbCrLf & _
                ws.Name & "!" & _
                ws.Cells(r, colNum).Address(False, False)

        End If


        If Not IsNumeric(ws.Cells(r, colNum).value) Then

            Err.Raise vbObjectError + 325, , _
                "回帰説明変数に数値でない値があります。" & vbCrLf & _
                ws.Name & "!" & _
                ws.Cells(r, colNum).Address(False, False)

        End If

    Next r

End Sub


'========================================================
' Basic calculations
'========================================================
Private Function LatestLevel( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal colNum As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    If lastRow < 2 Then

        LatestLevel = "NA"

    Else

        LatestLevel = _
            GetNumericValue( _
                ws, _
                lastRow, _
                colNum _
            ) * outputMultiplier

    End If

End Function


Private Function ChangeN( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal colNum As Long, _
    ByVal n As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    If lastRow - n < 2 Then

        ChangeN = "NA"
        Exit Function

    End If


    ChangeN = _
        ( _
            GetNumericValue(ws, lastRow, colNum) - _
            GetNumericValue(ws, lastRow - n, colNum) _
        ) * outputMultiplier

End Function


Private Function ZScoreLevel( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal colNum As Long, _
    ByVal n As Long _
) As Variant

    Dim startRow As Long
    Dim avgVal As Double
    Dim sdVal As Double
    Dim latestVal As Double

    startRow = lastRow - n + 1


    If startRow < 2 Then

        ZScoreLevel = "NA"
        Exit Function

    End If


    latestVal = _
        GetNumericValue(ws, lastRow, colNum)

    avgVal = _
        MeanValues( _
            ws, _
            startRow, _
            lastRow, _
            colNum _
        )

    sdVal = _
        StDevSampleValues( _
            ws, _
            startRow, _
            lastRow, _
            colNum _
        )


    If sdVal = 0 Then

        ZScoreLevel = "NA"
        Exit Function

    End If


    ZScoreLevel = _
        (latestVal - avgVal) / sdVal

End Function


Private Function VolDailyChange( _
    ByVal ws As Worksheet, _
    ByVal lastRow As Long, _
    ByVal colNum As Long, _
    ByVal n As Long, _
    ByVal outputMultiplier As Double _
) As Variant

    Dim firstChangeRow As Long
    Dim r As Long
    Dim changes() As Double
    Dim i As Long

    firstChangeRow = _
        lastRow - n + 1


    If firstChangeRow < 3 Then

        VolDailyChange = "NA"
        Exit Function

    End If


    ReDim changes(1 To n)

    i = 1


    For r = firstChangeRow To lastRow

        changes(i) = _
            GetNumericValue(ws, r, colNum) - _
            GetNumericValue(ws, r - 1, colNum)

        i = i + 1

    Next r


    VolDailyChange = _
        StDevSampleArray(changes) * _
        outputMultiplier

End Function


'========================================================
' 25d single regression
'
' targetとexplanatoryが別シートでも可
'
' y_t = alpha + beta*x_t + epsilon_t
'
' 日付を照合し、
' targetの当日・前日の両日がexplanatoryにも存在する
' 直近25個の日次変化を使用する
'========================================================
Private Sub Regression25DailyChangeCrossSheet( _
    ByVal wsTarget As Worksheet, _
    ByVal targetLastRow As Long, _
    ByVal targetCol As Long, _
    ByVal wsX As Worksheet, _
    ByVal xLastRow As Long, _
    ByVal xCol As Long, _
    ByRef betaOut As Variant, _
    ByRef r2Out As Variant, _
    ByRef tValueOut As Variant _
)

    Const n As Long = 25

    Dim xRowByDate As Object

    Dim x() As Double
    Dim y() As Double

    Dim r As Long
    Dim xrCurr As Long
    Dim xrPrev As Long

    Dim currKey As String
    Dim prevKey As String

    Dim countObs As Long
    Dim i As Long

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


    If targetLastRow < 3 Or _
       xLastRow < 3 Then

        Exit Sub

    End If


    Set xRowByDate = _
        CreateObject("Scripting.Dictionary")


    '====================================================
    ' 説明変数側：
    ' Date -> Row の辞書を作る
    '====================================================
    For r = 2 To xLastRow

        currKey = _
            DateKey(wsX.Cells(r, 1).value)

        xRowByDate(currKey) = r

    Next r


    ReDim x(1 To n)
    ReDim y(1 To n)


    '====================================================
    ' target側を最新日から遡る
    '
    ' targetの
    '   t
    '   t-1
    '
    ' の両方が説明変数シートにも存在する場合だけ
    ' 1つの回帰観測値として採用
    '====================================================
    countObs = 0


    For r = targetLastRow To 3 Step -1

        currKey = _
            DateKey(wsTarget.Cells(r, 1).value)

        prevKey = _
            DateKey(wsTarget.Cells(r - 1, 1).value)


        If xRowByDate.Exists(currKey) And _
           xRowByDate.Exists(prevKey) Then


            xrCurr = CLng(xRowByDate(currKey))
            xrPrev = CLng(xRowByDate(prevKey))


            countObs = countObs + 1


            ' 説明変数の日次変化
            x(countObs) = _
                GetNumericValue(wsX, xrCurr, xCol) - _
                GetNumericValue(wsX, xrPrev, xCol)


            ' 被説明変数の日次変化
            y(countObs) = _
                GetNumericValue(wsTarget, r, targetCol) - _
                GetNumericValue(wsTarget, r - 1, targetCol)


            If countObs = n Then
                Exit For
            End If

        End If

    Next r


    ' 25観測取れなければNA
    If countObs < n Then
        Exit Sub
    End If


    '====================================================
    ' Mean
    '====================================================
    For i = 1 To n

        xMean = xMean + x(i)
        yMean = yMean + y(i)

    Next i


    xMean = xMean / n
    yMean = yMean / n


    '====================================================
    ' Variance / covariance terms
    '====================================================
    For i = 1 To n

        sxx = _
            sxx + _
            (x(i) - xMean) ^ 2

        syy = _
            syy + _
            (y(i) - yMean) ^ 2

        sxy = _
            sxy + _
            (x(i) - xMean) * _
            (y(i) - yMean)

    Next i


    If sxx = 0 Or syy = 0 Then
        Exit Sub
    End If


    '====================================================
    ' Beta
    '====================================================
    beta = sxy / sxx


    '====================================================
    ' R2
    '====================================================
    r2 = _
        (sxy ^ 2) / _
        (sxx * syy)


    If r2 < 0 Then r2 = 0
    If r2 > 1 Then r2 = 1


    betaOut = beta
    r2Out = r2


    '====================================================
    ' t-value
    '
    ' se(beta)
    ' = sqrt[
    '       SSE / (n-2) / Sxx
    '   ]
    '====================================================
    sse = _
        syy - _
        beta * sxy


    ' floating point adjustment
    If sse < 0 And _
       Abs(sse) < 0.000000000001 Then

        sse = 0

    End If


    If sse <= 0 Then

        tValueOut = "NA"
        Exit Sub

    End If


    sigma2 = _
        sse / (n - 2)


    seBeta = _
        Sqr(sigma2 / sxx)


    If seBeta = 0 Then

        tValueOut = "NA"
        Exit Sub

    End If


    tValueOut = _
        beta / seBeta

End Sub


'========================================================
' 日付をDictionaryのキーへ変換
'
' 時刻部分は無視して日付単位で一致させる
'========================================================
Private Function DateKey( _
    ByVal v As Variant _
) As String

    If Not IsDate(v) Then

        Err.Raise vbObjectError + 540, , _
            "日付として認識できない値があります。"

    End If


    DateKey = _
        CStr( _
            CLng( _
                Int( _
                    CDbl( _
                        CDate(v) _
                    ) _
                ) _
            ) _
        )

End Function


'========================================================
' Helper statistics
'========================================================
Private Function MeanValues( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal colNum As Long _
) As Double

    Dim r As Long
    Dim total As Double
    Dim countVal As Long


    For r = startRow To endRow

        total = _
            total + _
            GetNumericValue(ws, r, colNum)

        countVal = _
            countVal + 1

    Next r


    If countVal = 0 Then

        Err.Raise vbObjectError + 500, , _
            "平均値を計算できません。"

    End If


    MeanValues = _
        total / countVal

End Function


Private Function StDevSampleValues( _
    ByVal ws As Worksheet, _
    ByVal startRow As Long, _
    ByVal endRow As Long, _
    ByVal colNum As Long _
) As Double

    Dim r As Long
    Dim n As Long

    Dim avgVal As Double
    Dim sumSq As Double
    Dim x As Double


    n = _
        endRow - startRow + 1


    If n < 2 Then

        StDevSampleValues = 0
        Exit Function

    End If


    avgVal = _
        MeanValues( _
            ws, _
            startRow, _
            endRow, _
            colNum _
        )


    For r = startRow To endRow

        x = _
            GetNumericValue( _
                ws, _
                r, _
                colNum _
            )

        sumSq = _
            sumSq + _
            (x - avgVal) ^ 2

    Next r


    StDevSampleValues = _
        Sqr(sumSq / (n - 1))

End Function


Private Function StDevSampleArray( _
    ByRef arr() As Double _
) As Double

    Dim i As Long
    Dim n As Long

    Dim avgVal As Double
    Dim sumVal As Double
    Dim sumSq As Double


    n = _
        UBound(arr) - _
        LBound(arr) + 1


    If n < 2 Then

        StDevSampleArray = 0
        Exit Function

    End If


    For i = LBound(arr) To UBound(arr)

        sumVal = _
            sumVal + arr(i)

    Next i


    avgVal = _
        sumVal / n


    For i = LBound(arr) To UBound(arr)

        sumSq = _
            sumSq + _
            (arr(i) - avgVal) ^ 2

    Next i


    StDevSampleArray = _
        Sqr(sumSq / (n - 1))

End Function


Private Function GetNumericValue( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal colNum As Long _
) As Double

    Dim v As Variant


    v = _
        ws.Cells( _
            rowNum, _
            colNum _
        ).value


    If IsError(v) Then

        Err.Raise vbObjectError + 530, , _
            "セルがエラー値です: " & _
            ws.Name & "!" & _
            ws.Cells( _
                rowNum, _
                colNum _
            ).Address(False, False)

    End If


    If IsEmpty(v) Or _
       Trim$(CStr(v)) = "" Then

        Err.Raise vbObjectError + 531, , _
            "空白セルがあります: " & _
            ws.Name & "!" & _
            ws.Cells( _
                rowNum, _
                colNum _
            ).Address(False, False)

    End If


    If Not IsNumeric(v) Then

        Err.Raise vbObjectError + 532, , _
            "数値として認識できないセルがあります: " & _
            ws.Name & "!" & _
            ws.Cells( _
                rowNum, _
                colNum _
            ).Address(False, False)

    End If


    GetNumericValue = _
        CDbl(v)

End Function


'========================================================
' Sheet utilities
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


Private Function GetOrCreateWorksheet( _
    ByVal sheetName As String _
) As Worksheet

    Dim ws As Worksheet


    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0


    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:= _
                    ThisWorkbook.Worksheets( _
                        ThisWorkbook.Worksheets.count _
                    ) _
            )

        ws.Name = sheetName

    End If


    Set GetOrCreateWorksheet = ws

End Function


Private Function FindHeaderColumnOrError( _
    ByVal ws As Worksheet, _
    ByVal headerName As String, _
    ByVal lastCol As Long _
) As Long

    Dim c As Long
    Dim foundCol As Long
    Dim foundCount As Long
    Dim h As String


    For c = 2 To lastCol

        h = _
            Trim$( _
                CStr( _
                    ws.Cells(1, c).value _
                ) _
            )


        If StrComp( _
            h, _
            Trim$(headerName), _
            vbTextCompare _
        ) = 0 Then

            foundCol = c
            foundCount = foundCount + 1

        End If

    Next c


    If foundCount = 0 Then

        Err.Raise vbObjectError + 610, , _
            "ヘッダーが見つかりません。" & vbCrLf & _
            "シート: " & ws.Name & vbCrLf & _
            "ヘッダー: " & headerName

    End If


    If foundCount > 1 Then

        Err.Raise vbObjectError + 611, , _
            "ヘッダーが重複しています。" & vbCrLf & _
            "シート: " & ws.Name & vbCrLf & _
            "ヘッダー: " & headerName

    End If


    FindHeaderColumnOrError = _
        foundCol

End Function


'========================================================
' Output headers
'
' A = Sheet
' B = Trade
'
' 書式には触れず値だけ書く
'========================================================
Private Sub WriteHeaders(ByVal ws As Worksheet)

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
        "t-value" _
    )


    For i = LBound(headers) To UBound(headers)

        ws.Cells( _
            1, _
            i + 1 _
        ).value = headers(i)

    Next i

End Sub

