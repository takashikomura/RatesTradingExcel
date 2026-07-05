Option Explicit

'========================================================
' Build covariance matrix from all time-series columns
'
' Source sheet assumption:
'   Row 1  : headers
'   Column A: dates
'   Column B onward: time-series data
'
' Output:
'   Sheet "Cov"
'   Sample covariance matrix of 1-day changes
'
' Notes:
'   - Covariance is calculated from changes, not levels.
'   - If source yield is in percent form, e.g. 0.010 = 1bp,
'     use multiplier = 100 to output covariance in bp^2.
'========================================================
Public Sub BuildCovMatrix()

    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsCov As Worksheet
    
    Dim srcName As String
    Dim lookbackInput As String
    Dim multiplierInput As String
    
    Dim lookbackDays As Long
    Dim multiplier As Double
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim nSeries As Long
    Dim obsCount As Long
    Dim firstChangeRow As Long
    Dim firstLevelRow As Long
    
    Dim headers() As String
    Dim changes() As Double
    Dim means() As Double
    Dim covMat() As Double
    Dim outArr() As Variant
    
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim outputStartRow As Long
    Dim outputStartCol As Long
    
    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim stateCaptured As Boolean
    
    Dim currentStep As String
    Dim errNum As Long
    Dim errDesc As String
    
    On Error GoTo ErrHandler
    
    currentStep = "Workbook取得"
    Set wb = ActiveWorkbook
    
    If wb Is Nothing Then
        Err.Raise vbObjectError + 1, , "アクティブなブックが見つかりません。"
    End If
    
    currentStep = "入力: 元データシート名"
    srcName = InputBox( _
        "共分散行列を計算する元データシート名を入力してください。" & vbCrLf & _
        "前提：A列=日付、B列以降=時系列データ、1行目=ヘッダー", _
        "Cov Matrix 作成" _
    )
    
    If Len(Trim$(srcName)) = 0 Then
        Err.Raise vbObjectError + 100, , "シート名が入力されていません。"
    End If
    
    currentStep = "元データシート取得"
    Set wsSrc = GetSheetOrError_Local(wb, srcName)
    
    If wsSrc.Name = "Cov" Then
        Err.Raise vbObjectError + 101, , "計算元シートに Cov は指定できません。"
    End If
    
    currentStep = "入力: Lookback"
    lookbackInput = InputBox( _
        "共分散行列の計算に使う直近日数を入力してください。" & vbCrLf & _
        "例：60, 125, 250" & vbCrLf & _
        "N日分の1日変化幅には、N+1個の水準データが必要です。", _
        "Lookback Days", _
        "60" _
    )
    
    If Len(Trim$(lookbackInput)) = 0 Then
        Err.Raise vbObjectError + 102, , "Lookback日数が入力されていません。"
    End If
    
    If Not IsNumeric(lookbackInput) Then
        Err.Raise vbObjectError + 103, , "Lookback日数が数値ではありません。"
    End If
    
    lookbackDays = CLng(lookbackInput)
    
    If lookbackDays < 2 Then
        Err.Raise vbObjectError + 104, , "Lookback日数は2以上を指定してください。"
    End If
    
    currentStep = "入力: 出力倍率"
    multiplierInput = InputBox( _
        "金利変化幅に掛ける出力倍率を入力してください。" & vbCrLf & _
        "例：" & vbCrLf & _
        "1   = 元データの単位のまま" & vbCrLf & _
        "100 = %表記の金利変化をbp表示に変換", _
        "Output Multiplier", _
        "1" _
    )
    
    If Len(Trim$(multiplierInput)) = 0 Then
        Err.Raise vbObjectError + 105, , "出力倍率が入力されていません。"
    End If
    
    If Not IsNumeric(multiplierInput) Then
        Err.Raise vbObjectError + 106, , "出力倍率が数値ではありません。"
    End If
    
    multiplier = CDbl(multiplierInput)
    
    If multiplier <= 0 Then
        Err.Raise vbObjectError + 107, , "出力倍率は0より大きい数値を指定してください。"
    End If
    
    currentStep = "Excel設定退避"
    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    stateCaptured = True
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    currentStep = "最終行・最終列取得"
    lastRow = wsSrc.Cells(wsSrc.rows.count, "A").End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.count).End(xlToLeft).Column
    
    If lastRow < 3 Then
        Err.Raise vbObjectError + 200, , _
            "データ行数が不足しています。最低でも2日分以上の水準データが必要です。"
    End If
    
    If lastCol < 2 Then
        Err.Raise vbObjectError + 201, , _
            "B列以降に時系列データがありません。1行目にヘッダーがあるか確認してください。"
    End If
    
    nSeries = lastCol - 1
    
    firstChangeRow = lastRow - lookbackDays + 1
    firstLevelRow = firstChangeRow - 1
    
    If firstChangeRow < 3 Then
        Err.Raise vbObjectError + 202, , _
            "Lookback日数に対してデータ数が不足しています。" & vbCrLf & _
            "指定Lookback: " & lookbackDays & vbCrLf & _
            "必要な水準データ数: " & (lookbackDays + 1) & vbCrLf & _
            "実際の水準データ数: " & (lastRow - 1)
    End If
    
    obsCount = lookbackDays
    
    currentStep = "ヘッダー検証"
    For c = 2 To lastCol
        
        If Len(Trim$(CStr(wsSrc.Cells(1, c).value))) = 0 Then
            Err.Raise vbObjectError + 300, , _
                "空白のヘッダーがあります。列番号: " & c
        End If
        
    Next c
    
    currentStep = "重複ヘッダー検証"
    For i = 2 To lastCol
        For j = i + 1 To lastCol
            
            If StrComp( _
                Trim$(CStr(wsSrc.Cells(1, i).value)), _
                Trim$(CStr(wsSrc.Cells(1, j).value)), _
                vbTextCompare _
            ) = 0 Then
            
                Err.Raise vbObjectError + 301, , _
                    "ヘッダー名が重複しています。" & vbCrLf & _
                    "ヘッダー: " & CStr(wsSrc.Cells(1, i).value) & vbCrLf & _
                    "列番号: " & i & " と " & j
            End If
            
        Next j
    Next i
    
    currentStep = "直近範囲の数値検証"
    For r = firstLevelRow To lastRow
        For c = 2 To lastCol
            
            If IsError(wsSrc.Cells(r, c).value) Then
                Err.Raise vbObjectError + 302, , _
                    "セルがエラー値です: " & wsSrc.Name & "!" & wsSrc.Cells(r, c).Address(False, False)
            End If
            
            If IsEmpty(wsSrc.Cells(r, c).value) Or Trim$(CStr(wsSrc.Cells(r, c).value)) = "" Then
                Err.Raise vbObjectError + 303, , _
                    "空白セルがあります: " & wsSrc.Name & "!" & wsSrc.Cells(r, c).Address(False, False)
            End If
            
            If Not IsNumeric(wsSrc.Cells(r, c).value) Then
                Err.Raise vbObjectError + 304, , _
                    "数値として認識できないセルがあります: " & wsSrc.Name & "!" & _
                    wsSrc.Cells(r, c).Address(False, False) & vbCrLf & _
                    "値: " & CStr(wsSrc.Cells(r, c).value)
            End If
            
        Next c
    Next r
    
    currentStep = "配列確保"
    ReDim headers(1 To nSeries)
    ReDim changes(1 To obsCount, 1 To nSeries)
    ReDim means(1 To nSeries)
    ReDim covMat(1 To nSeries, 1 To nSeries)
    ReDim outArr(1 To nSeries + 1, 1 To nSeries + 1)
    
    currentStep = "ヘッダー格納"
    For c = 2 To lastCol
        headers(c - 1) = Trim$(CStr(wsSrc.Cells(1, c).value))
    Next c
    
    currentStep = "1日変化幅作成"
    k = 1
    
    For r = firstChangeRow To lastRow
        For c = 2 To lastCol
            changes(k, c - 1) = _
                (CDbl(wsSrc.Cells(r, c).value) - CDbl(wsSrc.Cells(r - 1, c).value)) * multiplier
        Next c
        k = k + 1
    Next r
    
    currentStep = "平均変化幅計算"
    For j = 1 To nSeries
        For k = 1 To obsCount
            means(j) = means(j) + changes(k, j)
        Next k
        means(j) = means(j) / obsCount
    Next j
    
    currentStep = "共分散行列計算"
    For i = 1 To nSeries
        For j = i To nSeries
            
            covMat(i, j) = 0
            
            For k = 1 To obsCount
                covMat(i, j) = covMat(i, j) + _
                    (changes(k, i) - means(i)) * (changes(k, j) - means(j))
            Next k
            
            covMat(i, j) = covMat(i, j) / (obsCount - 1)
            covMat(j, i) = covMat(i, j)
            
        Next j
    Next i
    
    currentStep = "出力配列作成"
    outArr(1, 1) = "Cov"
    
    For j = 1 To nSeries
        outArr(1, j + 1) = headers(j)
    Next j
    
    For i = 1 To nSeries
        outArr(i + 1, 1) = headers(i)
        
        For j = 1 To nSeries
            outArr(i + 1, j + 1) = covMat(i, j)
        Next j
    Next i
    
    currentStep = "Covシート作成・クリア"
    Set wsCov = GetOrCreateSheet_Local(wb, "Cov")
    wsCov.Cells.Clear
    
    currentStep = "メタ情報出力"
    outputStartRow = 10
    outputStartCol = 1
    
    WriteCovMeta_Local wsCov, wsSrc, lookbackDays, obsCount, multiplier, firstChangeRow, lastRow
    
    currentStep = "共分散行列出力"
    wsCov.Cells(outputStartRow, outputStartCol).Resize(nSeries + 1, nSeries + 1).value = outArr
    
    currentStep = "フォーマット"
    FormatCovSheet_Local wsCov, nSeries, outputStartRow, outputStartCol
    
    If stateCaptured Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    
    MsgBox "Covシートへの共分散行列出力が完了しました。" & vbCrLf & _
           "計算元シート: " & wsSrc.Name & vbCrLf & _
           "系列数: " & nSeries & vbCrLf & _
           "Lookback: " & lookbackDays & "日" & vbCrLf & _
           "観測変化幅数: " & obsCount & vbCrLf & _
           "出力倍率: " & multiplier, vbInformation
    
    Exit Sub

ErrHandler:
    errNum = Err.Number
    errDesc = Err.Description
    
    On Error Resume Next
    
    If stateCaptured Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreenUpdating
        Application.EnableEvents = oldEnableEvents
    End If
    
    On Error GoTo 0
    
    If Len(errDesc) = 0 Then errDesc = "(Excelから詳細なエラー内容が返されていません)"
    
    MsgBox "共分散行列の作成中にエラーが発生しました。" & vbCrLf & _
           "Step: " & currentStep & vbCrLf & _
           "Err.Number: " & errNum & vbCrLf & _
           "内容: " & errDesc, vbCritical

End Sub

'========================================================
' Write metadata
'========================================================
Private Sub WriteCovMeta_Local( _
    ByVal wsCov As Worksheet, _
    ByVal wsSrc As Worksheet, _
    ByVal lookbackDays As Long, _
    ByVal obsCount As Long, _
    ByVal multiplier As Double, _
    ByVal firstChangeRow As Long, _
    ByVal lastRow As Long _
)

    With wsCov
        .Range("A1").value = "Source Sheet"
        .Range("B1").value = wsSrc.Name
        
        .Range("A2").value = "As Of Date"
        .Range("B2").value = wsSrc.Cells(lastRow, 1).value
        
        .Range("A3").value = "Lookback Days"
        .Range("B3").value = lookbackDays
        
        .Range("A4").value = "Observation Count"
        .Range("B4").value = obsCount
        
        .Range("A5").value = "Output Multiplier"
        .Range("B5").value = multiplier
        
        .Range("A6").value = "Method"
        .Range("B6").value = "Sample covariance of 1-day changes"
        
        .Range("A7").value = "Change Row Range"
        .Range("B7").value = firstChangeRow & ":" & lastRow
        
        .Range("A8").value = "Unit Note"
        .Range("B8").value = "If source is percent yield and multiplier=100, covariance unit is bp^2"
    End With

End Sub

'========================================================
' Formatting
' A列以外の列幅をそろえる
'========================================================
Private Sub FormatCovSheet_Local( _
    ByVal ws As Worksheet, _
    ByVal nSeries As Long, _
    ByVal outputStartRow As Long, _
    ByVal outputStartCol As Long _
)

    Dim lastRow As Long
    Dim lastCol As Long
    
    Dim matrixRange As Range
    Dim numberRange As Range
    Dim headerRowRange As Range
    Dim headerColRange As Range
    Dim diagRange As Range
    
    Dim i As Long
    Dim c As Long
    
    lastRow = outputStartRow + nSeries
    lastCol = outputStartCol + nSeries
    
    With ws
        
        Set matrixRange = .Range(.Cells(outputStartRow, outputStartCol), .Cells(lastRow, lastCol))
        Set numberRange = .Range(.Cells(outputStartRow + 1, outputStartCol + 1), .Cells(lastRow, lastCol))
        Set headerRowRange = .Range(.Cells(outputStartRow, outputStartCol), .Cells(outputStartRow, lastCol))
        Set headerColRange = .Range(.Cells(outputStartRow, outputStartCol), .Cells(lastRow, outputStartCol))
        
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
        
        '----------------------------------------
        ' Meta area
        '----------------------------------------
        .Range("A1:A8").Font.Bold = True
        .Range("A1:A8").Interior.Color = RGB(220, 230, 241)
        .Range("A1:B8").Borders.LineStyle = xlContinuous
        .Range("A1:B8").Borders.Color = RGB(200, 200, 200)
        .Range("B1:B8").WrapText = False
        
        '----------------------------------------
        ' Matrix area
        '----------------------------------------
        matrixRange.Borders.LineStyle = xlContinuous
        matrixRange.Borders.Color = RGB(210, 210, 210)
        
        headerRowRange.Font.Bold = True
        headerColRange.Font.Bold = True
        
        headerRowRange.Interior.Color = RGB(220, 230, 241)
        headerColRange.Interior.Color = RGB(220, 230, 241)
        
        headerRowRange.HorizontalAlignment = xlCenter
        headerColRange.HorizontalAlignment = xlLeft
        
        numberRange.NumberFormat = "0.000000"
        numberRange.HorizontalAlignment = xlRight
        
        '----------------------------------------
        ' Column width
        ' A列は別幅、B列以降はすべて同じ幅
        '----------------------------------------
        .Columns(1).ColumnWidth = 18
        
        For c = 2 To lastCol
            .Columns(c).ColumnWidth = 11
        Next c
        
        '----------------------------------------
        ' Row height
        '----------------------------------------
        .rows("1:8").RowHeight = 18
        .rows(outputStartRow).RowHeight = 18
        
        '----------------------------------------
        ' Diagonal variance highlight
        '----------------------------------------
        For i = 1 To nSeries
            If diagRange Is Nothing Then
                Set diagRange = .Cells(outputStartRow + i, outputStartCol + i)
            Else
                Set diagRange = Union(diagRange, .Cells(outputStartRow + i, outputStartCol + i))
            End If
        Next i
        
        If Not diagRange Is Nothing Then
            diagRange.Interior.Color = RGB(255, 242, 204)
            diagRange.Font.Bold = True
        End If
        
        '----------------------------------------
        ' Freeze panes
        '----------------------------------------
        On Error Resume Next
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(outputStartRow + 1, outputStartCol + 1).Select
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
        
    End With

End Sub

'========================================================
' Utilities
'========================================================
Private Function GetSheetOrError_Local( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
) As Worksheet

    On Error GoTo NotFound
    Set GetSheetOrError_Local = wb.Worksheets(sheetName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 600, , "指定されたシートが見つかりません: " & sheetName

End Function

Private Function GetOrCreateSheet_Local( _
    ByVal wb As Workbook, _
    ByVal sheetName As String _
) As Worksheet

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet_Local = ws

End Function

