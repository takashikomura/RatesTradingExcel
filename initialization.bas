Option Explicit

Private Sub Workbook_Open()

    InitializeWorkbook

End Sub

Option Explicit

Public Sub InitializeWorkbook()

    Dim ws As Worksheet
    Dim originalStatusBar As Variant
    Dim sheetCount As Long
    Dim currentSheet As Long

    On Error GoTo ErrorHandler

    originalStatusBar = Application.StatusBar
    sheetCount = ThisWorkbook.Worksheets.count

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    '----------------------------------------
    '1. 計算モードを手動に設定
    '----------------------------------------
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False

    '----------------------------------------
    '2. 各シートを1回ずつ個別に再計算
    '----------------------------------------
    currentSheet = 0

    For Each ws In ThisWorkbook.Worksheets

        currentSheet = currentSheet + 1

        Application.StatusBar = _
            "シートを再計算しています... " & _
            currentSheet & "/" & sheetCount & _
            "  [" & ws.Name & "]"

        'このワークシートだけを1回計算
        ws.Calculate

        'Excelが応答なしになるのを防ぐ
        DoEvents

    Next ws

    '----------------------------------------
    '3. 共分散行列を作成
    '----------------------------------------
    Application.StatusBar = _
        "BuildCovMatricesを実行しています..."

    BuildCovMatrices

    '処理終了後も手動計算を維持
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False

ExitHandler:

    Application.StatusBar = originalStatusBar
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    'エラー発生時も手動計算を維持
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False

    MsgBox _
        "ブックの初期処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
        "処理中のシート: " & GetWorksheetName(ws) & vbCrLf & _
        "エラー番号: " & Err.Number & vbCrLf & _
        "エラー内容: " & Err.Description, _
        vbExclamation, _
        "初期処理エラー"

    Resume ExitHandler

End Sub

Private Function GetWorksheetName(ByVal ws As Worksheet) As String

    If ws Is Nothing Then
        GetWorksheetName = "シート計算前またはBuildCovMatrices実行中"
    Else
        GetWorksheetName = ws.Name
    End If

End Function

