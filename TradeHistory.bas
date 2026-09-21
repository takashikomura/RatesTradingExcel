
Option Explicit

'========================================================
' Trade History
'
' RegisterTrade:
'   トレードを末尾に追加
'
' DeleteLastTrade:
'   最後に登録したトレードを削除
'
' Tradeの自動判別:
'   Outright : 1 Leg
'   Curve    : 2 Legs
'   Fly      : 3 Legs
'
' TradeType:
'   Flow : 顧客フロー
'   Prop : 自己勘定
'
' デフォルトTradeType:
'   Flow
'
' 対象シート:
'   TradeHisotry
'
' 出力列:
'   A: Date
'   B: Time
'   C: Trade
'   D: Level
'   E: Leg1
'   F: Leg2
'   G: Leg3
'   H: TradeType
'
'========================================================

Private Const HISTORY_SHEET As String = "TradeHisotry"


'========================================================
' 1. トレード登録
'========================================================

Public Sub RegisterTrade()

    Dim ws As Worksheet

    Dim tradeName As String
    Dim tradeCategory As String
    Dim tradeType As String

    Dim legCount As Long

    Dim tradeLevel As Double

    Dim leg1Size As Double
    Dim leg2Size As Double
    Dim leg3Size As Double

    Dim nextRow As Long

    Dim tradeDateTime As Date

    On Error GoTo ErrHandler


    '====================================================
    ' STEP 1
    '
    ' Trade名を入力
    '
    ' Trade名からOutright・Curve・Flyを自動判別
    '====================================================

    If Not AskTradeName( _
        tradeName, _
        tradeCategory, _
        legCount) Then

        Exit Sub

    End If


    '====================================================
    ' STEP 2
    '
    ' Level
    '====================================================

    If Not AskNumericValue( _
        "Levelを入力してください。" & vbCrLf & _
        vbCrLf & _
        "Trade: " & tradeName & vbCrLf & _
        "Category: " & tradeCategory, _
        "Trade History - Level", _
        tradeLevel, _
        False) Then

        Exit Sub

    End If


    '====================================================
    ' STEP 3
    '
    ' Leg1
    '
    ' すべての取引で入力
    '====================================================

    If Not AskNumericValue( _
        "Leg1のサイズを入力してください。" & _
        vbCrLf & vbCrLf & _
        "Trade: " & tradeName & vbCrLf & _
        "Category: " & tradeCategory & vbCrLf & _
        "例：100、-100", _
        "Trade History - Leg1", _
        leg1Size, _
        True) Then

        Exit Sub

    End If


    '====================================================
    ' STEP 4
    '
    ' Leg2
    '
    ' CurveまたはFlyの場合のみ入力
    '====================================================

    If legCount >= 2 Then

        If Not AskNumericValue( _
            "Leg2のサイズを入力してください。" & _
            vbCrLf & vbCrLf & _
            "Trade: " & tradeName & vbCrLf & _
            "Leg1: " & CStr(leg1Size), _
            "Trade History - Leg2", _
            leg2Size, _
            True) Then

            Exit Sub

        End If

    End If


    '====================================================
    ' STEP 5
    '
    ' Leg3
    '
    ' Flyの場合のみ入力
    '====================================================

    If legCount = 3 Then

        If Not AskNumericValue( _
            "Leg3のサイズを入力してください。" & _
            vbCrLf & vbCrLf & _
            "Trade: " & tradeName & vbCrLf & _
            "Leg1: " & CStr(leg1Size) & vbCrLf & _
            "Leg2: " & CStr(leg2Size), _
            "Trade History - Leg3", _
            leg3Size, _
            True) Then

            Exit Sub

        End If

    End If


    '====================================================
    ' STEP 6
    '
    ' TradeType
    '
    ' Flow = 顧客フロー
    ' Prop = 自己勘定
    '
    ' デフォルト値：Flow
    '====================================================

    If Not AskTradeType(tradeType) Then

        Exit Sub

    End If


    '====================================================
    ' STEP 7
    '
    ' 入力がすべて完了してからシートを取得
    '
    ' 途中でキャンセルされた場合、
    ' TradeHisotryの内容は変更しない。
    '====================================================

    Set ws = GetTradeHistorySheet(True)


    '====================================================
    ' ヘッダー確認
    '
    ' 新規シート:
    '   8列形式のヘッダーを作成
    '
    ' 既存シート:
    '   必要に応じて旧形式から移行
    '====================================================

    InitializeTradeHistory ws


    '====================================================
    ' STEP 8
    '
    ' 次の登録行を取得
    '====================================================

    nextRow = GetLastHistoryRow(ws) + 1


    If nextRow > ws.rows.count Then

        Err.Raise vbObjectError + 700, , _
            "TradeHisotryシートの最終行に到達しました。"

    End If


    '====================================================
    ' STEP 9
    '
    ' 登録日時を取得
    '
    ' 入力がすべて完了した時点のPC日時を使用
    '====================================================

    tradeDateTime = Now


    '====================================================
    ' STEP 10
    '
    ' トレードを記録
    '====================================================

    With ws


        '--------------------------------------------
        ' A: Date
        '--------------------------------------------

        .Cells(nextRow, 1).value = _
            dateValue(tradeDateTime)


        '--------------------------------------------
        ' B: Time
        '--------------------------------------------

        .Cells(nextRow, 2).value = _
            TimeValue(tradeDateTime)


        '--------------------------------------------
        ' C: Trade
        '--------------------------------------------

        .Cells(nextRow, 3).value2 = tradeName


        '--------------------------------------------
        ' D: Level
        '--------------------------------------------

        .Cells(nextRow, 4).value2 = tradeLevel


        '--------------------------------------------
        ' E: Leg1
        '--------------------------------------------

        .Cells(nextRow, 5).value2 = leg1Size


        '--------------------------------------------
        ' F: Leg2
        '
        ' Curve・Flyの場合のみ記録
        '--------------------------------------------

        If legCount >= 2 Then

            .Cells(nextRow, 6).value2 = leg2Size

        Else

            .Cells(nextRow, 6).ClearContents

        End If


        '--------------------------------------------
        ' G: Leg3
        '
        ' Flyの場合のみ記録
        '--------------------------------------------

        If legCount = 3 Then

            .Cells(nextRow, 7).value2 = leg3Size

        Else

            .Cells(nextRow, 7).ClearContents

        End If


        '--------------------------------------------
        ' H: TradeType
        '
        ' FlowまたはProp
        '--------------------------------------------

        .Cells(nextRow, 8).value2 = tradeType


        '--------------------------------------------
        ' Date / Time表示形式
        '
        ' 追加した行のA列・B列だけに設定
        '--------------------------------------------

        .Cells(nextRow, 1).numberFormat = _
            "yyyy/mm/dd"

        .Cells(nextRow, 2).numberFormat = _
            "hh:mm:ss"


    End With


    '====================================================
    ' 場中の操作を妨げないよう、
    ' 登録完了のMsgBoxは表示しない。
    '====================================================

    Exit Sub


ErrHandler:

    MsgBox _
        "トレード登録中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'========================================================
' 2. Trade名の入力
'
' 入力したTrade名からLeg数を自動判別
'
' 1 = Outright
' 2 = Curve
' 3 = Fly
'
' キャンセルした場合はFalse
'========================================================

Private Function AskTradeName( _
    ByRef tradeName As String, _
    ByRef tradeCategory As String, _
    ByRef legCount As Long _
) As Boolean

    Dim answer As Variant

    Dim inputText As String


    AskTradeName = False


    Do

        answer = Application.InputBox( _
            Prompt:= _
                "Tradeを入力してください。" & vbCrLf & _
                vbCrLf & _
                "Outright：" & vbCrLf & _
                "10Y、JGB10Y、JN123" & vbCrLf & _
                vbCrLf & _
                "Curve：" & vbCrLf & _
                "2s5s、5s10s、5Y-10Y" & vbCrLf & _
                vbCrLf & _
                "Fly：" & vbCrLf & _
                "2s5s10s、5s10s30s" & vbCrLf & _
                vbCrLf & _
                "取引区分は自動判別されます。", _
            Title:="Trade History - Trade", _
            Type:=2)


        '--------------------------------------------
        ' キャンセル判定
        '--------------------------------------------

        If VarType(answer) = vbBoolean Then

            If answer = False Then
                Exit Function
            End If

        End If


        inputText = Trim$(CStr(answer))


        '--------------------------------------------
        ' 空欄
        '--------------------------------------------

        If Len(inputText) = 0 Then

            MsgBox _
                "Tradeを入力してください。", _
                vbExclamation

        Else


            '----------------------------------------
            ' Leg数を判別
            '----------------------------------------

            legCount = DetectTradeLegCount(inputText)


            Select Case legCount

                Case 1

                    tradeCategory = "Outright"

                Case 2

                    tradeCategory = "Curve"

                Case 3

                    tradeCategory = "Fly"

                Case Else

                    tradeCategory = ""

            End Select


            '----------------------------------------
            ' 正常に判別された場合
            '----------------------------------------

            If legCount >= 1 And legCount <= 3 Then

                tradeName = inputText

                AskTradeName = True

                Exit Function

            End If


            '----------------------------------------
            ' 判別できない場合は再入力
            '----------------------------------------

            MsgBox _
                "Trade名から取引区分を判別できません。" & _
                vbCrLf & vbCrLf & _
                "対応する入力例：" & vbCrLf & _
                "Outright: 10Y" & vbCrLf & _
                "Curve: 5s10s" & vbCrLf & _
                "Fly: 2s5s10s", _
                vbExclamation

        End If

    Loop

End Function


'========================================================
' 3. Trade名からLeg数を自動判別
'
' 戻り値：
'
' 1 = Outright
' 2 = Curve
' 3 = Fly
' 0 = 判別不能
'
' 対応例：
'
' Outright:
'   10Y
'   JGB10Y
'   JN123
'   JS123
'   JB123
'
' Curve:
'   2s5s
'   5s10s
'   5Y-10Y
'   5Y/10Y
'   5Y10Y
'
' Fly:
'   2s5s10s
'   5s10s30s
'   2Y-5Y-10Y
'   2Y/5Y/10Y
'   2Y5Y10Y
'
'========================================================

Private Function DetectTradeLegCount( _
    ByVal tradeName As String _
) As Long

    Dim normalizedName As String

    Dim tenorPattern As String


    DetectTradeLegCount = 0


    '====================================================
    ' 文字列を正規化
    '====================================================

    normalizedName = UCase$(Trim$(tradeName))

    normalizedName = Replace(normalizedName, " ", "")
    normalizedName = Replace(normalizedName, "　", "")
    normalizedName = Replace(normalizedName, vbTab, "")


    If Len(normalizedName) = 0 Then
        Exit Function
    End If


    '====================================================
    ' Tenorの基本パターン
    '
    ' 対応：
    '   2
    '   2.5
    '   5Y
    '   10Y
    '====================================================

    tenorPattern = "([0-9]+(\.[0-9]+)?Y?)"


    '====================================================
    ' STEP 1
    '
    ' Fly判定
    '====================================================


    '--------------------------------------------
    ' 2s5s10s
    ' 2s5s10
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^" & tenorPattern & "S" & _
              tenorPattern & "S" & _
              tenorPattern & "S?$") Then

        DetectTradeLegCount = 3

        Exit Function

    End If


    '--------------------------------------------
    ' 2Y-5Y-10Y
    ' 2Y/5Y/10Y
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^" & tenorPattern & "[/-]" & _
              tenorPattern & "[/-]" & _
              tenorPattern & "$") Then

        DetectTradeLegCount = 3

        Exit Function

    End If


    '--------------------------------------------
    ' 2Y5Y10Y
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^([0-9]+(\.[0-9]+)?Y){3}$") Then

        DetectTradeLegCount = 3

        Exit Function

    End If


    '====================================================
    ' STEP 2
    '
    ' Curve判定
    '====================================================


    '--------------------------------------------
    ' 5s10s
    ' 5s10
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^" & tenorPattern & "S" & _
              tenorPattern & "S?$") Then

        DetectTradeLegCount = 2

        Exit Function

    End If


    '--------------------------------------------
    ' 5Y-10Y
    ' 5Y/10Y
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^" & tenorPattern & "[/-]" & _
              tenorPattern & "$") Then

        DetectTradeLegCount = 2

        Exit Function

    End If


    '--------------------------------------------
    ' 5Y10Y
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^([0-9]+(\.[0-9]+)?Y){2}$") Then

        DetectTradeLegCount = 2

        Exit Function

    End If


    '====================================================
    ' STEP 3
    '
    ' Outright判定
    '
    ' 複数Legを表す文字列をOutrightと
    ' 誤判定しないようにする。
    '====================================================


    '--------------------------------------------
    ' 単独の年限
    '
    ' 例：2Y、10Y、2.5Y
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^[0-9]+(\.[0-9]+)?Y$") Then

        DetectTradeLegCount = 1

        Exit Function

    End If


    '--------------------------------------------
    ' JGB銘柄名
    '
    ' 例：
    ' JN123
    ' JS123
    ' JB123
    ' JL123
    ' JX123
    ' JU123
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^(JN|JS|JB|JL|JX|JU)[A-Z0-9_.]+$") Then

        DetectTradeLegCount = 1

        Exit Function

    End If


    '--------------------------------------------
    ' JGB10Yなどの単独年限名
    '--------------------------------------------

    If TradePatternMatches( _
        normalizedName, _
        "^[A-Z]{2,8}[0-9]+(\.[0-9]+)?Y$") Then

        DetectTradeLegCount = 1

        Exit Function

    End If


    '====================================================
    ' いずれにも該当しなければ判別不能
    '====================================================

    DetectTradeLegCount = 0

End Function


'========================================================
' 4. 正規表現による判定
'========================================================

Private Function TradePatternMatches( _
    ByVal textValue As String, _
    ByVal pattern As String _
) As Boolean

    Dim regex As Object


    Set regex = CreateObject("VBScript.RegExp")


    With regex

        .pattern = pattern

        .IgnoreCase = True

        .Global = False

    End With


    TradePatternMatches = regex.Test(textValue)

End Function


'========================================================
' 5. TradeType入力
'
' Flow:
'   顧客フロー
'
' Prop:
'   自己勘定ポジション
'
' デフォルト値：
'   Flow
'
' EnterだけでFlowを確定可能
'
' 大文字・小文字は区別しない。
'
' キャンセルの場合はFalseを返す。
'========================================================

Private Function AskTradeType( _
    ByRef tradeType As String _
) As Boolean

    Dim answer As Variant

    Dim inputText As String


    AskTradeType = False


    Do

        answer = Application.InputBox( _
            Prompt:= _
                "TradeTypeを入力してください。" & vbCrLf & _
                vbCrLf & _
                "Flow = 顧客フロー" & vbCrLf & _
                "Prop = 自己勘定ポジション" & vbCrLf & _
                vbCrLf & _
                "通常はFlowのままEnterキーで確定してください。", _
            Title:="Trade History - TradeType", _
            Default:="Flow", _
            Type:=2)


        '--------------------------------------------
        ' キャンセル
        '--------------------------------------------

        If VarType(answer) = vbBoolean Then

            If answer = False Then
                Exit Function
            End If

        End If


        inputText = LCase$(Trim$(CStr(answer)))


        '--------------------------------------------
        ' Flow / Propの判定
        '--------------------------------------------

        Select Case inputText


            Case "flow", ""

                tradeType = "Flow"

                AskTradeType = True

                Exit Function


            Case "prop"

                tradeType = "Prop"

                AskTradeType = True

                Exit Function


            Case Else

                MsgBox _
                    "TradeTypeはFlowまたはPropを入力してください。", _
                    vbExclamation


        End Select

    Loop

End Function


'========================================================
' 6. 数値入力
'
' requireNonZero = True:
'   0を禁止する。
'
' requireNonZero = False:
'   0を許可する。
'
' 負の数値も入力可能。
'
' キャンセルした場合はFalseを返す。
'========================================================

Private Function AskNumericValue( _
    ByVal promptText As String, _
    ByVal titleText As String, _
    ByRef resultValue As Double, _
    ByVal requireNonZero As Boolean _
) As Boolean

    Dim answer As Variant

    Dim inputText As String

    Dim numericValue As Double

    Dim conversionError As Long


    AskNumericValue = False


    Do

        answer = Application.InputBox( _
            Prompt:=promptText, _
            Title:=titleText, _
            Type:=2)


        '--------------------------------------------
        ' キャンセル
        '--------------------------------------------

        If VarType(answer) = vbBoolean Then

            If answer = False Then
                Exit Function
            End If

        End If


        inputText = Trim$(CStr(answer))


        '--------------------------------------------
        ' 空欄
        '--------------------------------------------

        If Len(inputText) = 0 Then

            MsgBox _
                "数値を入力してください。", _
                vbExclamation


        '--------------------------------------------
        ' 非数値
        '--------------------------------------------

        ElseIf Not IsNumeric(inputText) Then

            MsgBox _
                "数値として認識できません。" & _
                vbCrLf & _
                "再入力してください。", _
                vbExclamation


        Else


            '----------------------------------------
            ' 数値変換
            '----------------------------------------

            conversionError = 0

            Err.Clear

            On Error Resume Next

            numericValue = CDbl(inputText)

            conversionError = Err.Number

            Err.Clear

            On Error GoTo 0


            '----------------------------------------
            ' 数値変換失敗
            '----------------------------------------

            If conversionError <> 0 Then

                MsgBox _
                    "数値を正しく変換できません。" & _
                    vbCrLf & _
                    "再入力してください。", _
                    vbExclamation


            '----------------------------------------
            ' サイズが0
            '----------------------------------------

            ElseIf requireNonZero And numericValue = 0 Then

                MsgBox _
                    "サイズには0以外の数値を入力してください。", _
                    vbExclamation


            '----------------------------------------
            ' 正常入力
            '----------------------------------------

            Else

                resultValue = numericValue

                AskNumericValue = True

                Exit Function

            End If

        End If

    Loop

End Function


'========================================================
' 7. 最後に登録したトレードを削除
'
' 削除対象：
'   A:Hの最終トレード
'
' ヘッダー行は削除しない。
'
' 削除前に確認ダイアログを表示する。
'
' 行全体は削除せず、
' A:Hのセル内容だけを削除する。
'========================================================

Public Sub DeleteLastTrade()

    Dim ws As Worksheet

    Dim lastRow As Long

    Dim tradeName As String

    Dim tradeLevel As Variant

    Dim leg1Size As Variant
    Dim leg2Size As Variant
    Dim leg3Size As Variant

    Dim tradeType As String

    Dim confirmResult As VbMsgBoxResult


    On Error GoTo ErrHandler


    '====================================================
    ' STEP 1
    '
    ' シート取得
    '====================================================

    Set ws = GetTradeHistorySheet(False)


    If ws Is Nothing Then

        MsgBox _
            "TradeHisotryシートが存在しません。", _
            vbExclamation

        Exit Sub

    End If


    '====================================================
    ' STEP 2
    '
    ' ヘッダー確認
    '====================================================

    If Not HasTradeHistoryHeaders(ws) Then

        MsgBox _
            "TradeHisotryのヘッダーが正しくありません。" & _
            vbCrLf & _
            "削除処理を中止しました。", _
            vbExclamation

        Exit Sub

    End If


    '====================================================
    ' STEP 3
    '
    ' 最終入力行
    '====================================================

    lastRow = GetLastHistoryRow(ws)


    If lastRow <= 1 Then

        MsgBox _
            "削除できるトレードがありません。", _
            vbInformation

        Exit Sub

    End If


    '====================================================
    ' STEP 4
    '
    ' 最終行がトレードであることを確認
    '====================================================

    If Not IsValidTradeRow(ws, lastRow) Then

        MsgBox _
            "最終行が通常のトレード記録ではありません。" & _
            vbCrLf & _
            "誤削除防止のため処理を中止しました。", _
            vbExclamation

        Exit Sub

    End If


    '====================================================
    ' STEP 5
    '
    ' 削除対象の取得
    '====================================================

    tradeName = CStr(ws.Cells(lastRow, 3).value2)

    tradeLevel = ws.Cells(lastRow, 4).value2

    leg1Size = ws.Cells(lastRow, 5).value2

    leg2Size = ws.Cells(lastRow, 6).value2

    leg3Size = ws.Cells(lastRow, 7).value2

    tradeType = CStr(ws.Cells(lastRow, 8).value2)


    '====================================================
    ' STEP 6
    '
    ' 削除確認
    '====================================================

    confirmResult = MsgBox( _
        "以下のトレードを削除しますか？" & _
        vbCrLf & vbCrLf & _
        "Trade: " & tradeName & vbCrLf & _
        "Level: " & CStr(tradeLevel) & vbCrLf & _
        "Leg1: " & DisplayLegValue(leg1Size) & vbCrLf & _
        "Leg2: " & DisplayLegValue(leg2Size) & vbCrLf & _
        "Leg3: " & DisplayLegValue(leg3Size) & vbCrLf & _
        "TradeType: " & tradeType & vbCrLf & _
        vbCrLf & _
        "行番号: " & CStr(lastRow), _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Delete Last Trade")


    If confirmResult <> vbYes Then

        Exit Sub

    End If


    '====================================================
    ' STEP 7
    '
    ' A:Hのセル内容だけを削除
    '
    ' セル書式や他の列のデータには触れない。
    '====================================================

    ws.Range( _
        ws.Cells(lastRow, 1), _
        ws.Cells(lastRow, 8) _
    ).ClearContents


    Exit Sub


ErrHandler:

    MsgBox _
        "トレード削除中にエラーが発生しました。" & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical

End Sub


'========================================================
' 8. Legサイズの表示
'
' 空欄の場合は "-" と表示
'========================================================

Private Function DisplayLegValue( _
    ByVal v As Variant _
) As String

    If IsError(v) Then

        DisplayLegValue = "(Error)"

    ElseIf IsEmpty(v) Then

        DisplayLegValue = "-"

    ElseIf IsNull(v) Then

        DisplayLegValue = "-"

    ElseIf Len(Trim$(CStr(v))) = 0 Then

        DisplayLegValue = "-"

    Else

        DisplayLegValue = CStr(v)

    End If

End Function


'========================================================
' 9. TradeHisotryシートの取得
'
' createIfMissing = True:
'   存在しなければ作成
'
' createIfMissing = False:
'   存在しなければNothingを返す
'========================================================

Private Function GetTradeHistorySheet( _
    ByVal createIfMissing As Boolean _
) As Worksheet

    Dim ws As Worksheet


    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(HISTORY_SHEET)

    On Error GoTo 0


    If ws Is Nothing Then

        If createIfMissing Then

            Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.count))

            ws.Name = HISTORY_SHEET

        End If

    End If


    Set GetTradeHistorySheet = ws

End Function


'========================================================
' 10. ヘッダー初期化
'
' 新形式：
'
' Date, Time, Trade, Level,
' Leg1, Leg2, Leg3, TradeType
'
' 旧7列形式：
'
' Date, Time, Trade, Level,
' Leg1, Leg2, TradeType
'
' 旧6列形式：
'
' Date, Time, Trade, Level,
' Size, TradeType
'
'========================================================

Private Sub InitializeTradeHistory( _
    ByVal ws As Worksheet _
)

    Dim headers As Variant

    Dim i As Long

    Dim lastRow As Long


    '====================================================
    ' 新形式なら何もしない
    '====================================================

    If HasTradeHistoryHeaders(ws) Then

        Exit Sub

    End If


    '====================================================
    ' 旧7列形式からの移行
    '
    ' G列のTradeTypeをH列へ移動
    ' G列をLeg3として使用する。
    '====================================================

    If HasSevenColumnHeaders(ws) Then


        '--------------------------------------------
        ' H列に既存データがある場合は移行しない
        '--------------------------------------------

        If ColumnHasData(ws, 8) Then

            Err.Raise vbObjectError + 710, , _
                "H列に既存データが存在するため、" & _
                "7列形式からの自動移行を中止しました。"

        End If


        lastRow = GetLastHistoryRow(ws)


        '--------------------------------------------
        ' G列のTradeTypeをH列へ移動
        '--------------------------------------------

        ws.Range("H1:H" & CStr(lastRow)).value2 = _
            ws.Range("G1:G" & CStr(lastRow)).value2


        '--------------------------------------------
        ' G列をLeg3として使用
        '--------------------------------------------

        ws.Range("G1:G" & CStr(lastRow)).ClearContents


        '--------------------------------------------
        ' 新ヘッダー
        '--------------------------------------------

        ws.Cells(1, 7).value2 = "Leg3"

        ws.Cells(1, 8).value2 = "TradeType"


        Exit Sub

    End If


    '====================================================
    ' 旧6列形式からの移行
    '
    ' E列のSizeをLeg1として使用
    '
    ' F列のTradeTypeをH列へ移動
    '
    ' F列をLeg2、G列をLeg3として使用
    '====================================================

    If HasSixColumnHeaders(ws) Then


        '--------------------------------------------
        ' G列・H列に既存データがある場合は中止
        '--------------------------------------------

        If ColumnHasData(ws, 7) Or _
           ColumnHasData(ws, 8) Then

            Err.Raise vbObjectError + 711, , _
                "G列またはH列に既存データが存在するため、" & _
                "6列形式からの自動移行を中止しました。"

        End If


        lastRow = GetLastHistoryRow(ws)


        '--------------------------------------------
        ' F列のTradeTypeをH列へ移動
        '--------------------------------------------

        ws.Range("H1:H" & CStr(lastRow)).value2 = _
            ws.Range("F1:F" & CStr(lastRow)).value2


        '--------------------------------------------
        ' F列をLeg2として使用
        '--------------------------------------------

        ws.Range("F1:F" & CStr(lastRow)).ClearContents


        '--------------------------------------------
        ' 新ヘッダー
        '--------------------------------------------

        ws.Cells(1, 5).value2 = "Leg1"

        ws.Cells(1, 6).value2 = "Leg2"

        ws.Cells(1, 7).value2 = "Leg3"

        ws.Cells(1, 8).value2 = "TradeType"


        Exit Sub

    End If


    '====================================================
    ' 新規シート
    '
    ' A:Hが空欄の場合のみヘッダーを作成
    '
    ' 既存データがあれば上書きせず中止する。
    '====================================================

    If Application.WorksheetFunction.CountA( _
        ws.Range("A:H")) > 0 Then

        Err.Raise vbObjectError + 712, , _
            "TradeHisotryのヘッダーが正しくありません。" & _
            vbCrLf & _
            "既存データを保護するため処理を中止しました。"

    End If


    '====================================================
    ' 8列形式のヘッダーを作成
    '====================================================

    headers = Array( _
        "Date", _
        "Time", _
        "Trade", _
        "Level", _
        "Leg1", _
        "Leg2", _
        "Leg3", _
        "TradeType")


    For i = LBound(headers) To UBound(headers)

        ws.Cells(1, i + 1).value2 = headers(i)

    Next i

End Sub


'========================================================
' 11. 新8列形式のヘッダー確認
'========================================================

Private Function HasTradeHistoryHeaders( _
    ByVal ws As Worksheet _
) As Boolean

    HasTradeHistoryHeaders = _
        HeadersMatch(ws, Array( _
            "Date", _
            "Time", _
            "Trade", _
            "Level", _
            "Leg1", _
            "Leg2", _
            "Leg3", _
            "TradeType"))

End Function


'========================================================
' 12. 旧7列形式のヘッダー確認
'========================================================

Private Function HasSevenColumnHeaders( _
    ByVal ws As Worksheet _
) As Boolean

    HasSevenColumnHeaders = _
        HeadersMatch(ws, Array( _
            "Date", _
            "Time", _
            "Trade", _
            "Level", _
            "Leg1", _
            "Leg2", _
            "TradeType"))

End Function


'========================================================
' 13. 旧6列形式のヘッダー確認
'========================================================

Private Function HasSixColumnHeaders( _
    ByVal ws As Worksheet _
) As Boolean

    HasSixColumnHeaders = _
        HeadersMatch(ws, Array( _
            "Date", _
            "Time", _
            "Trade", _
            "Level", _
            "Size", _
            "TradeType"))

End Function


'========================================================
' 14. ヘッダー比較
'========================================================

Private Function HeadersMatch( _
    ByVal ws As Worksheet, _
    ByVal expectedHeaders As Variant _
) As Boolean

    Dim i As Long

    Dim v As Variant


    HeadersMatch = False


    For i = LBound(expectedHeaders) To UBound(expectedHeaders)

        v = ws.Cells(1, i + 1).value2


        If IsError(v) Then

            Exit Function

        End If


        If StrComp( _
            Trim$(CStr(v)), _
            CStr(expectedHeaders(i)), _
            vbTextCompare) <> 0 Then

            Exit Function

        End If

    Next i


    HeadersMatch = True

End Function


'========================================================
' 15. 指定列の既存データ確認
'
' 書式のみのセルは対象としない。
'
'========================================================

Private Function ColumnHasData( _
    ByVal ws As Worksheet, _
    ByVal colNum As Long _
) As Boolean

    Dim foundCell As Range


    Set foundCell = ws.Columns(colNum).Find( _
        What:="*", _
        After:=ws.Cells(ws.rows.count, colNum), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False, _
        SearchFormat:=False)


    ColumnHasData = Not foundCell Is Nothing

End Function


'========================================================
' 16. 最終入力行
'
' A:Hの各列から最大の最終行を取得する。
'
'========================================================

Private Function GetLastHistoryRow( _
    ByVal ws As Worksheet _
) As Long

    Dim c As Long

    Dim currentLastRow As Long
    Dim maxLastRow As Long


    maxLastRow = 1


    For c = 1 To 8

        currentLastRow = ws.Cells( _
            ws.rows.count, c).End(xlUp).Row


        If currentLastRow > maxLastRow Then

            maxLastRow = currentLastRow

        End If

    Next c


    GetLastHistoryRow = maxLastRow

End Function


'========================================================
' 17. 最終トレードの有効性確認
'
' Date、Time、Trade、Level、Leg1、TradeType
' の存在を確認する。
'
' Leg2・Leg3は空欄を許可する。
'
'========================================================

Private Function IsValidTradeRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long _
) As Boolean

    Dim c As Long

    Dim v As Variant


    IsValidTradeRow = False


    '====================================================
    ' Excelエラー値を確認
    '====================================================

    For c = 1 To 8

        If IsError(ws.Cells(rowNum, c).value2) Then

            Exit Function

        End If

    Next c


    '====================================================
    ' Date
    '====================================================

    v = ws.Cells(rowNum, 1).value2

    If Not IsNumeric(v) Then

        Exit Function

    End If


    '====================================================
    ' Time
    '====================================================

    v = ws.Cells(rowNum, 2).value2

    If Not IsNumeric(v) Then

        Exit Function

    End If


    '====================================================
    ' Trade
    '====================================================

    v = ws.Cells(rowNum, 3).value2

    If Len(Trim$(CStr(v))) = 0 Then

        Exit Function

    End If


    '====================================================
    ' Level
    '====================================================

    v = ws.Cells(rowNum, 4).value2

    If Not IsNumeric(v) Then

        Exit Function

    End If


    '====================================================
    ' Leg1
    '====================================================

    v = ws.Cells(rowNum, 5).value2

    If Not IsNumeric(v) Then

        Exit Function

    End If


    '====================================================
    ' Leg2
    '
    ' 空欄または数値
    '====================================================

    v = ws.Cells(rowNum, 6).value2

    If Len(Trim$(CStr(v))) > 0 Then

        If Not IsNumeric(v) Then

            Exit Function

        End If

    End If


    '====================================================
    ' Leg3
    '
    ' 空欄または数値
    '====================================================

    v = ws.Cells(rowNum, 7).value2

    If Len(Trim$(CStr(v))) > 0 Then

        If Not IsNumeric(v) Then

            Exit Function

        End If

    End If


    '====================================================
    ' TradeType
    '
    ' FlowまたはProp
    '====================================================

    v = ws.Cells(rowNum, 8).value2

    If Len(Trim$(CStr(v))) = 0 Then

        Exit Function

    End If


    IsValidTradeRow = True

End Function


