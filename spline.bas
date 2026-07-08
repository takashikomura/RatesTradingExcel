Option Explicit

' ============================================================
' 1. Public UDF: CubicSpline
' ============================================================

Public Function CubicSpline( _
    ByVal Known_X As Range, _
    ByVal Known_Y As Range, _
    ByVal Target_X As Variant, _
    Optional ByVal AllowExtrapolate As Boolean = False, _
    Optional ByVal ShowErrorMessage As Boolean = False _
) As Variant

    On Error GoTo ErrHandler

    Dim n As Long
    Dim i As Long
    Dim x() As Double
    Dim y() As Double
    Dim xq As Double

    If Not IsSingleVectorRange(Known_X) Or Not IsSingleVectorRange(Known_Y) Then
        CubicSpline = SplineError(xlErrValue, _
            "Known_X と Known_Y は1行または1列の範囲で指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    If Known_X.Cells.CountLarge <> Known_Y.Cells.CountLarge Then
        CubicSpline = SplineError(xlErrValue, _
            "Known_X と Known_Y のデータ数が一致していません。", _
            ShowErrorMessage)
        Exit Function
    End If

    n = CLng(Known_X.Cells.CountLarge)

    If n < 3 Then
        CubicSpline = SplineError(xlErrValue, _
            "3次スプラインには3点以上のデータが必要です。", _
            ShowErrorMessage)
        Exit Function
    End If

    If IsError(Target_X) Or IsEmpty(Target_X) Or Not IsNumeric(Target_X) Then
        CubicSpline = SplineError(xlErrValue, _
            "Target_X が数値ではありません。", _
            ShowErrorMessage)
        Exit Function
    End If

    xq = CDbl(Target_X)

    ReDim x(1 To n)
    ReDim y(1 To n)

    For i = 1 To n

        If IsError(Known_X.Cells(i).Value2) _
            Or IsEmpty(Known_X.Cells(i).Value2) _
            Or Not IsNumeric(Known_X.Cells(i).Value2) Then

            CubicSpline = SplineError(xlErrValue, _
                "Known_X の " & i & " 番目が数値ではありません。", _
                ShowErrorMessage)
            Exit Function
        End If

        If IsError(Known_Y.Cells(i).Value2) _
            Or IsEmpty(Known_Y.Cells(i).Value2) _
            Or Not IsNumeric(Known_Y.Cells(i).Value2) Then

            CubicSpline = SplineError(xlErrValue, _
                "Known_Y の " & i & " 番目が数値ではありません。", _
                ShowErrorMessage)
            Exit Function
        End If

        x(i) = CDbl(Known_X.Cells(i).Value2)
        y(i) = CDbl(Known_Y.Cells(i).Value2)

    Next i

    Call SortXYAscending(x, y, n)

    CubicSpline = NaturalCubicSplineFromArrays( _
        x, y, n, xq, AllowExtrapolate, ShowErrorMessage)

    Exit Function

ErrHandler:
    CubicSpline = SplineError(xlErrValue, _
        "CubicSpline 実行中のVBAエラー: " & Err.Description, _
        ShowErrorMessage)

End Function


' ============================================================
' 2. Public UDF: MofSpline
'
' 既存の4列範囲入力版。
'
' 使用例:
'   =MofSpline(Ticker列, 満期列, クーポン列, 複利利回り列, 10, 基準日, FALSE, TRUE)
' ============================================================

Public Function MofSpline( _
    ByVal TickerRange As Range, _
    ByVal MaturityRange As Range, _
    ByVal CouponRange As Range, _
    ByVal CompoundYieldRange As Range, _
    ByVal TargetMaturity As Variant, _
    Optional ByVal AsOfDate As Variant, _
    Optional ByVal AllowExtrapolate As Boolean = False, _
    Optional ByVal ShowErrorMessage As Boolean = False _
) As Variant

    On Error GoTo ErrHandler

    Application.Volatile True

    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim j As Long
    Dim grid As Long
    Dim termCls As Long
    Dim innerIdx As Long
    Dim outerIdx As Long

    Dim xq As Double
    Dim AsOfSerial As Double
    Dim ok As Boolean

    Dim tickers() As String
    Dim issueNo() As Long
    Dim originalTerm() As Long
    Dim matYears() As Double
    Dim coupon() As Double
    Dim compYield() As Double
    Dim selected() As Boolean

    Dim x() As Double
    Dim y() As Double
    Dim nodeCoupon() As Double
    Dim nodeIssue() As Long

    ' -----------------------------
    ' 入力範囲チェック
    ' -----------------------------

    If Not IsSingleVectorRange(TickerRange) _
        Or Not IsSingleVectorRange(MaturityRange) _
        Or Not IsSingleVectorRange(CouponRange) _
        Or Not IsSingleVectorRange(CompoundYieldRange) Then

        MofSpline = SplineError(xlErrValue, _
            "Ticker列、満期列、クーポン列、複利列は、すべて1行または1列の範囲で指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    n = CLng(TickerRange.Cells.CountLarge)

    If MaturityRange.Cells.CountLarge <> n _
        Or CouponRange.Cells.CountLarge <> n _
        Or CompoundYieldRange.Cells.CountLarge <> n Then

        MofSpline = SplineError(xlErrValue, _
            "Ticker列、満期列、クーポン列、複利列のデータ数が一致していません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If n < 3 Then
        MofSpline = SplineError(xlErrValue, _
            "3次スプラインには最低3銘柄以上が必要です。", _
            ShowErrorMessage)
        Exit Function
    End If

    ' -----------------------------
    ' 基準日処理
    ' -----------------------------

    If IsMissing(AsOfDate) Then
        AsOfSerial = CDbl(Date)
    Else
        AsOfSerial = DateValueToSerial(AsOfDate, ok)

        If Not ok Then
            MofSpline = SplineError(xlErrValue, _
                "AsOfDate が日付として認識できません。", _
                ShowErrorMessage)
            Exit Function
        End If
    End If

    ' -----------------------------
    ' TargetMaturity処理
    '
    ' 1000以下の数値 → 年限として扱う
    ' 1000超の数値   → Excel日付シリアルとして扱う
    ' 文字列日付      → 日付として扱う
    ' -----------------------------

    xq = MaturityToYears(TargetMaturity, AsOfSerial, ok)

    If Not ok Then
        MofSpline = SplineError(xlErrValue, _
            "TargetMaturity が年限または日付として認識できません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If xq <= 0# Then
        MofSpline = SplineError(xlErrValue, _
            "TargetMaturity は正の年限、またはAsOfDateより後の日付で指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    ' -----------------------------
    ' 配列確保
    ' -----------------------------

    ReDim tickers(1 To n)
    ReDim issueNo(1 To n)
    ReDim originalTerm(1 To n)
    ReDim matYears(1 To n)
    ReDim coupon(1 To n)
    ReDim compYield(1 To n)
    ReDim selected(1 To n)

    ' -----------------------------
    ' 入力データ読み込み
    ' -----------------------------

    For i = 1 To n

        If IsError(TickerRange.Cells(i).Value2) _
            Or IsEmpty(TickerRange.Cells(i).Value2) Then

            MofSpline = SplineError(xlErrValue, _
                "Ticker列の " & i & " 番目が空欄またはエラーです。", _
                ShowErrorMessage)
            Exit Function
        End If

        tickers(i) = UCase$(Trim$(CStr(TickerRange.Cells(i).Value2)))

        originalTerm(i) = MofTermClassFromTicker(tickers(i))

        If originalTerm(i) = 0 Then
            MofSpline = SplineError(xlErrValue, _
                "Ticker """ & tickers(i) & """ の年限区分を判定できません。JN/JS/JB/JL/JX/JU のいずれかで始まるTickerを指定してください。", _
                ShowErrorMessage)
            Exit Function
        End If

        issueNo(i) = IssueNumberFromTicker(tickers(i))

        If issueNo(i) < 0 Then
            MofSpline = SplineError(xlErrValue, _
                "Ticker """ & tickers(i) & """ から回号を取得できません。", _
                ShowErrorMessage)
            Exit Function
        End If

        matYears(i) = MaturityToYears(MaturityRange.Cells(i).Value2, AsOfSerial, ok)

        If Not ok Then
            MofSpline = SplineError(xlErrValue, _
                "満期列の " & i & " 番目が日付または残存年数として認識できません。", _
                ShowErrorMessage)
            Exit Function
        End If

        If matYears(i) <= 0# Then
            MofSpline = SplineError(xlErrValue, _
                "Ticker """ & tickers(i) & """ は基準日時点で満期到来済み、または残存年数が0以下です。", _
                ShowErrorMessage)
            Exit Function
        End If

        If IsError(CouponRange.Cells(i).Value2) _
            Or IsEmpty(CouponRange.Cells(i).Value2) _
            Or Not IsNumeric(CouponRange.Cells(i).Value2) Then

            MofSpline = SplineError(xlErrValue, _
                "クーポン列の " & i & " 番目が数値ではありません。", _
                ShowErrorMessage)
            Exit Function
        End If

        If IsError(CompoundYieldRange.Cells(i).Value2) _
            Or IsEmpty(CompoundYieldRange.Cells(i).Value2) _
            Or Not IsNumeric(CompoundYieldRange.Cells(i).Value2) Then

            MofSpline = SplineError(xlErrValue, _
                "複利列の " & i & " 番目が数値ではありません。", _
                ShowErrorMessage)
            Exit Function
        End If

        coupon(i) = CDbl(CouponRange.Cells(i).Value2)
        compYield(i) = CDbl(CompoundYieldRange.Cells(i).Value2)
        selected(i) = False

    Next i

    ' -----------------------------
    ' 財務省方式に近い対象銘柄選定
    '
    ' 1. 各年限区分のカレント銘柄、最大回号銘柄を選定
    ' 2. 1年～40年の各グリッドについて、
    '    対応する年限区分の内側・外側の最近傍銘柄を選定
    ' 3. 同一残存年数ならクーポン高、さらに同じなら回号大を優先
    ' -----------------------------

    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 2)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 5)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 10)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 20)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 30)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 40)

    For grid = 1 To 40

        termCls = MofTermClassForGrid(grid)

        innerIdx = FindNearestIssue( _
            originalTerm, matYears, coupon, issueNo, _
            n, termCls, CDbl(grid), -1)

        outerIdx = FindNearestIssue( _
            originalTerm, matYears, coupon, issueNo, _
            n, termCls, CDbl(grid), 1)

        If innerIdx > 0 Then selected(innerIdx) = True
        If outerIdx > 0 Then selected(outerIdx) = True

    Next grid

    ' -----------------------------
    ' 選定銘柄をスプラインノードへ変換
    ' -----------------------------

    m = 0

    For i = 1 To n
        If selected(i) Then m = m + 1
    Next i

    If m < 3 Then
        MofSpline = SplineError(xlErrValue, _
            "選定後の銘柄数が3未満です。対象データが不足しています。", _
            ShowErrorMessage)
        Exit Function
    End If

    ReDim x(1 To m)
    ReDim y(1 To m)
    ReDim nodeCoupon(1 To m)
    ReDim nodeIssue(1 To m)

    j = 0

    For i = 1 To n

        If selected(i) Then
            j = j + 1

            x(j) = matYears(i)
            y(j) = compYield(i)
            nodeCoupon(j) = coupon(i)
            nodeIssue(j) = issueNo(i)
        End If

    Next i

    Call SortSplineNodes(x, y, nodeCoupon, nodeIssue, m)

    m = CompressDuplicateMaturities(x, y, nodeCoupon, nodeIssue, m)

    If m < 3 Then
        MofSpline = SplineError(xlErrValue, _
            "重複残存年数を圧縮した結果、スプラインノードが3点未満になりました。", _
            ShowErrorMessage)
        Exit Function
    End If

    MofSpline = NaturalCubicSplineFromArrays( _
        x, y, m, xq, AllowExtrapolate, ShowErrorMessage)

    Exit Function

ErrHandler:
    MofSpline = SplineError(xlErrValue, _
        "MofSpline 実行中のVBAエラー: " & Err.Description, _
        ShowErrorMessage)

End Function


' ============================================================
' 3. Public UDF: MOFSplineTable
'
' ヘッダー付きの名前付き範囲・通常範囲を直接受け取る版。
'
' 使用例:
'   =MOFSplineTable(pastetable,10,$C$3,FALSE,TRUE)
'
' DataRange:
'   ヘッダー行を含む範囲。
'
' 必須カラム:
'   Ticker
'   Maturity
'   Cpn(%) / Coupon
'   CY(%) / CompoundYield
'
' 不完全な行は無視する。
' ============================================================

Public Function MOFSplineTable( _
    ByVal DataRange As Range, _
    ByVal TargetMaturity As Variant, _
    Optional ByVal AsOfDate As Variant, _
    Optional ByVal AllowExtrapolate As Boolean = False, _
    Optional ByVal ShowErrorMessage As Boolean = False, _
    Optional ByVal HeaderRowIndex As Long = 1 _
) As Variant

    On Error GoTo ErrHandler

    Application.Volatile True

    If DataRange Is Nothing Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange が指定されていません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If DataRange.Areas.count <> 1 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange は単一エリアで指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    If DataRange.rows.count < 2 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange にはヘッダー行とデータ行が必要です。", _
            ShowErrorMessage)
        Exit Function
    End If

    If HeaderRowIndex < 1 Or HeaderRowIndex >= DataRange.rows.count Then
        MOFSplineTable = SplineError(xlErrValue, _
            "HeaderRowIndex が不正です。ヘッダー行とデータ行を含む範囲を指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    Dim idxTicker As Long
    Dim idxMaturity As Long
    Dim idxCoupon As Long
    Dim idxYield As Long

    idxTicker = MNB_FindColumnIndex( _
        DataRange, HeaderRowIndex, _
        "Ticker", "TickerRange", "銘柄", "銘柄コード", "ティッカー", "コード")

    idxMaturity = MNB_FindColumnIndex( _
        DataRange, HeaderRowIndex, _
        "Maturity", "MaturityRange", "MaturityDate", _
        "満期", "満期日", "償還日", "償還年月日", "残存年数")

    idxCoupon = MNB_FindColumnIndex( _
        DataRange, HeaderRowIndex, _
        "Coupon", "CouponRange", "CouponRate", _
        "Cpn", "Cpn%", "Cpn(%)", _
        "クーポン", "表面利率", "利率", "利率%")

    idxYield = MNB_FindColumnIndex( _
        DataRange, HeaderRowIndex, _
        "CompoundYield", "CompoundYieldRange", "CompoundYieldYtm", _
        "CY", "CY%", "CY(%)", _
        "複利", "複利利回り", "複利利回", "複利利回り%")

    If idxTicker = 0 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange 内に Ticker 列が見つかりません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If idxMaturity = 0 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange 内に Maturity 列が見つかりません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If idxCoupon = 0 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange 内に Coupon / Cpn(%) 列が見つかりません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If idxYield = 0 Then
        MOFSplineTable = SplineError(xlErrValue, _
            "DataRange 内に CompoundYield / CY(%) 列が見つかりません。", _
            ShowErrorMessage)
        Exit Function
    End If

    Dim AsOfValue As Variant

    If IsMissing(AsOfDate) Then
        AsOfValue = Empty
    Else
        AsOfValue = AsOfDate
    End If

    MOFSplineTable = MNB_MofSplineCoreFromBlock( _
        DataRange, _
        HeaderRowIndex, _
        idxTicker, _
        idxMaturity, _
        idxCoupon, _
        idxYield, _
        TargetMaturity, _
        AsOfValue, _
        AllowExtrapolate, _
        ShowErrorMessage)

    Exit Function

ErrHandler:
    MOFSplineTable = SplineError(xlErrValue, _
        "MOFSplineTable 実行中のVBAエラー: " & Err.Description, _
        ShowErrorMessage)

End Function


' ============================================================
' 4. Optional: UDF help registration
' ============================================================

Public Sub RegisterSplineUdfHelp()

    Application.MacroOptions _
        Macro:="CubicSpline", _
        Description:="既知のX/YからNatural Cubic Splineで任意のXに対応するYを補間します。Xは内部で昇順ソートされます。", _
        Category:="JGB Analytics", _
        ArgumentDescriptions:=Array( _
            "既知のX値の範囲です。1行または1列で指定してください。", _
            "既知のY値の範囲です。Known_Xと同じデータ数にしてください。", _
            "補間したいX値です。", _
            "省略可能。TRUEなら範囲外の外挿を許可します。省略時はFALSEです。", _
            "省略可能。TRUEならExcelエラーではなくエラー原因を文字列で返します。" _
        )

    Application.MacroOptions _
        Macro:="MofSpline", _
        Description:="財務省の国債金利情報の公開算出方法に近い形で、Ticker・満期・クーポン・複利利回りから対象銘柄を選定し、3次スプラインで任意年限のconstant maturity金利を返します。TargetMaturityは年限または日付で指定できます。", _
        Category:="JGB Analytics", _
        ArgumentDescriptions:=Array( _
            "Ticker列。例: JN480, JS176, JB380, JL190, JX085, JU018", _
            "満期列。Excel日付、文字列日付、または残存年数を指定します。", _
            "クーポン列。同一残存年数の銘柄選定時のタイブレークに使います。", _
            "半年複利利回り列。単利ではなく複利利回りを指定してください。", _
            "補間したい年限または日付。例: 10, 15.5, DATE(2035,9,20)", _
            "省略可能。基準日。省略時は今日を使用します。", _
            "省略可能。TRUEなら外挿を許可します。省略時はFALSEです。", _
            "省略可能。TRUEならExcelエラーではなくエラー原因を文字列で返します。" _
        )

    Application.MacroOptions _
        Macro:="MOFSplineTable", _
        Description:="ヘッダー付き範囲からTicker・Maturity・Cpn/Coupon・CY/CompoundYield列を自動検出し、不完全行を無視してMOF風スプラインを計算します。", _
        Category:="JGB Analytics", _
        ArgumentDescriptions:=Array( _
            "ヘッダー行を含むデータ範囲です。名前付き範囲 pastetable などを指定できます。", _
            "補間したい年限または日付。例: 10, 15.5, DATE(2035,9,20)", _
            "省略可能。基準日。省略時は今日を使用します。", _
            "省略可能。TRUEなら外挿を許可します。省略時はFALSEです。", _
            "省略可能。TRUEならExcelエラーではなくエラー原因を文字列で返します。", _
            "省略可能。ヘッダー行が範囲内の何行目かを指定します。省略時は1です。" _
        )

End Sub


' ============================================================
' 5. Shared helpers
' ============================================================

Private Function SplineError( _
    ByVal ExcelErrorCode As Long, _
    ByVal Message As String, _
    ByVal ShowErrorMessage As Boolean _
) As Variant

    If ShowErrorMessage Then
        SplineError = Message
    Else
        SplineError = CVErr(ExcelErrorCode)
    End If

End Function


Private Function IsSingleVectorRange(ByVal rng As Range) As Boolean

    If rng.Areas.count <> 1 Then
        IsSingleVectorRange = False
    ElseIf rng.rows.count = 1 Or rng.Columns.count = 1 Then
        IsSingleVectorRange = True
    Else
        IsSingleVectorRange = False
    End If

End Function


Private Sub SortXYAscending( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal n As Long _
)

    Dim i As Long
    Dim j As Long
    Dim keyX As Double
    Dim keyY As Double

    For i = 2 To n

        keyX = x(i)
        keyY = y(i)
        j = i - 1

        Do While j >= 1 And x(j) > keyX
            x(j + 1) = x(j)
            y(j + 1) = y(j)
            j = j - 1
        Loop

        x(j + 1) = keyX
        y(j + 1) = keyY

    Next i

End Sub


Private Function NaturalCubicSplineFromArrays( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByVal n As Long, _
    ByVal Target_X As Double, _
    Optional ByVal AllowExtrapolate As Boolean = False, _
    Optional ByVal ShowErrorMessage As Boolean = False _
) As Variant

    On Error GoTo ErrHandler

    Dim i As Long
    Dim j As Long
    Dim idx As Long

    Dim h() As Double
    Dim alpha() As Double
    Dim l() As Double
    Dim mu() As Double
    Dim z() As Double
    Dim c() As Double
    Dim b() As Double
    Dim d() As Double

    Dim dx As Double

    If n < 3 Then
        NaturalCubicSplineFromArrays = SplineError(xlErrValue, _
            "3次スプラインには3点以上のノードが必要です。", _
            ShowErrorMessage)
        Exit Function
    End If

    ReDim h(1 To n - 1)
    ReDim alpha(1 To n)
    ReDim l(1 To n)
    ReDim mu(1 To n)
    ReDim z(1 To n)
    ReDim c(1 To n)
    ReDim b(1 To n - 1)
    ReDim d(1 To n - 1)

    For i = 1 To n - 1

        h(i) = x(i + 1) - x(i)

        If h(i) <= 0# Then
            NaturalCubicSplineFromArrays = SplineError(xlErrValue, _
                "Xは昇順かつ重複なしである必要があります。", _
                ShowErrorMessage)
            Exit Function
        End If

    Next i

    If Not AllowExtrapolate Then
        If Target_X < x(1) Or Target_X > x(n) Then
            NaturalCubicSplineFromArrays = SplineError(xlErrNA, _
                "補間対象がスプラインノードの範囲外です。", _
                ShowErrorMessage)
            Exit Function
        End If
    End If

    For i = 2 To n - 1

        alpha(i) = _
            (3# / h(i)) * (y(i + 1) - y(i)) _
            - (3# / h(i - 1)) * (y(i) - y(i - 1))

    Next i

    l(1) = 1#
    mu(1) = 0#
    z(1) = 0#

    For i = 2 To n - 1

        l(i) = 2# * (x(i + 1) - x(i - 1)) - h(i - 1) * mu(i - 1)

        If Abs(l(i)) < 0.000000000001 Then
            NaturalCubicSplineFromArrays = SplineError(xlErrDiv0, _
                "スプライン係数計算中にゼロ除算に近い状態が発生しました。", _
                ShowErrorMessage)
            Exit Function
        End If

        mu(i) = h(i) / l(i)
        z(i) = (alpha(i) - h(i - 1) * z(i - 1)) / l(i)

    Next i

    l(n) = 1#
    z(n) = 0#
    c(n) = 0#

    For j = n - 1 To 1 Step -1

        c(j) = z(j) - mu(j) * c(j + 1)

        b(j) = (y(j + 1) - y(j)) / h(j) _
            - h(j) * (c(j + 1) + 2# * c(j)) / 3#

        d(j) = (c(j + 1) - c(j)) / (3# * h(j))

    Next j

    If Target_X <= x(1) Then
        idx = 1
    ElseIf Target_X >= x(n) Then
        idx = n - 1
    Else
        idx = FindSplineIntervalArray(x, n, Target_X)
    End If

    dx = Target_X - x(idx)

    NaturalCubicSplineFromArrays = y(idx) _
        + b(idx) * dx _
        + c(idx) * dx ^ 2 _
        + d(idx) * dx ^ 3

    Exit Function

ErrHandler:
    NaturalCubicSplineFromArrays = SplineError(xlErrValue, _
        "スプライン計算中のVBAエラー: " & Err.Description, _
        ShowErrorMessage)

End Function


Private Function FindSplineIntervalArray( _
    ByRef x() As Double, _
    ByVal n As Long, _
    ByVal Target_X As Double _
) As Long

    Dim lo As Long
    Dim hi As Long
    Dim mid As Long

    lo = 1
    hi = n

    Do While hi - lo > 1

        mid = (lo + hi) \ 2

        If Target_X >= x(mid) Then
            lo = mid
        Else
            hi = mid
        End If

    Loop

    If lo >= n Then lo = n - 1

    FindSplineIntervalArray = lo

End Function


' ============================================================
' 6. Date and maturity helpers
' ============================================================

Private Function DateValueToSerial( _
    ByVal DateValue As Variant, _
    ByRef ok As Boolean _
) As Double

    ok = False

    If IsError(DateValue) Then Exit Function

    If IsEmpty(DateValue) Then
        DateValueToSerial = CDbl(Date)
        ok = True
        Exit Function
    End If

    If IsNumeric(DateValue) Then
        DateValueToSerial = CDbl(DateValue)
        ok = True
        Exit Function
    End If

    If IsDate(DateValue) Then
        DateValueToSerial = CDbl(CDate(DateValue))
        ok = True
        Exit Function
    End If

End Function


Private Function MaturityToYears( _
    ByVal MaturityValue As Variant, _
    ByVal AsOfSerial As Double, _
    ByRef ok As Boolean _
) As Double

    Dim v As Double
    Dim dt As Date

    ok = False

    If IsError(MaturityValue) Or IsEmpty(MaturityValue) Then Exit Function

    If IsNumeric(MaturityValue) Then

        v = CDbl(MaturityValue)

        If v > 1000# Then
            MaturityToYears = (v - AsOfSerial) / 365.25
        Else
            MaturityToYears = v
        End If

        ok = True
        Exit Function

    End If

    If IsDate(MaturityValue) Then
        dt = CDate(MaturityValue)
        MaturityToYears = (CDbl(dt) - AsOfSerial) / 365.25
        ok = True
        Exit Function
    End If

End Function


' ============================================================
' 7. Ticker helpers
' ============================================================

Private Function MofTermClassFromTicker(ByVal ticker As String) As Long

    Dim s As String
    Dim p As String

    s = UCase$(Trim$(ticker))
    p = TickerPrefix(s)

    Select Case p

        Case "JN"
            MofTermClassFromTicker = 2

        Case "JS"
            MofTermClassFromTicker = 5

        Case "JB"
            MofTermClassFromTicker = 10

        Case "JL"
            MofTermClassFromTicker = 20

        Case "JX"
            MofTermClassFromTicker = 30

        Case "JU"
            MofTermClassFromTicker = 40

        Case Else
            MofTermClassFromTicker = 0

    End Select

End Function


Private Function TickerPrefix(ByVal ticker As String) As String

    Dim i As Long

    For i = 1 To Len(ticker)

        If mid$(ticker, i, 1) Like "#" Then
            TickerPrefix = Left$(ticker, i - 1)
            Exit Function
        End If

    Next i

    TickerPrefix = ticker

End Function


Private Function IssueNumberFromTicker(ByVal ticker As String) As Long

    Dim i As Long
    Dim ch As String
    Dim digits As String

    digits = ""

    For i = Len(ticker) To 1 Step -1

        ch = mid$(ticker, i, 1)

        If ch Like "#" Then
            digits = ch & digits
        ElseIf Len(digits) > 0 Then
            Exit For
        End If

    Next i

    If Len(digits) = 0 Then
        IssueNumberFromTicker = -1
    Else
        IssueNumberFromTicker = CLng(digits)
    End If

End Function


' ============================================================
' 8. MOF grid mapping
' ============================================================

Private Function MofTermClassForGrid(ByVal GridYear As Long) As Long

    If GridYear <= 2 Then
        MofTermClassForGrid = 2
    ElseIf GridYear <= 5 Then
        MofTermClassForGrid = 5
    ElseIf GridYear <= 10 Then
        MofTermClassForGrid = 10
    ElseIf GridYear <= 20 Then
        MofTermClassForGrid = 20
    ElseIf GridYear <= 30 Then
        MofTermClassForGrid = 30
    Else
        MofTermClassForGrid = 40
    End If

End Function


' ============================================================
' 9. MOF issue selection helpers
' ============================================================

Private Sub SelectCurrentIssue( _
    ByRef selected() As Boolean, _
    ByRef originalTerm() As Long, _
    ByRef issueNo() As Long, _
    ByVal n As Long, _
    ByVal TargetTermClass As Long _
)

    Dim bestIdx As Long

    bestIdx = CurrentIssueIndex(originalTerm, issueNo, n, TargetTermClass)

    If bestIdx > 0 Then selected(bestIdx) = True

End Sub


Private Function CurrentIssueIndex( _
    ByRef originalTerm() As Long, _
    ByRef issueNo() As Long, _
    ByVal n As Long, _
    ByVal TargetTermClass As Long _
) As Long

    Dim i As Long
    Dim bestIdx As Long

    bestIdx = 0

    For i = 1 To n

        If originalTerm(i) = TargetTermClass Then

            If bestIdx = 0 Then
                bestIdx = i
            ElseIf issueNo(i) > issueNo(bestIdx) Then
                bestIdx = i
            End If

        End If

    Next i

    CurrentIssueIndex = bestIdx

End Function


Private Function FindNearestIssue( _
    ByRef originalTerm() As Long, _
    ByRef matYears() As Double, _
    ByRef coupon() As Double, _
    ByRef issueNo() As Long, _
    ByVal n As Long, _
    ByVal TargetTermClass As Long, _
    ByVal GridYear As Double, _
    ByVal Side As Long _
) As Long

    ' Side = -1 : 内側。残存年数 <= GridYear の最近傍
    ' Side =  1 : 外側。残存年数 >= GridYear の最近傍

    Dim i As Long
    Dim bestIdx As Long
    Dim candidateOK As Boolean

    bestIdx = 0

    For i = 1 To n

        If originalTerm(i) = TargetTermClass Then

            candidateOK = False

            If Side < 0 Then
                If matYears(i) <= GridYear Then candidateOK = True
            Else
                If matYears(i) >= GridYear Then candidateOK = True
            End If

            If candidateOK Then

                If bestIdx = 0 Then

                    bestIdx = i

                ElseIf IsBetterNearestCandidate( _
                    matYears(i), coupon(i), issueNo(i), _
                    matYears(bestIdx), coupon(bestIdx), issueNo(bestIdx), _
                    GridYear) Then

                    bestIdx = i

                End If

            End If

        End If

    Next i

    FindNearestIssue = bestIdx

End Function


Private Function IsBetterNearestCandidate( _
    ByVal xCandidate As Double, _
    ByVal couponCandidate As Double, _
    ByVal issueCandidate As Long, _
    ByVal xBest As Double, _
    ByVal couponBest As Double, _
    ByVal issueBest As Long, _
    ByVal GridYear As Double _
) As Boolean

    Const EPS As Double = 0.000000000001

    Dim dCandidate As Double
    Dim dBest As Double

    dCandidate = Abs(xCandidate - GridYear)
    dBest = Abs(xBest - GridYear)

    If dCandidate < dBest - EPS Then
        IsBetterNearestCandidate = True
        Exit Function
    End If

    ' 同一残存年数の場合のみ、クーポン高・回号大でタイブレーク
    If Abs(xCandidate - xBest) <= EPS Then

        If couponCandidate > couponBest + EPS Then
            IsBetterNearestCandidate = True
            Exit Function
        End If

        If Abs(couponCandidate - couponBest) <= EPS Then
            If issueCandidate > issueBest Then
                IsBetterNearestCandidate = True
                Exit Function
            End If
        End If

    End If

    IsBetterNearestCandidate = False

End Function


' ============================================================
' 10. Node sorting and duplicate compression
' ============================================================

Private Sub SortSplineNodes( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByRef nodeCoupon() As Double, _
    ByRef nodeIssue() As Long, _
    ByVal n As Long _
)

    Dim i As Long
    Dim j As Long

    Dim keyX As Double
    Dim keyY As Double
    Dim keyCoupon As Double
    Dim keyIssue As Long

    For i = 2 To n

        keyX = x(i)
        keyY = y(i)
        keyCoupon = nodeCoupon(i)
        keyIssue = nodeIssue(i)

        j = i - 1

        Do While j >= 1 And x(j) > keyX

            x(j + 1) = x(j)
            y(j + 1) = y(j)
            nodeCoupon(j + 1) = nodeCoupon(j)
            nodeIssue(j + 1) = nodeIssue(j)

            j = j - 1

        Loop

        x(j + 1) = keyX
        y(j + 1) = keyY
        nodeCoupon(j + 1) = keyCoupon
        nodeIssue(j + 1) = keyIssue

    Next i

End Sub


Private Function CompressDuplicateMaturities( _
    ByRef x() As Double, _
    ByRef y() As Double, _
    ByRef nodeCoupon() As Double, _
    ByRef nodeIssue() As Long, _
    ByVal n As Long _
) As Long

    Const EPS As Double = 0.000000000001

    Dim i As Long
    Dim m As Long

    m = 1

    For i = 2 To n

        If Abs(x(i) - x(m)) <= EPS Then

            If nodeCoupon(i) > nodeCoupon(m) + EPS _
                Or (Abs(nodeCoupon(i) - nodeCoupon(m)) <= EPS And nodeIssue(i) > nodeIssue(m)) Then

                x(m) = x(i)
                y(m) = y(i)
                nodeCoupon(m) = nodeCoupon(i)
                nodeIssue(m) = nodeIssue(i)

            End If

        Else

            m = m + 1

            If m <> i Then
                x(m) = x(i)
                y(m) = y(i)
                nodeCoupon(m) = nodeCoupon(i)
                nodeIssue(m) = nodeIssue(i)
            End If

        End If

    Next i

    CompressDuplicateMaturities = m

End Function


' ============================================================
' 11. MOFSplineTable helpers
' ============================================================

Private Function MNB_FindColumnIndex( _
    ByVal rng As Range, _
    ByVal HeaderRowIndex As Long, _
    ParamArray Aliases() As Variant _
) As Long

    Dim c As Long
    Dim a As Long
    Dim headerName As String
    Dim aliasName As String

    For c = 1 To rng.Columns.count

        headerName = MNB_NormalizeHeader(rng.Cells(HeaderRowIndex, c).Value2)

        For a = LBound(Aliases) To UBound(Aliases)

            aliasName = MNB_NormalizeHeader(CStr(Aliases(a)))

            If headerName = aliasName Then
                MNB_FindColumnIndex = c
                Exit Function
            End If

        Next a

    Next c

    MNB_FindColumnIndex = 0

End Function


Private Function MNB_NormalizeHeader(ByVal v As Variant) As String

    Dim s As String

    If IsError(v) Or IsEmpty(v) Then
        MNB_NormalizeHeader = ""
        Exit Function
    End If

    s = LCase$(Trim$(CStr(v)))

    s = Replace(s, " ", "")
    s = Replace(s, "　", "")
    s = Replace(s, "_", "")
    s = Replace(s, "-", "")
    s = Replace(s, "－", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "（", "")
    s = Replace(s, "）", "")
    s = Replace(s, "%", "")

    MNB_NormalizeHeader = s

End Function


Private Function MNB_MofSplineCoreFromBlock( _
    ByVal rng As Range, _
    ByVal HeaderRowIndex As Long, _
    ByVal idxTicker As Long, _
    ByVal idxMaturity As Long, _
    ByVal idxCoupon As Long, _
    ByVal idxYield As Long, _
    ByVal TargetMaturity As Variant, _
    ByVal AsOfValue As Variant, _
    Optional ByVal AllowExtrapolate As Boolean = False, _
    Optional ByVal ShowErrorMessage As Boolean = False _
) As Variant

    On Error GoTo ErrHandler

    Dim AsOfSerial As Double
    Dim ok As Boolean

    If IsEmpty(AsOfValue) Then
        AsOfSerial = CDbl(Date)
    Else
        AsOfSerial = DateValueToSerial(AsOfValue, ok)

        If Not ok Then
            MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
                "AsOfDate が日付として認識できません。", _
                ShowErrorMessage)
            Exit Function
        End If
    End If

    Dim xq As Double
    xq = MaturityToYears(TargetMaturity, AsOfSerial, ok)

    If Not ok Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "TargetMaturity が年限または日付として認識できません。", _
            ShowErrorMessage)
        Exit Function
    End If

    If xq <= 0# Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "TargetMaturity は正の年限、またはAsOfDateより後の日付で指定してください。", _
            ShowErrorMessage)
        Exit Function
    End If

    Dim rawCount As Long
    rawCount = rng.rows.count - HeaderRowIndex

    If rawCount < 1 Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "DataRange にデータ行がありません。", _
            ShowErrorMessage)
        Exit Function
    End If

    Dim tickers() As String
    Dim issueNo() As Long
    Dim originalTerm() As Long
    Dim matYears() As Double
    Dim coupon() As Double
    Dim compYield() As Double
    Dim selected() As Boolean

    ReDim tickers(1 To rawCount)
    ReDim issueNo(1 To rawCount)
    ReDim originalTerm(1 To rawCount)
    ReDim matYears(1 To rawCount)
    ReDim coupon(1 To rawCount)
    ReDim compYield(1 To rawCount)

    Dim i As Long
    Dim rowIndex As Long
    Dim n As Long

    Dim rowTicker As String
    Dim rowTermClass As Long
    Dim rowIssueNo As Long
    Dim rowMatYears As Double
    Dim rowCoupon As Double
    Dim rowYield As Double

    n = 0

    For rowIndex = HeaderRowIndex + 1 To rng.rows.count

        If MNB_TryReadMofInputRow( _
            rng.Cells(rowIndex, idxTicker).Value2, _
            rng.Cells(rowIndex, idxMaturity).Value2, _
            rng.Cells(rowIndex, idxCoupon).Value2, _
            rng.Cells(rowIndex, idxYield).Value2, _
            AsOfSerial, _
            rowTicker, rowTermClass, rowIssueNo, rowMatYears, rowCoupon, rowYield) Then

            n = n + 1

            tickers(n) = rowTicker
            originalTerm(n) = rowTermClass
            issueNo(n) = rowIssueNo
            matYears(n) = rowMatYears
            coupon(n) = rowCoupon
            compYield(n) = rowYield

        End If

    Next rowIndex

    If n < 3 Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "有効なデータ行が3行未満です。空欄・不正値・満期到来済み行を除外した結果、データが不足しています。", _
            ShowErrorMessage)
        Exit Function
    End If

    ReDim Preserve tickers(1 To n)
    ReDim Preserve issueNo(1 To n)
    ReDim Preserve originalTerm(1 To n)
    ReDim Preserve matYears(1 To n)
    ReDim Preserve coupon(1 To n)
    ReDim Preserve compYield(1 To n)
    ReDim selected(1 To n)

    ' -----------------------------
    ' 財務省方式に近い対象銘柄選定
    ' -----------------------------

    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 2)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 5)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 10)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 20)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 30)
    Call SelectCurrentIssue(selected, originalTerm, issueNo, n, 40)

    Dim grid As Long
    Dim termCls As Long
    Dim innerIdx As Long
    Dim outerIdx As Long

    For grid = 1 To 40

        termCls = MofTermClassForGrid(grid)

        innerIdx = FindNearestIssue( _
            originalTerm, matYears, coupon, issueNo, _
            n, termCls, CDbl(grid), -1)

        outerIdx = FindNearestIssue( _
            originalTerm, matYears, coupon, issueNo, _
            n, termCls, CDbl(grid), 1)

        If innerIdx > 0 Then selected(innerIdx) = True
        If outerIdx > 0 Then selected(outerIdx) = True

    Next grid

    ' -----------------------------
    ' 選定銘柄をスプラインノードへ変換
    ' -----------------------------

    Dim m As Long
    m = 0

    For i = 1 To n
        If selected(i) Then m = m + 1
    Next i

    If m < 3 Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "選定後の銘柄数が3未満です。対象データが不足しています。", _
            ShowErrorMessage)
        Exit Function
    End If

    Dim x() As Double
    Dim y() As Double
    Dim nodeCoupon() As Double
    Dim nodeIssue() As Long

    ReDim x(1 To m)
    ReDim y(1 To m)
    ReDim nodeCoupon(1 To m)
    ReDim nodeIssue(1 To m)

    Dim j As Long
    j = 0

    For i = 1 To n

        If selected(i) Then
            j = j + 1

            x(j) = matYears(i)
            y(j) = compYield(i)
            nodeCoupon(j) = coupon(i)
            nodeIssue(j) = issueNo(i)

        End If

    Next i

    Call SortSplineNodes(x, y, nodeCoupon, nodeIssue, m)

    m = CompressDuplicateMaturities(x, y, nodeCoupon, nodeIssue, m)

    If m < 3 Then
        MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
            "重複残存年数を圧縮した結果、スプラインノードが3点未満になりました。", _
            ShowErrorMessage)
        Exit Function
    End If

    MNB_MofSplineCoreFromBlock = NaturalCubicSplineFromArrays( _
        x, y, m, xq, AllowExtrapolate, ShowErrorMessage)

    Exit Function

ErrHandler:
    MNB_MofSplineCoreFromBlock = SplineError(xlErrValue, _
        "MOFSplineTable内部計算中のVBAエラー: " & Err.Description, _
        ShowErrorMessage)

End Function


Private Function MNB_TryReadMofInputRow( _
    ByVal TickerValue As Variant, _
    ByVal MaturityValue As Variant, _
    ByVal CouponValue As Variant, _
    ByVal YieldValue As Variant, _
    ByVal AsOfSerial As Double, _
    ByRef TickerOut As String, _
    ByRef OriginalTermOut As Long, _
    ByRef IssueNoOut As Long, _
    ByRef MatYearsOut As Double, _
    ByRef CouponOut As Double, _
    ByRef CompYieldOut As Double _
) As Boolean

    On Error GoTo BadRow

    Dim ok As Boolean

    ' Ticker
    If IsError(TickerValue) _
        Or IsEmpty(TickerValue) _
        Or Trim$(CStr(TickerValue)) = "" Then

        MNB_TryReadMofInputRow = False
        Exit Function

    End If

    TickerOut = UCase$(Trim$(CStr(TickerValue)))

    OriginalTermOut = MofTermClassFromTicker(TickerOut)

    If OriginalTermOut = 0 Then
        MNB_TryReadMofInputRow = False
        Exit Function
    End If

    IssueNoOut = IssueNumberFromTicker(TickerOut)

    If IssueNoOut < 0 Then
        MNB_TryReadMofInputRow = False
        Exit Function
    End If

    ' Maturity
    MatYearsOut = MaturityToYears(MaturityValue, AsOfSerial, ok)

    If Not ok Then
        MNB_TryReadMofInputRow = False
        Exit Function
    End If

    If MatYearsOut <= 0# Then
        MNB_TryReadMofInputRow = False
        Exit Function
    End If

    ' Coupon
    If IsError(CouponValue) _
        Or IsEmpty(CouponValue) _
        Or Not IsNumeric(CouponValue) Then

        MNB_TryReadMofInputRow = False
        Exit Function

    End If

    ' Compound Yield
    If IsError(YieldValue) _
        Or IsEmpty(YieldValue) _
        Or Not IsNumeric(YieldValue) Then

        MNB_TryReadMofInputRow = False
        Exit Function

    End If

    CouponOut = CDbl(CouponValue)
    CompYieldOut = CDbl(YieldValue)

    MNB_TryReadMofInputRow = True
    Exit Function

BadRow:
    MNB_TryReadMofInputRow = False

End Function

