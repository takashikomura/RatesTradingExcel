Option Explicit

'=========================================================
' ポートフォリオVaR
'
' VaR = Z × √(S'QS)
'
' 使用例
' =VaRCalc(B2:B16, D2:R16)
' =VaRCalc(B2:B16, D2:R16, 2.326)
'=========================================================
Public Function VaRCalc( _
    ByVal RiskVector As Range, _
    ByVal CovMatrix As Range, _
    Optional ByVal Z As Double = 1# _
) As Variant

    Dim S As Variant
    Dim QS As Variant
    Dim Variance As Double
    Dim n As Long

    On Error GoTo ErrorHandler

    S = ToColumnVector(RiskVector)
    n = UBound(S, 1)

    If CovMatrix.rows.count <> n _
       Or CovMatrix.Columns.count <> n Then

        VaRCalc = CVErr(xlErrRef)
        Exit Function
    End If

    If Z < 0# Then
        VaRCalc = CVErr(xlErrNum)
        Exit Function
    End If

    ' Q × S
    QS = Application.MMult(CovMatrix.value2, S)

    ' S' × Q × S
    Variance = CDbl(Application.SumProduct(S, QS))

    If Variance < -0.0000000001 Then
        VaRCalc = CVErr(xlErrNum)
        Exit Function
    End If

    ' 浮動小数点誤差による微小な負値をゼロにする
    Variance = Application.Max(0#, Variance)

    VaRCalc = Z * Sqr(Variance)
    Exit Function

ErrorHandler:
    VaRCalc = CVErr(xlErrValue)

End Function


'=========================================================
' VaR Allocation
'
' Allocation_i
'   = Z × S_i × (QS)_i / √(S'QS)
'
' 使用例
' =VaRAllocate(B2:B16, D2:R16)
' =VaRAllocate(B2:B16, D2:R16, 2.326)
'
' Microsoft 365では各Allocationがスピル表示される
'=========================================================
Public Function VaRAllocate( _
    ByVal RiskVector As Range, _
    ByVal CovMatrix As Range, _
    Optional ByVal Z As Double = 1# _
) As Variant

    Dim S As Variant
    Dim QS As Variant
    Dim Result() As Double

    Dim Variance As Double
    Dim BaseRisk As Double

    Dim n As Long
    Dim i As Long

    On Error GoTo ErrorHandler

    S = ToColumnVector(RiskVector)
    n = UBound(S, 1)

    If CovMatrix.rows.count <> n _
       Or CovMatrix.Columns.count <> n Then

        VaRAllocate = CVErr(xlErrRef)
        Exit Function
    End If

    If Z < 0# Then
        VaRAllocate = CVErr(xlErrNum)
        Exit Function
    End If

    ' Q × S
    QS = Application.MMult(CovMatrix.value2, S)

    ' S' × Q × S
    Variance = CDbl(Application.SumProduct(S, QS))

    If Variance < -0.0000000001 Then
        VaRAllocate = CVErr(xlErrNum)
        Exit Function
    End If

    Variance = Application.Max(0#, Variance)
    BaseRisk = Sqr(Variance)

    ReDim Result(1 To n, 1 To 1)

    If BaseRisk = 0# Then
        VaRAllocate = Result
        Exit Function
    End If

    For i = 1 To n
        Result(i, 1) = _
            Z * CDbl(S(i, 1)) * CDbl(QS(i, 1)) / BaseRisk
    Next i

    ' 元のRiskVectorが横方向なら横方向で返す
    If RiskVector.rows.count = 1 Then
        VaRAllocate = Application.Transpose(Result)
    Else
        VaRAllocate = Result
    End If

    Exit Function

ErrorHandler:
    VaRAllocate = CVErr(xlErrValue)

End Function


'=========================================================
' 縦・横どちらの入力も縦ベクトルへ統一する
'=========================================================
Private Function ToColumnVector( _
    ByVal InputRange As Range _
) As Variant

    If InputRange.Columns.count = 1 Then

        ToColumnVector = InputRange.value2

    ElseIf InputRange.rows.count = 1 Then

        ToColumnVector = Application.Transpose(InputRange.value2)

    Else

        Err.Raise vbObjectError + 1000, _
                  "ToColumnVector", _
                  "RiskVectorは1列または1行で指定してください。"

    End If

End Function

