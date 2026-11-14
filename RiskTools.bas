
Option Explicit

'==========================================================
' JGB Risk Analytics - Unified VBA UDF
'
' INPUT CONVENTIONS
'
' b   : n x 1 or 1 x n Signed BPV
' Cov : n x n covariance of yield changes
' T   : n x n TradeMatrix
'       Rows = Tenors
'       Columns = Base Trades
' V   : n x n orthonormal eigenvector matrix
'       Rows = Tenors
'       Columns = PC1, PC2, ..., PCn
'
' Yield changes must be in bp.
' BPV must be in currency per bp.
'
' FUNCTIONS
'
' 1. TradeVector(T, b)
' 2. TradeCov(T, Cov)
' 3. PortfolioVaR(Cov, b, z)
' 4. VaRAllocate(Cov, b, z)
' 5. TradeVaRAllocate(Cov, b, T, z)
' 6. PCAVaRAllocate(Cov, b, V, T, z)
' 7. PCAVaRDebug(Cov, b, V, T, z)
'
' Default z = 1
'
'==========================================================

Private Const RISK_INPUT_ERROR As Long = vbObjectError + 701
Private Const RISK_MODEL_ERROR As Long = vbObjectError + 702


'==========================================================
' 1. TradeVector
'
' x = T^(-1) * b
'
' Output: n x 1
'==========================================================

Public Function TradeVector( _
    ByVal T As Range, _
    ByVal b As Range _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim tm As Variant
    Dim bv As Variant

    n = T.rows.Count

    If T.Columns.Count <> n Or n < 2 Then
        Err.Raise RISK_INPUT_ERROR, , _
            "TradeMatrix must be square."
    End If

    tm = Risk_ReadMatrix(T)
    bv = Risk_ReadBPV(b, n)

    TradeVector = Risk_SolveTrade(tm, bv)
    Exit Function

Failed:
    TradeVector = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' 2. TradeCov
'
' Cov_T = T' * Cov * T
'
' Output: n x n
'==========================================================

Public Function TradeCov( _
    ByVal T As Range, _
    ByVal Cov As Range _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim tm As Variant
    Dim s As Variant

    Risk_PrepareCov Cov, s, n

    If T.rows.Count <> n Or T.Columns.Count <> n Then
        Err.Raise RISK_INPUT_ERROR, , _
            "TradeMatrix dimension mismatch."
    End If

    tm = Risk_ReadMatrix(T)

    TradeCov = Risk_MM( _
        Risk_MM(Risk_Transpose(tm), s), tm)

    Exit Function

Failed:
    TradeCov = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' 3. PortfolioVaR
'
' VaR = z * sqrt(b' * Cov * b)
'
' Default z = 1
'==========================================================

Public Function PortfolioVaR( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    Optional ByVal z As Double = 1# _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim s As Variant
    Dim bv As Variant

    Risk_CheckZ z
    Risk_PrepareCov Cov, s, n

    bv = Risk_ReadBPV(b, n)

    PortfolioVaR = z * Risk_Volatility(s, bv, n)
    Exit Function

Failed:
    PortfolioVaR = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' 4. VaRAllocate
'
' Tenor-level Euler VaR allocation
'
' A(j) = z * b(j) * (Cov*b)(j) / sigma
'
' Output: n x 1
'==========================================================

Public Function VaRAllocate( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    Optional ByVal z As Double = 1# _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim j As Long

    Dim s As Variant
    Dim bv As Variant
    Dim sb As Variant

    Dim sigma As Double
    Dim output() As Variant

    Risk_CheckZ z
    Risk_PrepareCov Cov, s, n

    bv = Risk_ReadBPV(b, n)
    sb = Risk_MM(s, bv)

    sigma = Risk_Volatility(s, bv, n)

    ReDim output(1 To n, 1 To 1)

    For j = 1 To n

        If sigma > 0# Then

            output(j, 1) = _
                z * bv(j, 1) * sb(j, 1) / sigma

        Else

            output(j, 1) = 0#

        End If

    Next j

    VaRAllocate = output
    Exit Function

Failed:
    VaRAllocate = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' 5. TradeVaRAllocate
'
' Trade-level Euler VaR allocation
'
' x = T^(-1)b
' Cov_T = T'CovT
'
' A(k) = z * x(k)*(Cov_T*x)(k)/sigma
'
' Output: n x 1
'==========================================================

Public Function TradeVaRAllocate( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    ByVal T As Range, _
    Optional ByVal z As Double = 1# _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim k As Long

    Dim s As Variant
    Dim bv As Variant
    Dim tm As Variant

    Dim x As Variant
    Dim covT As Variant
    Dim covTx As Variant

    Dim sigma As Double
    Dim output() As Variant

    Risk_CheckZ z
    Risk_PrepareCov Cov, s, n

    If T.rows.Count <> n Or T.Columns.Count <> n Then
        Err.Raise RISK_INPUT_ERROR, , _
            "TradeMatrix dimension mismatch."
    End If

    bv = Risk_ReadBPV(b, n)
    tm = Risk_ReadMatrix(T)

    x = Risk_SolveTrade(tm, bv)

    covT = Risk_MM( _
        Risk_MM(Risk_Transpose(tm), s), tm)

    covTx = Risk_MM(covT, x)

    sigma = Risk_Volatility(s, bv, n)

    ReDim output(1 To n, 1 To 1)

    For k = 1 To n

        If sigma > 0# Then

            output(k, 1) = _
                z * x(k, 1) * covTx(k, 1) / sigma

        Else

            output(k, 1) = 0#

        End If

    Next k

    TradeVaRAllocate = output
    Exit Function

Failed:
    TradeVaRAllocate = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' 6. PCAVaRAllocate
'
' Factor x Trade Euler VaR allocation
'
' q = V' b
' L = V' T
' Q = V' Cov V
' x = T^(-1)b
'
' A(k,i) =
'   z/sigma * x(k)*L(i,k)*(Q*q)(i)
'
' Output: n x n
'
' Rows    = Base Trades
' Columns = PC1, PC2, ..., PCn
'
' IMPORTANT:
' If V diagonalizes Cov, Q = Lambda.
' If not, Q is the full factor covariance matrix.
'
' Therefore, the same formula remains additive
' even if Cov and V use different estimation windows.
'==========================================================


Public Function PCAVaRAllocate( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    ByVal EigenMatrix As Range, _
    ByVal T As Range, _
    Optional ByVal z As Double = 1# _
) As Variant

    PCAVaRAllocate = Risk_PCAVaRCore( _
        Cov, b, EigenMatrix, T, z, False, False)

End Function




'==========================================================
' 7. PCAVaRDebug
'
' Returns diagnostic information as a single cell.
'
' Example:
' =PCAVaRDebug(Cov, b, EigenMatrix, T)
'
' On failure:
' ERROR [stage] | number | description
'
' On success:
' OK | n=... | sigma=... | ...
'==========================================================

Public Function PCAVaRDebug( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    ByVal EigenMatrix As Range, _
    ByVal T As Range, _
    Optional ByVal z As Double = 1# _
) As Variant

    PCAVaRDebug = Risk_PCAVaRCore( _
        Cov, b, EigenMatrix, T, z, True, True)

End Function

'==========================================================
' PCTradeExposure
'
' Signed PCA Exposure by Base Trade
'
' E(k,i) = x(k) * (V' * T)(i,k)
'
' x = T^(-1) * b
'
' INPUT:
'   EigenMatrix : n x n (Tenor x PC)
'   T           : n x n (Tenor x Base Trade)
'   b           : n x 1 or 1 x n (Signed BPV)
'
' OUTPUT:
'   Rows    = Base Trades
'   Columns = PC1, PC2, ..., PCn
'
' Row sum does NOT represent total BPV.
' Column sum = Portfolio PCA Exposure.
'==========================================================

Public Function PCTradeExposure( _
    ByVal EigenMatrix As Range, _
    ByVal T As Range, _
    ByVal b As Range _
) As Variant

    On Error GoTo Failed

    Dim n As Long
    Dim i As Long, k As Long

    Dim v As Variant
    Dim tm As Variant
    Dim bv As Variant

    Dim x As Variant
    Dim L As Variant

    Dim output() As Variant

    '--------------------------------------------
    ' 1. Validate dimensions
    '--------------------------------------------

    n = T.rows.Count

    If n < 2 Or T.Columns.Count <> n Then
        Err.Raise RISK_INPUT_ERROR
    End If

    If EigenMatrix.rows.Count <> n Or _
       EigenMatrix.Columns.Count <> n Then
        Err.Raise RISK_INPUT_ERROR
    End If

    '--------------------------------------------
    ' 2. Read inputs
    '--------------------------------------------

    tm = Risk_ReadMatrix(T)
    bv = Risk_ReadBPV(b, n)
    v = Risk_ReadMatrix(EigenMatrix)

    '--------------------------------------------
    ' 3. Trade exposure
    '
    ' x = T^(-1) * b
    '--------------------------------------------

    x = Risk_SolveTrade(tm, bv)

    '--------------------------------------------
    ' 4. PCA loading matrix
    '
    ' L = V' * T
    '
    ' L: PC x Trade
    '--------------------------------------------

    L = Risk_MM(Risk_Transpose(v), tm)

    '--------------------------------------------
    ' 5. Signed PCA Exposure
    '
    ' E(k,i) = x(k) * L(i,k)
    '
    ' Output: Trade x PC
    '--------------------------------------------

    ReDim output(1 To n, 1 To n)

    For k = 1 To n

        For i = 1 To n

            output(k, i) = _
                x(k, 1) * L(i, k)

        Next i

    Next k

    PCTradeExposure = output
    Exit Function

Failed:
    PCTradeExposure = Risk_ErrorToExcel(Err.Number)

End Function


'==========================================================
' SHARED CORE FOR PCA VaR ALLOCATION AND DEBUGGING
'==========================================================

Private Function Risk_PCAVaRCore( _
    ByVal Cov As Range, _
    ByVal b As Range, _
    ByVal EigenMatrix As Range, _
    ByVal T As Range, _
    ByVal z As Double, _
    ByVal DebugMode As Boolean, _
    ByVal DiagnosticsOnly As Boolean _
) As Variant

    On Error GoTo Failed

    Dim debugStage As String
    Dim errorNumber As Long
    Dim errorDescription As String

    Dim n As Long
    Dim i As Long, j As Long, k As Long

    Dim s As Variant
    Dim bv As Variant
    Dim tm As Variant

    Dim v As Variant
    Dim vt As Variant

    Dim x As Variant
    Dim q As Variant
    Dim L As Variant

    Dim QCov As Variant
    Dim Qq As Variant
    Dim gram As Variant

    Dim sigma As Double
    Dim orthoError As Double

    Dim maxOffDiag As Double
    Dim covScale As Double

    Dim totalAllocated As Double
    Dim expectedVaR As Double
    Dim difference As Double

    Dim output() As Variant

    '------------------------------------------------------
    ' 1. Validate z
    '------------------------------------------------------

    debugStage = "Risk_CheckZ"
    Risk_CheckZ z

    '------------------------------------------------------
    ' 2. Read and validate covariance
    '------------------------------------------------------

    debugStage = "Risk_PrepareCov"
    Risk_PrepareCov Cov, s, n

    '------------------------------------------------------
    ' 3. Validate dimensions
    '------------------------------------------------------

    debugStage = "Dimension Check"

    If EigenMatrix.rows.Count <> n Or _
       EigenMatrix.Columns.Count <> n Then

        Err.Raise RISK_INPUT_ERROR, , _
            "EigenMatrix dimension mismatch."

    End If

    If T.rows.Count <> n Or T.Columns.Count <> n Then

        Err.Raise RISK_INPUT_ERROR, , _
            "TradeMatrix dimension mismatch."

    End If

    '------------------------------------------------------
    ' 4. Read all input matrices
    '------------------------------------------------------

    debugStage = "Risk_ReadBPV"
    bv = Risk_ReadBPV(b, n)

    debugStage = "Risk_ReadMatrix(T)"
    tm = Risk_ReadMatrix(T)

    debugStage = "Risk_ReadMatrix(EigenMatrix)"
    v = Risk_ReadMatrix(EigenMatrix)

    '------------------------------------------------------
    ' 5. Transpose eigenvector matrix
    '------------------------------------------------------

    debugStage = "Transpose EigenMatrix"
    vt = Risk_Transpose(v)

    '------------------------------------------------------
    ' 6. Check orthonormality: V'V = I
    '
    ' Only orthonormality is required.
    ' V does not need to diagonalize Cov.
    '------------------------------------------------------

    debugStage = "PCA Orthonormality Check"

    gram = Risk_MM(vt, v)

    orthoError = 0#

    For i = 1 To n

        For j = 1 To n

            Dim expectedEntry As Double
            Dim currentError As Double

            If i = j Then
                expectedEntry = 1#
            Else
                expectedEntry = 0#
            End If

            currentError = Abs( _
                CDbl(gram(i, j)) - expectedEntry)

            If currentError > orthoError Then
                orthoError = currentError
            End If

        Next j

    Next i

    If orthoError > 0.0001 Then

        Err.Raise RISK_MODEL_ERROR, , _
            "V is not orthonormal. Max error=" & _
            CStr(orthoError)

    End If

    '------------------------------------------------------
    ' 7. Solve trade exposures
    '
    ' x = T^(-1)b
    '------------------------------------------------------

    debugStage = "TradeVector / MInverse"

    x = Risk_SolveTrade(tm, bv)

    '------------------------------------------------------
    ' 8. PCA exposure
    '
    ' q = V'b
    '------------------------------------------------------

    debugStage = "PCA Exposure: V'b"

    q = Risk_MM(vt, bv)

    '------------------------------------------------------
    ' 9. PCA loading of each trade
    '
    ' L = V'T
    '------------------------------------------------------

    debugStage = "PCA Trade Loading: V'T"

    L = Risk_MM(vt, tm)

    '------------------------------------------------------
    ' 10. Covariance in PCA coordinates
    '
    ' QCov = V'CovV
    '------------------------------------------------------

    debugStage = "Factor Covariance: V'CovV"

    QCov = Risk_MM( _
        Risk_MM(vt, s), v)

    '------------------------------------------------------
    ' 11. Factor marginal covariance
    '
    ' Qq = QCov*q
    '------------------------------------------------------

    debugStage = "Factor Marginal Risk: QCov*q"

    Qq = Risk_MM(QCov, q)

    '------------------------------------------------------
    ' 12. Portfolio volatility
    '------------------------------------------------------

    debugStage = "Portfolio Volatility"

    sigma = Risk_Volatility(s, bv, n)

    '------------------------------------------------------
    ' 13. Factor x Trade VaR allocation
    '------------------------------------------------------

    debugStage = "Factor x Trade Allocation"

    ReDim output(1 To n, 1 To n)

    totalAllocated = 0#

    For k = 1 To n

        For i = 1 To n

            If sigma > 0# Then

                output(k, i) = _
                    (z / sigma) _
                    * x(k, 1) _
                    * L(i, k) _
                    * Qq(i, 1)

            Else

                output(k, i) = 0#

            End If

            totalAllocated = totalAllocated + _
                             CDbl(output(k, i))

        Next i

    Next k

    '------------------------------------------------------
    ' 14. Check total VaR consistency
    '------------------------------------------------------

    debugStage = "VaR Reconciliation"

    expectedVaR = z * sigma

    difference = Abs(totalAllocated - expectedVaR)

    If difference > _
       0.000001 * (1# + Abs(expectedVaR)) Then

        Err.Raise RISK_MODEL_ERROR, , _
            "VaR allocation total differs from " & _
            "portfolio VaR. Difference=" & _
            CStr(difference)

    End If

    '------------------------------------------------------
    ' 15. Diagnostic metrics
    '------------------------------------------------------

    If DiagnosticsOnly Then

        debugStage = "Diagnostic Metrics"

        covScale = 0#
        maxOffDiag = 0#

        For i = 1 To n

            If Abs(QCov(i, i)) > covScale Then
                covScale = Abs(QCov(i, i))
            End If

            For j = 1 To n

                If i <> j Then

                    If Abs(QCov(i, j)) > maxOffDiag Then
                        maxOffDiag = Abs(QCov(i, j))
                    End If

                End If

            Next j

        Next i

        Risk_PCAVaRCore = _
            "OK" & _
            " | n=" & CStr(n) & _
            " | sigma=" & Format(sigma, "0.000000E+00") & _
            " | MaxOrthoError=" & _
                Format(orthoError, "0.000000E+00") & _
            " | MaxFactorCovOffDiag=" & _
                Format(maxOffDiag, "0.000000E+00") & _
            " | CovScale=" & _
                Format(covScale, "0.000000E+00") & _
            " | AllocationDiff=" & _
                Format(difference, "0.000000E+00")

    Else

        Risk_PCAVaRCore = output

    End If

    Exit Function

Failed:

    errorNumber = Err.Number
    errorDescription = Err.Description

    If DebugMode Then

        Risk_PCAVaRCore = _
            "ERROR [" & debugStage & "]" & _
            " | " & CStr(errorNumber) & _
            " | " & errorDescription

    Else

        Risk_PCAVaRCore = _
            Risk_ErrorToExcel(errorNumber)

    End If

End Function


'==========================================================
' PRIVATE HELPER FUNCTIONS
'==========================================================


'==========================================================
' Risk_MM
'
' Matrix multiplication with Excel MMULT
'==========================================================

Private Function Risk_MM( _
    ByVal a As Variant, _
    ByVal b As Variant _
) As Variant

    Risk_MM = Application.WorksheetFunction.MMult(a, b)

End Function


'==========================================================
' Risk_Transpose
'
' Explicit transpose of 2D arrays.
'
' Avoids unexpected one-dimensional array returns
' from WorksheetFunction.Transpose.
'==========================================================

Private Function Risk_Transpose( _
    ByVal a As Variant _
) As Variant

    Dim nr As Long
    Dim nc As Long

    Dim i As Long
    Dim j As Long

    Dim output() As Double

    nr = UBound(a, 1)
    nc = UBound(a, 2)

    ReDim output(1 To nc, 1 To nr)

    For i = 1 To nr

        For j = 1 To nc

            output(j, i) = CDbl(a(i, j))

        Next j

    Next i

    Risk_Transpose = output

End Function


'==========================================================
' Risk_SolveTrade
'
' Solve T*x = b using Excel MINVERSE and MMULT.
'
' n = 15 is well within Excel's matrix-function limits.
'==========================================================

Private Function Risk_SolveTrade( _
    ByVal T As Variant, _
    ByVal b As Variant _
) As Variant

    Dim invT As Variant

    invT = Application.WorksheetFunction.MInverse(T)

    Risk_SolveTrade = Risk_MM(invT, b)

End Function


'==========================================================
' Risk_ReadMatrix
'
' Read numeric Range into 1-based 2D Double array.
'==========================================================

Private Function Risk_ReadMatrix( _
    ByVal rg As Range _
) As Variant

    Dim nr As Long
    Dim nc As Long

    Dim i As Long
    Dim j As Long

    Dim raw As Variant
    Dim cellValue As Variant

    Dim output() As Double

    If rg.Areas.Count <> 1 Then

        Err.Raise RISK_INPUT_ERROR, , _
            "Multi-area ranges are not supported."

    End If

    nr = rg.rows.Count
    nc = rg.Columns.Count

    ReDim output(1 To nr, 1 To nc)

    raw = rg.value2

    For i = 1 To nr

        For j = 1 To nc

            If nr = 1 And nc = 1 Then
                cellValue = raw
            Else
                cellValue = raw(i, j)
            End If

            If IsError(cellValue) Then

                Err.Raise RISK_INPUT_ERROR, , _
                    "Excel error in input at (" & _
                    CStr(i) & "," & CStr(j) & ")."

            End If

            If IsEmpty(cellValue) Then

                Err.Raise RISK_INPUT_ERROR, , _
                    "Empty input at (" & _
                    CStr(i) & "," & CStr(j) & ")."

            End If

            If Not IsNumeric(cellValue) Then

                Err.Raise RISK_INPUT_ERROR, , _
                    "Nonnumeric input at (" & _
                    CStr(i) & "," & CStr(j) & ")."

            End If

            output(i, j) = CDbl(cellValue)

        Next j

    Next i

    Risk_ReadMatrix = output

End Function


'==========================================================
' Risk_ReadBPV
'
' Accept n x 1 or 1 x n.
' Return n x 1.
'==========================================================

Private Function Risk_ReadBPV( _
    ByVal rg As Range, _
    ByVal n As Long _
) As Variant

    Dim raw As Variant
    Dim output() As Double

    Dim i As Long

    If Not ( _
        (rg.rows.Count = n And rg.Columns.Count = 1) Or _
        (rg.rows.Count = 1 And rg.Columns.Count = n) _
    ) Then

        Err.Raise RISK_INPUT_ERROR, , _
            "BPV must be n x 1 or 1 x n."

    End If

    raw = Risk_ReadMatrix(rg)

    ReDim output(1 To n, 1 To 1)

    For i = 1 To n

        If rg.Columns.Count = 1 Then

            output(i, 1) = raw(i, 1)

        Else

            output(i, 1) = raw(1, i)

        End If

    Next i

    Risk_ReadBPV = output

End Function


'==========================================================
' Risk_PrepareCov
'
' Validate covariance:
'   - square
'   - numeric
'   - nonnegative diagonal
'   - symmetric within tolerance
'
' Small asymmetry is removed by averaging.
'==========================================================

Private Sub Risk_PrepareCov( _
    ByVal rg As Range, _
    ByRef s As Variant, _
    ByRef n As Long _
)

    Dim i As Long
    Dim j As Long

    Dim tolerance As Double
    Dim avg As Double

    n = rg.rows.Count

    If n < 2 Or rg.Columns.Count <> n Then

        Err.Raise RISK_INPUT_ERROR, , _
            "Covariance matrix must be square."

    End If

    s = Risk_ReadMatrix(rg)

    For i = 1 To n

        If s(i, i) < 0# Then

            Err.Raise RISK_MODEL_ERROR, , _
                "Negative covariance diagonal at " & _
                CStr(i)

        End If

        For j = i + 1 To n

            tolerance = 0.00000001 * _
                (1# + Abs(s(i, j)) + Abs(s(j, i)))

            If Abs(s(i, j) - s(j, i)) > tolerance Then

                Err.Raise RISK_MODEL_ERROR, , _
                    "Covariance is asymmetric at (" & _
                    CStr(i) & "," & CStr(j) & ")."

            End If

            avg = 0.5 * (s(i, j) + s(j, i))

            s(i, j) = avg
            s(j, i) = avg

        Next j

    Next i

End Sub


'==========================================================
' Risk_Volatility
'
' sigma = sqrt(b' Cov b)
'==========================================================

Private Function Risk_Volatility( _
    ByVal s As Variant, _
    ByVal b As Variant, _
    ByVal n As Long _
) As Double

    Dim sb As Variant

    Dim i As Long
    Dim variance As Double
    Dim absSum As Double
    Dim term As Double

    sb = Risk_MM(s, b)

    variance = 0#
    absSum = 0#

    For i = 1 To n

        term = b(i, 1) * sb(i, 1)

        variance = variance + term
        absSum = absSum + Abs(term)

    Next i

    If variance < -0.0000000001 * absSum Then

        Err.Raise RISK_MODEL_ERROR, , _
            "Negative portfolio variance."

    End If

    If variance < 0# Then variance = 0#

    Risk_Volatility = Sqr(variance)

End Function


'==========================================================
' Risk_CheckZ
'==========================================================

Private Sub Risk_CheckZ(ByVal z As Double)

    If z < 0# Then

        Err.Raise RISK_INPUT_ERROR, , _
            "z must be nonnegative."

    End If

End Sub


'==========================================================
' Risk_ErrorToExcel
'
' Input errors -> #VALUE!
' Calculation/model errors -> #NUM!
'==========================================================

Private Function Risk_ErrorToExcel( _
    ByVal errorNumber As Long _
) As Variant

    If errorNumber = RISK_INPUT_ERROR Then

        Risk_ErrorToExcel = CVErr(xlErrValue)

    Else

        Risk_ErrorToExcel = CVErr(xlErrNum)

    End If

End Function


