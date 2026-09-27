
Option Explicit

'========================================================
' Covariance Eigenvalue Decomposition UDF
'
' 使用例:
'
' =CovEigen(CovMatrix30,"Lambda")
'
' =CovEigen(CovMatrix30,"Vectors")
'
' =CovEigen(CovMatrix30)
'
' 入力:
'   対称な共分散行列
'
' 出力:
'   Lambda  固有値対角行列
'   Vectors 固有ベクトル行列
'   Both    LambdaとVectorsを横方向に結合
'
' 欠損系列:
'   対角要素が #N/A の系列を除外する。
'
'========================================================

Private Const EIGEN_TOL As Double = 0.000000000001

Private Const EIGEN_MAX_SWEEPS As Long = 100

Private Const EIGEN_MAX_DIMENSION As Long = 300


'========================================================
' Excel UDF
'
' covInput:
'   共分散行列
'
' outputMode:
'   "Lambda"
'   "Vectors"
'   "Both"
'
'========================================================

Public Function CovEigen( _
    ByVal covInput As Variant, _
    Optional ByVal outputMode As String = "Both" _
) As Variant

    Dim rawData As Variant

    Dim A() As Double
    Dim eigenValues() As Double
    Dim EigenVectors() As Double

    Dim validIndices() As Long

    Dim Result() As Variant

    Dim fullN As Long
    Dim n As Long

    Dim firstRow As Long
    Dim firstCol As Long

    Dim i As Long
    Dim j As Long

    Dim sourceI As Long
    Dim sourceJ As Long

    Dim mode As String

    Dim outputCols As Long

    Dim v As Variant

    Dim valueIJ As Double
    Dim valueJI As Double

    Dim matrixScale As Double
    Dim matrixTolerance As Double

    On Error GoTo ErrHandler


    '====================================================
    ' 1. 出力モード
    '====================================================

    mode = UCase$(Trim$(outputMode))


    Select Case mode

        Case "LAMBDA"
            ' 固有値対角行列

        Case "VECTORS"
            ' 固有ベクトル行列

        Case "BOTH", ""
            mode = "BOTH"

        Case Else

            CovEigen = CVErr(xlErrValue)

            Exit Function

    End Select


    '====================================================
    ' 2. 入力データ取得
    '
    ' Rangeを受け取った場合はValue2を使用する。
    '
    ' 配列を受け取った場合はそのまま使用する。
    '====================================================

    If IsObject(covInput) Then

        If TypeName(covInput) <> "Range" Then

            CovEigen = CVErr(xlErrValue)

            Exit Function

        End If

        rawData = covInput.value2

    Else

        rawData = covInput

    End If


    '====================================================
    ' 3. 入力を2次元配列に統一
    '====================================================

    If Not IsArray(rawData) Then

        v = rawData

        ReDim rawData(1 To 1, 1 To 1)

        rawData(1, 1) = v

    End If


    firstRow = LBound(rawData, 1)
    firstCol = LBound(rawData, 2)


    fullN = _
        UBound(rawData, 1) - firstRow + 1


    If fullN < 1 Then

        CovEigen = CVErr(xlErrValue)

        Exit Function

    End If


    '====================================================
    ' 正方行列の検証
    '====================================================

    If UBound(rawData, 2) - firstCol + 1 _
       <> fullN Then

        CovEigen = CVErr(xlErrValue)

        Exit Function

    End If


    '====================================================
    ' 4. 欠損系列の除外
    '
    ' 対角要素が数値なら有効候補。
    '
    ' 対角要素が #N/A なら系列を除外する。
    '
    ' その他のエラーや文字列は不正入力。
    '====================================================

    ReDim validIndices(1 To fullN)

    n = 0


    For i = 1 To fullN

        v = rawData( _
            firstRow + i - 1, _
            firstCol + i - 1)


        If IsError(v) Then

            If v = CVErr(xlErrNA) Then

                ' #N/A の系列を除外

            Else

                CovEigen = CVErr(xlErrValue)

                Exit Function

            End If


        ElseIf IsFiniteNumeric_Local(v) Then

            n = n + 1

            validIndices(n) = i


        Else

            CovEigen = CVErr(xlErrValue)

            Exit Function

        End If

    Next i


    '====================================================
    ' 有効系列が存在しない場合
    '====================================================

    If n = 0 Then

        CovEigen = CVErr(xlErrNA)

        Exit Function

    End If


    '====================================================
    ' Jacobi法の計算量を考慮して上限を設定
    '====================================================

    If n > EIGEN_MAX_DIMENSION Then

        CovEigen = CVErr(xlErrNum)

        Exit Function

    End If


    '====================================================
    ' 5. 有効系列の部分行列を作成
    '====================================================

    ReDim A(1 To n, 1 To n)

    matrixScale = 0#


    For i = 1 To n

        sourceI = validIndices(i)


        For j = 1 To n

            sourceJ = validIndices(j)


            v = rawData( _
                firstRow + sourceI - 1, _
                firstCol + sourceJ - 1)


            '----------------------------------------
            ' 有効系列間に欠損値が残っている場合
            '----------------------------------------

            If Not IsFiniteNumeric_Local(v) Then

                CovEigen = CVErr(xlErrValue)

                Exit Function

            End If


            A(i, j) = CDbl(v)


            If Abs(A(i, j)) > matrixScale Then

                matrixScale = Abs(A(i, j))

            End If

        Next j

    Next i


    '====================================================
    ' 6. 対称性の検証
    '
    ' 共分散行列は対称行列である必要がある。
    '====================================================

    matrixTolerance = EIGEN_TOL * matrixScale


    For i = 1 To n

        For j = i + 1 To n

            valueIJ = A(i, j)
            valueJI = A(j, i)


            If Abs(valueIJ - valueJI) > _
               matrixTolerance Then

                CovEigen = CVErr(xlErrNum)

                Exit Function

            End If


            '----------------------------------------
            ' 許容誤差以内の非対称成分を平均化
            '----------------------------------------

            A(i, j) = _
                (valueIJ + valueJI) / 2#

            A(j, i) = A(i, j)

        Next j

    Next i


    '====================================================
    ' 7. 固有値分解
    '====================================================

    If Not JacobiEigen_Local( _
        A, _
        n, _
        eigenValues, _
        EigenVectors, _
        matrixTolerance) Then

        CovEigen = CVErr(xlErrNum)

        Exit Function

    End If


    '====================================================
    ' 8. 固有値を降順にソート
    '====================================================

    SortEigenpairs_Local _
        eigenValues, _
        EigenVectors, _
        n


    '====================================================
    ' 9. 共分散行列の半正定値性を検証
    '
    ' 有意な負の固有値がある場合は #NUM!
    '
    ' 丸め誤差程度の負の固有値は0に補正する。
    '====================================================

    For i = 1 To n

        If eigenValues(i) < -matrixTolerance Then

            CovEigen = CVErr(xlErrNum)

            Exit Function

        End If


        If eigenValues(i) < 0# Then

            eigenValues(i) = 0#

        End If

    Next i


    '====================================================
    ' 10. 出力配列
    '====================================================

    Select Case mode

        Case "LAMBDA"

            outputCols = n


        Case "VECTORS"

            outputCols = n


        Case "BOTH"

            ' Lambda + 空白列 + Vectors

            outputCols = 2 * n + 1

    End Select


    ReDim Result( _
        1 To n, _
        1 To outputCols)


    '====================================================
    ' 11. Lambda出力
    '====================================================

    If mode = "LAMBDA" Or mode = "BOTH" Then

        For i = 1 To n

            For j = 1 To n

                If i = j Then

                    Result(i, j) = eigenValues(i)

                Else

                    Result(i, j) = 0#

                End If

            Next j

        Next i

    End If


    '====================================================
    ' 12. 固有ベクトル出力
    '
    ' V(row, component)
    '
    ' 各列が1つの主成分を表す。
    '====================================================

    If mode = "VECTORS" Then

        For i = 1 To n

            For j = 1 To n

                Result(i, j) = _
                    EigenVectors(i, j)

            Next j

        Next i


    ElseIf mode = "BOTH" Then

        For i = 1 To n

            For j = 1 To n

                Result(i, n + 1 + j) = _
                    EigenVectors(i, j)

            Next j

        Next i

    End If


    CovEigen = Result

    Exit Function


ErrHandler:

    CovEigen = CVErr(xlErrValue)

End Function


'========================================================
' 数値の有効性
'
' 数値以外の型を除外する。
'
' 空白、文字列、Boolean、Excelエラー、
' Null等は無効。
'
'========================================================

Private Function IsFiniteNumeric_Local( _
    ByVal v As Variant _
) As Boolean

    IsFiniteNumeric_Local = False


    If IsError(v) Then Exit Function

    If IsEmpty(v) Then Exit Function

    If IsNull(v) Then Exit Function


    Select Case VarType(v)

        Case vbByte, vbInteger, vbLong, _
             vbSingle, vbDouble, vbCurrency, vbDecimal

            If Not IsNumeric(v) Then
                Exit Function
            End If


            IsFiniteNumeric_Local = True

    End Select

End Function


'========================================================
' Jacobi固有値分解
'
' A: 対称行列
'
' EigenValues:
'   固有値
'
' EigenVectors:
'   固有ベクトルを列方向に格納
'
' A * V = V * Lambda
'
'========================================================

Private Function JacobiEigen_Local( _
    ByRef inputA() As Double, _
    ByVal n As Long, _
    ByRef eigenValues() As Double, _
    ByRef EigenVectors() As Double, _
    ByVal tolerance As Double _
) As Boolean

    Dim A() As Double

    Dim i As Long
    Dim j As Long

    Dim p As Long
    Dim q As Long
    Dim r As Long

    Dim sweep As Long

    Dim app As Double
    Dim aqq As Double
    Dim apq As Double

    Dim tau As Double
    Dim t As Double

    Dim c As Double
    Dim s As Double

    Dim arp As Double
    Dim arq As Double

    Dim vrp As Double
    Dim vrq As Double

    Dim maxOffDiag As Double

    JacobiEigen_Local = False


    '====================================================
    ' 入力行列をコピー
    '====================================================

    ReDim A(1 To n, 1 To n)

    ReDim eigenValues(1 To n)

    ReDim EigenVectors(1 To n, 1 To n)


    For i = 1 To n

        EigenVectors(i, i) = 1#


        For j = 1 To n

            A(i, j) = inputA(i, j)

        Next j

    Next i


    '====================================================
    ' 1×1行列
    '====================================================

    If n = 1 Then

        eigenValues(1) = A(1, 1)

        JacobiEigen_Local = True

        Exit Function

    End If


    '====================================================
    ' Jacobi sweeps
    '====================================================

    For sweep = 1 To EIGEN_MAX_SWEEPS

        maxOffDiag = 0#


        '--------------------------------------------
        ' 最大非対角成分を取得
        '--------------------------------------------

        For p = 1 To n - 1

            For q = p + 1 To n

                If Abs(A(p, q)) > maxOffDiag Then

                    maxOffDiag = Abs(A(p, q))

                End If

            Next q

        Next p


        '--------------------------------------------
        ' 収束判定
        '--------------------------------------------

        If maxOffDiag <= tolerance Then

            For i = 1 To n

                eigenValues(i) = A(i, i)

            Next i


            JacobiEigen_Local = True

            Exit Function

        End If


        '================================================
        ' 非対角要素をJacobi回転で消去
        '================================================

        For p = 1 To n - 1

            For q = p + 1 To n

                apq = A(p, q)


                If Abs(apq) > tolerance Then

                    app = A(p, p)
                    aqq = A(q, q)


                    '------------------------------------
                    ' 回転パラメータ
                    '------------------------------------

                    tau = _
                        (aqq - app) / (2# * apq)


                    If tau >= 0# Then

                        t = 1# / _
                            (tau + Sqr(1# + tau * tau))

                    Else

                        t = -1# / _
                            (-tau + Sqr(1# + tau * tau))

                    End If


                    c = 1# / Sqr(1# + t * t)

                    s = t * c


                    '------------------------------------
                    ' 対角要素更新
                    '------------------------------------

                    A(p, p) = app - t * apq

                    A(q, q) = aqq + t * apq


                    A(p, q) = 0#
                    A(q, p) = 0#


                    '------------------------------------
                    ' 他の行・列の更新
                    '------------------------------------

                    For r = 1 To n

                        If r <> p And r <> q Then

                            arp = A(r, p)
                            arq = A(r, q)


                            A(r, p) = _
                                c * arp - s * arq

                            A(p, r) = A(r, p)


                            A(r, q) = _
                                s * arp + c * arq

                            A(q, r) = A(r, q)

                        End If

                    Next r


                    '------------------------------------
                    ' 固有ベクトル更新
                    '------------------------------------

                    For r = 1 To n

                        vrp = EigenVectors(r, p)
                        vrq = EigenVectors(r, q)


                        EigenVectors(r, p) = _
                            c * vrp - s * vrq

                        EigenVectors(r, q) = _
                            s * vrp + c * vrq

                    Next r

                End If

            Next q

        Next p

    Next sweep


    '====================================================
    ' 最大Sweep数到達後の収束判定
    '====================================================

    maxOffDiag = 0#


    For p = 1 To n - 1

        For q = p + 1 To n

            If Abs(A(p, q)) > maxOffDiag Then

                maxOffDiag = Abs(A(p, q))

            End If

        Next q

    Next p


    If maxOffDiag <= tolerance Then

        For i = 1 To n

            eigenValues(i) = A(i, i)

        Next i


        JacobiEigen_Local = True

    End If

End Function


'========================================================
' 固有値・固有ベクトルの並べ替え
'
' 固有値を降順に並べる。
'
' 固有ベクトルは対応関係を維持して列を交換。
'
' 符号は絶対値最大要素が正になるよう統一する。
'
'========================================================

Private Sub SortEigenpairs_Local( _
    ByRef eigenValues() As Double, _
    ByRef EigenVectors() As Double, _
    ByVal n As Long _
)

    Dim i As Long
    Dim j As Long
    Dim k As Long

    Dim maxIndex As Long

    Dim tmpValue As Double
    Dim tmpVector As Double

    Dim pivotRow As Long
    Dim maxAbs As Double


    '====================================================
    ' 固有値降順
    '====================================================

    For i = 1 To n - 1

        maxIndex = i


        For j = i + 1 To n

            If eigenValues(j) > eigenValues(maxIndex) Then

                maxIndex = j

            End If

        Next j


        If maxIndex <> i Then

            tmpValue = eigenValues(i)

            eigenValues(i) = eigenValues(maxIndex)

            eigenValues(maxIndex) = tmpValue


            '----------------------------------------
            ' 固有ベクトルの列交換
            '----------------------------------------

            For k = 1 To n

                tmpVector = EigenVectors(k, i)

                EigenVectors(k, i) = _
                    EigenVectors(k, maxIndex)

                EigenVectors(k, maxIndex) = _
                    tmpVector

            Next k

        End If

    Next i


    '====================================================
    ' 固有ベクトルの符号統一
    '
    ' 絶対値最大の要素が正になるようにする。
    '====================================================

    For j = 1 To n

        maxAbs = 0#
        pivotRow = 1


        For i = 1 To n

            If Abs(EigenVectors(i, j)) > maxAbs Then

                maxAbs = Abs(EigenVectors(i, j))

                pivotRow = i

            End If

        Next i


        If EigenVectors(pivotRow, j) < 0# Then

            For i = 1 To n

                EigenVectors(i, j) = _
                    -EigenVectors(i, j)

            Next i

        End If

    Next j

End Sub



' =====================================================================
' PCAHedge
'
' 指定したspread structureについて、
' 指定したPrincipal Componentsへのexposureが0になる
' hedge ratio / leg weightを計算する。
'
'
' =====================================================================
' USAGE
' =====================================================================
'
' 3-leg fly:
'
'   =PCAHedge("5s7s10s", EigenVectors, "1,2")
'
'
' 2-leg curve:
'
'   =PCAHedge("5s10s", EigenVectors, "1")
'
'
' 4-leg structure:
'
'   =PCAHedge("2s5s10s20s", EigenVectors, "1,2,5")
'
'
' =====================================================================
' EigenVectors FORMAT
' =====================================================================
'
'                   PC1       PC2       PC3       PC4
'   yield_2y        ...
'   yield_3y        ...
'   yield_5y        ...
'   7y              ...
'   yield_10y       ...
'   yield_20y       ...
'
'
' ・1行目 = PC headers
' ・1列目 = tenor labels
' ・数値部分 = eigenvector loadings
'
'
' =====================================================================
' MATHEMATICS
' =====================================================================
'
' Selected-leg weight vector:
'
'       w
'
' Hedge対象PCのloading matrix:
'
'       V_H
'
' PCA neutrality:
'
'       V_H' w = 0
'
'
' m legsについてm-1個のPCをneutralizeすると、
' overall scaleを除いてweight ratioが一意に決まる。
'
'
' Normalization:
'
'   odd number of legs:
'
'       middle leg = +1
'
'   even number of legs:
'
'       first leg = +1
'
'
' =====================================================================
' OUTPUT
' =====================================================================
'
' Output size:
'
'       EigenVectors.Rows.Count - 1
'
' つまりEigenVectorsのtenor数と同じ長さの
' column vectorとして縦方向にspillする。
'
'
' Example:
'
' EigenVectors rows:
'
'   yield_2y
'   yield_3y
'   yield_5y
'   yield_7y
'   yield_10y
'   yield_20y
'
'
' =PCAHedge("5s7s10s",EigenVectors,"1,2")
'
' might return:
'
'       0
'       0
'      -0.72
'       1.00
'      -0.39
'       0
'
'
' =====================================================================

Public Function PCAHedge( _
    ByVal SpreadName As String, _
    ByVal EigenVectors As Range, _
    ByVal HedgePCs As String _
) As Variant

    On Error GoTo ErrorHandler


    ' =================================================================
    ' 1. Parse spread tenors
    ' =================================================================

    Dim Legs As Variant
    Legs = ParseSpreadTenors(SpreadName)


    Dim m As Long

    m = UBound(Legs) - LBound(Legs) + 1


    If m < 2 Then

        PCAHedge = CVErr(xlErrValue)
        Exit Function

    End If


    ' =================================================================
    ' 2. Parse hedge PC list
    '
    ' Example:
    '
    '   "1,2,5"
    '
    ' becomes
    '
    '   {1,2,5}
    '
    ' =================================================================

    Dim PCs As Variant
    PCs = ParsePCList(HedgePCs)


    Dim q As Long

    q = UBound(PCs) - LBound(PCs) + 1


    ' =================================================================
    ' 3. Number of constraints
    '
    ' m legs require m-1 hedge PCs
    '
    ' Examples:
    '
    '   2 legs -> 1 PC
    '   3 legs -> 2 PCs
    '   4 legs -> 3 PCs
    '
    ' =================================================================

    If q <> m - 1 Then

        PCAHedge = CVErr(xlErrValue)
        Exit Function

    End If


    ' =================================================================
    ' 4. Validate EigenVectors
    ' =================================================================

    If EigenVectors Is Nothing Then

        PCAHedge = CVErr(xlErrRef)
        Exit Function

    End If


    ' Need:
    '
    ' header row
    ' + at least one tenor row

    If EigenVectors.rows.Count < 2 Then

        PCAHedge = CVErr(xlErrRef)
        Exit Function

    End If


    ' Need:
    '
    ' tenor-name column
    ' + at least one PC column

    If EigenVectors.Columns.Count < 2 Then

        PCAHedge = CVErr(xlErrRef)
        Exit Function

    End If


    ' =================================================================
    ' 5. Reject duplicated PC numbers
    '
    ' Example:
    '
    '   "1,1"
    '
    ' is invalid.
    '
    ' =================================================================

    If HasDuplicateLongValues(PCs) Then

        PCAHedge = CVErr(xlErrValue)
        Exit Function

    End If


    ' =================================================================
    ' 6. Find selected tenor rows inside EigenVectors
    ' =================================================================

    Dim TenorRows() As Long

    ReDim TenorRows(1 To m)


    Dim j As Long


    For j = 1 To m

        TenorRows(j) = FindTenorRowInMatrix( _
                            EigenVectors, _
                            CDbl(Legs(j)) _
                       )


        Select Case TenorRows(j)


            Case 0

                ' Tenor not found

                PCAHedge = CVErr(xlErrNA)
                Exit Function


            Case -1

                ' Duplicate / ambiguous tenor

                PCAHedge = CVErr(xlErrRef)
                Exit Function


        End Select

    Next j


    ' =================================================================
    ' 7. Find selected PC columns inside EigenVectors
    ' =================================================================

    Dim PCCols() As Long

    ReDim PCCols(1 To q)


    Dim i As Long


    For i = 1 To q

        PCCols(i) = FindPCColumn( _
                        EigenVectors, _
                        CLng(PCs(i)) _
                    )


        Select Case PCCols(i)


            Case 0

                ' PC header not found

                PCAHedge = CVErr(xlErrNA)
                Exit Function


            Case -1

                ' Duplicate PC header

                PCAHedge = CVErr(xlErrRef)
                Exit Function


        End Select

    Next i


    ' =================================================================
    ' 8. Construct linear system
    '
    '
    '      [ V_H'  ]        [ 0 ]
    '      [       ] w   =  [ . ]
    '      [anchor ]        [ 1 ]
    '
    '
    ' A w = rhs
    '
    '
    ' First q rows:
    '
    '       selected-PC neutrality
    '
    ' Last row:
    '
    '       normalization
    '
    ' =================================================================

    Dim A() As Variant
    Dim rhs() As Variant


    ReDim A(1 To m, 1 To m)
    ReDim rhs(1 To m, 1 To 1)


    ' -----------------------------------------------------------------
    ' PCA-neutrality constraints
    ' -----------------------------------------------------------------

    Dim LoadingValue As Variant


    For i = 1 To q

        For j = 1 To m

            LoadingValue = EigenVectors.Cells( _
                                TenorRows(j), _
                                PCCols(i) _
                           ).value2


            If IsError(LoadingValue) Then

                PCAHedge = CVErr(xlErrValue)
                Exit Function

            End If


            If Not IsNumeric(LoadingValue) Then

                PCAHedge = CVErr(xlErrValue)
                Exit Function

            End If


            A(i, j) = CDbl(LoadingValue)

        Next j


        rhs(i, 1) = 0#

    Next i


    ' =================================================================
    ' 9. Normalization
    '
    '
    ' Odd number of legs:
    '
    '       middle leg = +1
    '
    '
    ' Example:
    '
    '       5s7s10s
    '
    '       w_7 = 1
    '
    '
    ' Even number of legs:
    '
    '       first leg = +1
    '
    ' =================================================================

    Dim AnchorLeg As Long


    If m Mod 2 = 1 Then

        AnchorLeg = (m + 1) \ 2

    Else

        AnchorLeg = 1

    End If


    For j = 1 To m

        A(m, j) = 0#

    Next j


    A(m, AnchorLeg) = 1#

    rhs(m, 1) = 1#


    ' =================================================================
    ' 10. Solve linear system
    '
    '
    '       A w = rhs
    '
    '
    '       w = A^(-1) rhs
    '
    ' =================================================================

    Dim InvA As Variant
    Dim Weights As Variant


    InvA = Application.MInverse(A)


    If IsError(InvA) Then

        PCAHedge = CVErr(xlErrNum)
        Exit Function

    End If


    Weights = Application.MMult(InvA, rhs)


    If IsError(Weights) Then

        PCAHedge = CVErr(xlErrNum)
        Exit Function

    End If


    ' =================================================================
    ' 11. Construct full-size output vector
    '
    '
    ' Number of output rows:
    '
    '       number of tenor rows in EigenVectors
    '
    '
    ' Since EigenVectors row 1 is header:
    '
    '       n = Rows.Count - 1
    '
    '
    ' Non-traded tenors:
    '
    '       0
    '
    '
    ' Traded tenors:
    '
    '       calculated PCA hedge weights
    '
    ' =================================================================

    Dim n As Long

    n = EigenVectors.rows.Count - 1


    Dim Result() As Variant

    ReDim Result(1 To n, 1 To 1)


    Dim r As Long


    ' -----------------------------------------------------------------
    ' Initialise all rows to zero
    ' -----------------------------------------------------------------

    For r = 1 To n

        Result(r, 1) = 0#

    Next r


    ' -----------------------------------------------------------------
    ' Insert calculated hedge ratios into corresponding tenor rows
    '
    '
    ' TenorRows(j):
    '
    '       relative row inside EigenVectors
    '
    '
    ' EigenVectors:
    '
    '       row 1 = header
    '
    '
    ' Therefore:
    '
    '       Result row = TenorRows(j) - 1
    '
    ' -----------------------------------------------------------------

    Dim w As Double


    For j = 1 To m

        w = CDbl(Weights(j, 1))


        ' Remove numerical floating-point noise

        If Abs(w) < 0.000000000001 Then

            w = 0#

        End If


        Result(TenorRows(j) - 1, 1) = w

    Next j


    ' =================================================================
    ' 12. Return spill array
    ' =================================================================

    PCAHedge = Result

    Exit Function



ErrorHandler:

    PCAHedge = CVErr(xlErrValue)

End Function



' =====================================================================
' ParseSpreadTenors
'
' Extract tenors from spread name.
'
'
' Examples
' ---------------------------------------------------------------------
'
'   "5s10s"
'
'       -> {5,10}
'
'
'   "5s7s10s"
'
'       -> {5,7,10}
'
'
'   "2s5s10s20s"
'
'       -> {2,5,10,20}
'
'
'   "5Y-7Y-10Y"
'
'       -> {5,7,10}
'
'
'   "5/7/10"
'
'       -> {5,7,10}
'
'
'   "2.5s5s10s"
'
'       -> {2.5,5,10}
'
' =====================================================================

Private Function ParseSpreadTenors( _
    ByVal SpreadName As String _
) As Variant


    Dim txt As String

    txt = Trim$(SpreadName)


    If Len(txt) = 0 Then

        Err.Raise _
            vbObjectError + 1000, _
            "PCAHedge", _
            "SpreadName cannot be empty."

    End If


    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")


    re.Global = True
    re.IgnoreCase = True


    re.pattern = "[0-9]+(?:\.[0-9]+)?"


    Dim Matches As Object

    Set Matches = re.Execute(txt)


    If Matches.Count < 2 Then

        Err.Raise _
            vbObjectError + 1001, _
            "PCAHedge", _
            "SpreadName must contain at least two tenors."

    End If


    Dim Result() As Double

    ReDim Result(1 To Matches.Count)


    Dim i As Long


    For i = 1 To Matches.Count

        Result(i) = CDbl(Matches(i - 1).value)

    Next i


    ParseSpreadTenors = Result

End Function



' =====================================================================
' ParsePCList
'
' Parse comma-separated PC numbers.
'
'
' Examples
' ---------------------------------------------------------------------
'
'   "1"
'
'       -> {1}
'
'
'   "1,2"
'
'       -> {1,2}
'
'
'   "1,2,5"
'
'       -> {1,2,5}
'
'
'   "1, 2, 5"
'
'       -> {1,2,5}
'
'
' Rules:
'
'   - must be numeric
'   - must be integer
'   - must be >= 1
'
' =====================================================================

Private Function ParsePCList( _
    ByVal PCText As String _
) As Variant


    Dim txt As String

    txt = Trim$(PCText)


    If Len(txt) = 0 Then

        Err.Raise _
            vbObjectError + 1100, _
            "PCAHedge", _
            "HedgePCs cannot be empty."

    End If


    Dim Parts As Variant

    Parts = Split(txt, ",")


    Dim Count As Long

    Count = UBound(Parts) - LBound(Parts) + 1


    Dim Result() As Long

    ReDim Result(1 To Count)


    Dim i As Long

    Dim Item As String

    Dim x As Double


    For i = LBound(Parts) To UBound(Parts)


        Item = Trim$(CStr(Parts(i)))


        ' -------------------------------------------------------------
        ' Empty item
        '
        ' Example:
        '
        '       "1,,2"
        '
        ' -------------------------------------------------------------

        If Len(Item) = 0 Then

            Err.Raise _
                vbObjectError + 1101, _
                "PCAHedge", _
                "Invalid empty PC number."

        End If


        ' -------------------------------------------------------------
        ' Must be numeric
        ' -------------------------------------------------------------

        If Not IsNumeric(Item) Then

            Err.Raise _
                vbObjectError + 1102, _
                "PCAHedge", _
                "Invalid PC number: " & Item

        End If


        x = CDbl(Item)


        ' -------------------------------------------------------------
        ' Must be integer
        ' -------------------------------------------------------------

        If x <> Fix(x) Then

            Err.Raise _
                vbObjectError + 1103, _
                "PCAHedge", _
                "PC numbers must be integers."

        End If


        ' -------------------------------------------------------------
        ' Must be >= 1
        ' -------------------------------------------------------------

        If x < 1 Then

            Err.Raise _
                vbObjectError + 1104, _
                "PCAHedge", _
                "PC numbers must be >= 1."

        End If


        Result(i - LBound(Parts) + 1) = CLng(x)


    Next i


    ParsePCList = Result

End Function



' =====================================================================
' FindTenorRowInMatrix
'
' Search the first column of EigenVectors for a given tenor.
'
'
' Recognised examples
' ---------------------------------------------------------------------
'
'   yield_5y
'
'   yield_10y
'
'   5y
'
'   10Y
'
'   JGB_5Y
'
'   JGB|5Y
'
'   jgb_yield_7y
'
'
' Return values
' ---------------------------------------------------------------------
'
'   > 0
'
'       unique row found
'
'
'   0
'
'       tenor not found
'
'
'   -1
'
'       multiple rows matched
'
'
' Row number is relative to the EigenVectors range.
'
' =====================================================================

Private Function FindTenorRowInMatrix( _
    ByVal EigenVectors As Range, _
    ByVal TargetTenor As Double _
) As Long


    Dim r As Long

    Dim TenorValue As Double

    Dim FoundCount As Long

    Dim FoundRow As Long


    ' Row 1 = PC header row

    For r = 2 To EigenVectors.rows.Count


        TenorValue = ExtractTenor( _
                         EigenVectors.Cells(r, 1).value2 _
                     )


        If TenorValue >= 0# Then


            If Abs(TenorValue - TargetTenor) < 0.0000001 Then


                FoundCount = FoundCount + 1

                FoundRow = r


            End If


        End If


    Next r


    Select Case FoundCount


        Case 0

            FindTenorRowInMatrix = 0


        Case 1

            FindTenorRowInMatrix = FoundRow


        Case Else

            FindTenorRowInMatrix = -1


    End Select

End Function



' =====================================================================
' ExtractTenor
'
' Extract numeric tenor from a row label.
'
'
' Examples
' ---------------------------------------------------------------------
'
'   "yield_5y"
'
'       -> 5
'
'
'   "5y"
'
'       -> 5
'
'
'   "JGB_10Y"
'
'       -> 10
'
'
'   "yield_2.5y"
'
'       -> 2.5
'
'
' If tenor cannot be identified:
'
'       -> -1
'
' =====================================================================

Private Function ExtractTenor( _
    ByVal LabelValue As Variant _
) As Double


    If IsError(LabelValue) Then

        ExtractTenor = -1#

        Exit Function

    End If


    Dim txt As String

    txt = Trim$(CStr(LabelValue))


    If Len(txt) = 0 Then

        ExtractTenor = -1#

        Exit Function

    End If


    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")


    re.Global = False

    re.IgnoreCase = True


    ' Numeric tenor followed by Y/y
    '
    ' Examples:
    '
    '   5y
    '   yield_5y
    '   JGB|10Y
    '   yield_2.5y

    re.pattern = "([0-9]+(?:\.[0-9]+)?)\s*[yY]"


    Dim Matches As Object

    Set Matches = re.Execute(txt)


    If Matches.Count > 0 Then


        ExtractTenor = CDbl( _
                           Matches(0).SubMatches(0) _
                       )


        Exit Function


    End If


    ' Pure numeric labels are also allowed

    If IsNumeric(txt) Then


        ExtractTenor = CDbl(txt)


    Else


        ExtractTenor = -1#


    End If

End Function



' =====================================================================
' FindPCColumn
'
' Search first row of EigenVectors for PC header.
'
'
' Example:
'
'   PCNumber = 1
'
' searches for:
'
'       PC1
'
'
' The following are treated as equivalent:
'
'       PC1
'       PC 1
'
'
' Return values
' ---------------------------------------------------------------------
'
'   > 0
'
'       unique PC column found
'
'
'   0
'
'       PC not found
'
'
'   -1
'
'       duplicate PC headers
'
' =====================================================================

Private Function FindPCColumn( _
    ByVal EigenVectors As Range, _
    ByVal PCNumber As Long _
) As Long


    Dim Target As String


    Target = "PC" & CStr(PCNumber)

    Target = UCase$(Target)


    Dim c As Long

    Dim HeaderText As String

    Dim FoundCount As Long

    Dim FoundCol As Long


    ' Column 1 = tenor names

    For c = 2 To EigenVectors.Columns.Count


        If Not IsError(EigenVectors.Cells(1, c).value2) Then


            HeaderText = CStr( _
                             EigenVectors.Cells(1, c).value2 _
                         )


            HeaderText = UCase$(Trim$(HeaderText))


            ' Ignore spaces

            HeaderText = Replace(HeaderText, " ", "")


            If HeaderText = Target Then


                FoundCount = FoundCount + 1

                FoundCol = c


            End If


        End If


    Next c


    Select Case FoundCount


        Case 0

            FindPCColumn = 0


        Case 1

            FindPCColumn = FoundCol


        Case Else

            FindPCColumn = -1


    End Select

End Function



' =====================================================================
' HasDuplicateLongValues
'
' Returns True if a Long array contains duplicate values.
'
'
' Examples
' ---------------------------------------------------------------------
'
'   {1,2,5}
'
'       -> False
'
'
'   {1,2,2}
'
'       -> True
'
' =====================================================================

Private Function HasDuplicateLongValues( _
    ByVal Values As Variant _
) As Boolean


    Dim i As Long

    Dim j As Long


    For i = LBound(Values) To UBound(Values) - 1


        For j = i + 1 To UBound(Values)


            If CLng(Values(i)) = CLng(Values(j)) Then


                HasDuplicateLongValues = True

                Exit Function


            End If


        Next j


    Next i


    HasDuplicateLongValues = False

End Function

