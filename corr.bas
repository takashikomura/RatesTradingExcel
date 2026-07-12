Option Explicit

'========================================================
' CovToCorr
'
' 共分散行列を相関行列へ変換するUDF
'
' 使用例:
'   =CovToCorr(CovMatrix)
'
' Excel 365:
'   1セルに入力すると相関行列がスピル出力される。
'
' 旧Excel:
'   出力範囲を事前選択し、配列数式として確定する。
'
' 入力条件:
'   - 正方行列であること
'   - 対角成分が各系列の分散であること
'
' 欠損処理:
'   - 共分散セルがエラーの場合、その相関セルは #N/A
'   - 対応する分散がエラー、非数値、0以下の場合も #N/A
'
' 戻り値:
'   入力行列と同じ行数・列数の相関行列
'========================================================
Public Function CovToCorr( _
    ByVal covInput As Variant _
) As Variant

    Dim covData As Variant
    Dim result() As Variant

    Dim nRows As Long
    Dim nCols As Long

    Dim i As Long
    Dim j As Long

    Dim covIJ As Variant
    Dim varI As Variant
    Dim varJ As Variant

    Dim denominator As Double
    Dim corrValue As Double

    On Error GoTo FunctionError

    Application.Volatile True

    '----------------------------------------------------
    ' Rangeまたは配列をVariant配列へ変換
    '----------------------------------------------------
    If TypeName(covInput) = "Range" Then
        covData = covInput.Value2
    Else
        covData = covInput
    End If

    '----------------------------------------------------
    ' 入力が2次元配列か確認
    '----------------------------------------------------
    If Not IsArray(covData) Then
        CovToCorr = CVErr(xlErrValue)
        Exit Function
    End If

    On Error GoTo InvalidArray

    nRows = UBound(covData, 1) - LBound(covData, 1) + 1
    nCols = UBound(covData, 2) - LBound(covData, 2) + 1

    On Error GoTo FunctionError

    '----------------------------------------------------
    ' 正方行列の確認
    '----------------------------------------------------
    If nRows <> nCols Then
        CovToCorr = CVErr(xlErrValue)
        Exit Function
    End If

    If nRows < 1 Then
        CovToCorr = CVErr(xlErrValue)
        Exit Function
    End If

    ReDim result(1 To nRows, 1 To nCols)

    '----------------------------------------------------
    ' 相関行列計算
    '----------------------------------------------------
    For i = 1 To nRows

        varI = covData( _
            LBound(covData, 1) + i - 1, _
            LBound(covData, 2) + i - 1 _
        )

        For j = 1 To nCols

            covIJ = covData( _
                LBound(covData, 1) + i - 1, _
                LBound(covData, 2) + j - 1 _
            )

            varJ = covData( _
                LBound(covData, 1) + j - 1, _
                LBound(covData, 2) + j - 1 _
            )

            '共分散または分散がExcelエラーなら #N/A
            If IsError(covIJ) _
               Or IsError(varI) _
               Or IsError(varJ) Then

                result(i, j) = CVErr(xlErrNA)

            '数値でなければ #N/A
            ElseIf Not IsUsableNumericValue_Local(covIJ) _
               Or Not IsUsableNumericValue_Local(varI) _
               Or Not IsUsableNumericValue_Local(varJ) Then

                result(i, j) = CVErr(xlErrNA)

            '分散が0以下なら相関を定義できない
            ElseIf CDbl(varI) <= 0# _
               Or CDbl(varJ) <= 0# Then

                result(i, j) = CVErr(xlErrNA)

            Else

                denominator = Sqr(CDbl(varI) * CDbl(varJ))

                If denominator <= 0# Then

                    result(i, j) = CVErr(xlErrNA)

                Else

                    corrValue = CDbl(covIJ) / denominator

                    '浮動小数点誤差により1をわずかに超える場合を補正
                    If corrValue > 1# And corrValue < 1# + 0.0000000001 Then
                        corrValue = 1#
                    ElseIf corrValue < -1# And _
                           corrValue > -1# - 0.0000000001 Then
                        corrValue = -1#
                    End If

                    '明確に[-1,1]を超える場合は入力行列が不整合
                    If corrValue > 1# Or corrValue < -1# Then
                        result(i, j) = CVErr(xlErrNum)
                    Else
                        result(i, j) = corrValue
                    End If

                End If

            End If

        Next j

    Next i

    CovToCorr = result
    Exit Function

InvalidArray:

    CovToCorr = CVErr(xlErrValue)
    Exit Function

FunctionError:

    CovToCorr = CVErr(xlErrValue)

End Function

'========================================================
' 数値として相関計算に使用可能か判定
'
' False:
'   - Error
'   - Empty
'   - Null
'   - 空文字
'   - 非数値
'
' True:
'   - 数値
'   - 数値文字列
'========================================================
Private Function IsUsableNumericValue_Local( _
    ByVal valueInput As Variant _
) As Boolean

    IsUsableNumericValue_Local = False

    If IsError(valueInput) Then Exit Function
    If IsNull(valueInput) Then Exit Function
    If IsEmpty(valueInput) Then Exit Function

    If VarType(valueInput) = vbString Then

        If Len(Trim$(CStr(valueInput))) = 0 Then
            Exit Function
        End If

    End If

    If Not IsNumeric(valueInput) Then Exit Function

    IsUsableNumericValue_Local = True

End Function

