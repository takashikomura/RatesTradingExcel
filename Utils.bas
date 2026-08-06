Option Explicit

'=========================================================
' MatrixTBA
'
' A' × B × A を返す
'
' A : m行×n列
' B : m行×m列
' 出力 : n行×n列
'
' 使用例:
'   =MatrixTBA(TradeMatrix, CovGrid)
'
' J'QgridJ を計算する場合:
'   A = J
'   B = Qgrid
'=========================================================
Public Function MatrixTBA( _
    ByVal A As Range, _
    ByVal B As Range _
) As Variant

    On Error GoTo ErrorHandler

    ' Bは正方行列である必要がある
    If B.rows.count <> B.Columns.count Then
        MatrixTBA = CVErr(xlErrValue)
        Exit Function
    End If

    ' Bの行数・列数はAの行数と一致する必要がある
    If B.rows.count <> A.rows.count Then
        MatrixTBA = CVErr(xlErrRef)
        Exit Function
    End If

    ' A' × B × A
    MatrixTBA = Application.MMult( _
                    Application.Transpose(A.value2), _
                    Application.MMult(B.value2, A.value2) _
                )

    Exit Function

ErrorHandler:
    MatrixTBA = CVErr(xlErrValue)

End Function

