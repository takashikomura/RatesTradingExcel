Option Explicit

' =========================================================
' spd(tenor列, yield列, spread名セル, multiplier)
'
' 例:
'   =spd($A$2:$A$10,$B$2:$B$10,"1s2s",100)
'
' 解釈:
'   1s2s     = Yield(2y) - Yield(1y)
'   1s2s3s   = 2 * Yield(2y) - Yield(1y) - Yield(3y)
'
' tenor列:
'   1y, 2y, 3y, 1m, 6m など
'
' spread名:
'   1s2s, 2s5s, 5s10s20s など
'   s は y と同じく「年」として扱う
'
' multiplier:
'   最終結果に掛ける倍率
' =========================================================
Public Function spd( _
    ByVal TenorRange As Range, _
    ByVal YieldRange As Range, _
    ByVal SpreadName As Variant, _
    Optional ByVal Multiplier As Double = 1 _
) As Variant

    On Error GoTo ErrHandler

    Dim nTenor As Long
    Dim nYield As Long

    nTenor = TenorRange.Cells.count
    nYield = YieldRange.Cells.count

    If nTenor <> nYield Then
        spd = CVErr(xlErrValue)
        Exit Function
    End If

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim tenorKey As String
    Dim y As Variant

    For i = 1 To nTenor
        tenorKey = TenorToMonthKey(TenorRange.Cells(i).value)

        If tenorKey <> "" Then
            y = YieldRange.Cells(i).value

            If IsNumeric(y) Then
                dict(tenorKey) = CDbl(y)
            End If
        End If
    Next i

    Dim legs As Collection
    Set legs = ParseSpreadLegs(CStr(SpreadName))

    If legs.count <> 2 And legs.count <> 3 Then
        spd = CVErr(xlErrValue)
        Exit Function
    End If

    Dim k1 As String, k2 As String, k3 As String
    Dim y1 As Double, y2 As Double, y3 As Double

    If legs.count = 2 Then

        k1 = TenorToMonthKey(legs(1))
        k2 = TenorToMonthKey(legs(2))

        If Not dict.Exists(k1) Or Not dict.Exists(k2) Then
            spd = CVErr(xlErrNA)
            Exit Function
        End If

        y1 = dict(k1)
        y2 = dict(k2)

        spd = Multiplier * (y2 - y1)
        Exit Function

    ElseIf legs.count = 3 Then

        k1 = TenorToMonthKey(legs(1))
        k2 = TenorToMonthKey(legs(2))
        k3 = TenorToMonthKey(legs(3))

        If Not dict.Exists(k1) Or Not dict.Exists(k2) Or Not dict.Exists(k3) Then
            spd = CVErr(xlErrNA)
            Exit Function
        End If

        y1 = dict(k1)
        y2 = dict(k2)
        y3 = dict(k3)

        spd = Multiplier * (2 * y2 - y1 - y3)
        Exit Function

    End If

    spd = CVErr(xlErrValue)
    Exit Function

ErrHandler:
    spd = CVErr(xlErrValue)

End Function


' ---------------------------------------------------------
' tenor文字列を「月数」のキーに変換する
'
' 例:
'   1m  -> "1"
'   6m  -> "6"
'   1y  -> "12"
'   1s  -> "12"
'   10s -> "120"
' ---------------------------------------------------------
Private Function TenorToMonthKey(ByVal v As Variant) As String

    On Error GoTo ErrHandler

    Dim s As String
    s = LCase$(Trim$(CStr(v)))

    If s = "" Then
        TenorToMonthKey = ""
        Exit Function
    End If

    s = Replace(s, " ", "")
    s = Replace(s, "年", "y")
    s = Replace(s, "月", "m")
    s = Replace(s, "yrs", "y")
    s = Replace(s, "yr", "y")
    s = Replace(s, "years", "y")
    s = Replace(s, "year", "y")
    s = Replace(s, "months", "m")
    s = Replace(s, "month", "m")

    Dim unitChar As String
    Dim numPart As String

    unitChar = Right$(s, 1)

    If unitChar = "y" Or unitChar = "s" Or unitChar = "m" Then
        numPart = Left$(s, Len(s) - 1)
    Else
        ' 単位なしの場合は年として扱う
        unitChar = "y"
        numPart = s
    End If

    If Not IsNumeric(numPart) Then
        TenorToMonthKey = ""
        Exit Function
    End If

    Dim x As Double
    x = CDbl(numPart)

    If unitChar = "m" Then
        TenorToMonthKey = CStr(CLng(x))
    Else
        TenorToMonthKey = CStr(CLng(x * 12))
    End If

    Exit Function

ErrHandler:
    TenorToMonthKey = ""

End Function


' ---------------------------------------------------------
' spread名をlegに分解する
'
' 例:
'   1s2s       -> 1s, 2s
'   5s10s20s   -> 5s, 10s, 20s
'   1m3m       -> 1m, 3m
' ---------------------------------------------------------
Private Function ParseSpreadLegs(ByVal spreadText As String) As Collection

    Dim legs As New Collection

    Dim s As String
    s = LCase$(Trim$(spreadText))

    s = Replace(s, " ", "")
    s = Replace(s, "-", "")
    s = Replace(s, "_", "")
    s = Replace(s, "/", "")

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Global = True
        .IgnoreCase = True
        .Pattern = "([0-9]+(?:\.[0-9]+)?)([sm y]?)"
    End With

    ' 空白除去後なので [sm y] の y は実質 y を拾う
    re.Pattern = "([0-9]+(?:\.[0-9]+)?)([smy]?)"

    Dim matches As Object
    Set matches = re.Execute(s)

    Dim m As Object
    Dim leg As String
    Dim numPart As String
    Dim unitPart As String

    For Each m In matches
        numPart = m.SubMatches(0)
        unitPart = m.SubMatches(1)

        If unitPart = "" Then
            unitPart = "s"
        End If

        leg = numPart & unitPart
        legs.Add leg
    Next m

    Set ParseSpreadLegs = legs

End Function

