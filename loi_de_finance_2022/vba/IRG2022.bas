Attribute VB_Name = "IRG2022"
Function IRG(taxableSalary As Variant, Optional handicape As Boolean = False) As Double
    Dim SI As Double: SI = Int(taxableSalary / 10) * 10
    IRG = 0
    If (SI >= 20001 And SI <= 40000) Then
        IRG = (SI - 20000) * 0.23
    End If
    If (SI >= 40001 And SI <= 80000) Then
        IRG = 4600 + (SI - 40000) * 0.27
    End If
    If (SI >= 80001 And SI <= 160000) Then
        IRG = 15400 + (SI - 80000) * 0.3
    End If
    If (SI >= 160001 And SI <= 320000) Then
        IRG = 39400 + (SI - 160000) * 0.33
    End If
    If (SI >= 320001) Then
        IRG = 92200 + (SI - 320000) * 0.35
    End If
    Dim abat As Double
    abat = Application.WorksheetFunction.Max(IRG * 0.4, 1000)
    abat = Application.WorksheetFunction.Min(abat, 1500)
    IRG = IRG - abat
    If (SI <= 30000) Then
        IRG = 0
    End If
    If (SI > 30000 And SI <= 35000 And (Not handicape)) Then
        IRG = IRG * (137 / 51) - (27925 / 8)
    End If
    If (SI > 30000 And SI <= 42500 And handicape) Then
        IRG = IRG * (93 / 61) - (81213 / 41)
    End If
    IRG = Int(IRG * 10) / 10
End Function

