Option Explicit

' 引張を正とするひずみの主値計算。
' gxy は工学せん断ひずみ、theta は反時計回りの radian で返す。
' 等方ひずみ（または全成分ゼロ）では主軸方向を 0 と定義する。

Public Sub FEMStrainPrincipal2D(ByVal ex As Double, ByVal ey As Double, ByVal gxy As Double, _
                                ByRef pMax As Double, ByRef pMin As Double, ByRef theta As Double)
    Dim center As Double, halfDiff As Double, shear As Double
    Dim radius As Double, angle As Double, pi As Double

    ' 半分ずつ計算して、同符号の大きな値の加算による overflow を避ける。
    center = ex * 0.5 + ey * 0.5
    halfDiff = ex * 0.5 - ey * 0.5
    shear = gxy * 0.5
    radius = FEMStrainHypot(halfDiff, shear)
    pMax = center + radius
    pMin = center - radius

    If radius = 0# Then
        theta = 0#
        Exit Sub
    End If

    ' 2 theta = atan2(gxy/2, (ex-ey)/2)。VBA の Atn は atan2 を持たない。
    angle = FEMStrainAtan2(shear, halfDiff)
    pi = 4# * Atn(1#)
    theta = angle * 0.5
    If theta < 0# Then theta = theta + pi
    If theta >= pi Then theta = theta - pi
End Sub

Public Function FEMStrainMaxShear(ByVal ex As Double, ByVal ey As Double, ByVal gxy As Double, _
                                  Optional ByVal ez As Double = 0#) As Double
    Dim pMax As Double, pMin As Double, theta As Double
    Dim high As Double, low As Double
    FEMStrainPrincipal2D ex, ey, gxy, pMax, pMin, theta
    high = pMax
    low = pMin
    If ez > high Then high = ez
    If ez < low Then low = ez
    FEMStrainMaxShear = high - low
End Function

Private Function FEMStrainHypot(ByVal x As Double, ByVal y As Double) As Double
    Dim scaleValue As Double, ax As Double, ay As Double
    ax = Abs(x): ay = Abs(y)
    If ax >= ay Then
        scaleValue = ax
    Else
        scaleValue = ay
    End If
    If scaleValue = 0# Then
        FEMStrainHypot = 0#
    Else
        FEMStrainHypot = scaleValue * Sqr((x / scaleValue) * (x / scaleValue) + (y / scaleValue) * (y / scaleValue))
    End If
End Function

Private Function FEMStrainAtan2(ByVal y As Double, ByVal x As Double) As Double
    Dim pi As Double
    pi = 4# * Atn(1#)
    If Abs(x) > Abs(y) Then
        FEMStrainAtan2 = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then FEMStrainAtan2 = FEMStrainAtan2 + pi Else FEMStrainAtan2 = FEMStrainAtan2 - pi
        End If
    ElseIf y > 0# Then
        FEMStrainAtan2 = pi * 0.5 - Atn(x / y)
    ElseIf y < 0# Then
        FEMStrainAtan2 = -pi * 0.5 - Atn(x / y)
    Else
        FEMStrainAtan2 = 0#
    End If
End Function
