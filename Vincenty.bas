Attribute VB_Name = "Vincenty"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Vincenty's Direct and Inverse Solution of Geodesics on the Ellipsoid
' algorithms by Thaddeus Vincenty (1975)
' https://en.wikipedia.org/wiki/Vincenty%27s_formulae
' https://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf
' https://geographiclib.sourceforge.io/geodesic-papers/vincenty75b.pdf
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Ported to VBA by (c) Tomasz Jastrzebski 2018-2019 MIT Licence
' Version: 2019-07-26
' Latest version available at:
' https://github.com/tdjastrzebski/Vincenty-Excel
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Based on the implementation by Chris Veness, ver 2.2.0
' https://www.movable-type.co.uk/scripts/latlong-vincenty.html
' https://github.com/chrisveness/geodesy
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const PI = 3.14159265358979
Private Const EPSILON12 As Double = 0.000000000001 ' 1E-12
Private Const EPSILON16 As Double = 2 ^ -52 ' ~2.2E-16
' WGS-84 ellipsiod
Private Const low_a As Double = 6378137
Private Const low_b As Double = 6356752.3142
Private Const f As Double = 1 / 298.257223563
Private Const MaxIterations As Integer = 100

Private Type DirParams
    sinU1 As Double
    cosU1 As Double
    sigma As Double
    sinSigma As Double
    cosSigma As Double
    cosAlpha1 As Double
    cosSqAlpha As Double
    cos2sigmaM As Double
    sinAlpha As Double
    sinAlpha1 As Double
    lambda1 As Double
End Type

Private Type InvParams
    upper_A As Double
    sigma As Double
    sinSqSigma As Double
    deltaSigma As Double
    sinU1 As Double
    cosU1 As Double
    sinU2 As Double
    cosU2 As Double
    cosLambda As Double
    sinLambda As Double
    s As Double
End Type

' Calculates geodesic latitude (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirLat(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As Variant
Attribute VincentyDirLat.VB_Description = "Calculates geodesic latitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirLat.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: p = VincentyDir(lat, lon, azimuth, distance)
    
    Dim x As Double: x = p.sinU1 * p.sinSigma - p.cosU1 * p.cosSigma * p.cosAlpha1
    Dim phi2 As Double: phi2 = Atan2(p.sinU1 * p.cosSigma + p.cosU1 * p.sinSigma * p.cosAlpha1, (1 - f) * Sqr(p.sinAlpha * p.sinAlpha + x * x))
 
    VincentyDirLat = ToDegrees(phi2)
    Exit Function
error:
If Err.Number = Excel.xlErrNA Then
    VincentyDirLat = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

' Calculates geodesic longitude (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirLon(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As Variant
Attribute VincentyDirLon.VB_Description = "Calculates geodesic longitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirLon.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: p = VincentyDir(lat, lon, azimuth, distance)
    
    Dim lambda As Double: lambda = Atan2(p.sinSigma * p.sinAlpha1, p.cosU1 * p.cosSigma - p.sinU1 * p.sinSigma * p.cosAlpha1)
    Dim C As Double: C = f / 16 * p.cosSqAlpha * (4 + f * (4 - 3 * p.cosSqAlpha))
    Dim fix1 As Double: fix1 = p.cos2sigmaM + C * p.cosSigma * (-1 + 2 * p.cos2sigmaM * p.cos2sigmaM)
    Dim L As Double: L = lambda - (1 - C) * f * p.sinAlpha * (p.sigma + C * p.sinSigma * fix1)
    
    Dim lambda2 As Double: lambda2 = p.lambda1 + L
    
    If lambda2 = PI Then
        VincentyDirLon = 180
    Else
        VincentyDirLon = NormalizeLon(ToDegrees(lambda2))
    End If
    
    Exit Function
error:
If Err.Number = Excel.xlErrNA Then
    VincentyDirLon = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

' Calculates geodesic reverse azimuth (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirRevAzimuth(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double, Optional returnAzimuth As Boolean = False) As Variant
Attribute VincentyDirRevAzimuth.VB_Description = "Calculates geodesic azimuth in degrees clockwise from north based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirRevAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: p = VincentyDir(lat, lon, azimuth, distance)

    Dim x As Double: x = p.sinU1 * p.sinSigma - p.cosU1 * p.cosSigma * p.cosAlpha1
    Dim alpha2 As Double: alpha2 = Atan2(p.sinAlpha, -x)
    
    If returnAzimuth Then
        VincentyDirRevAzimuth = NormalizeAzimuth(ToDegrees(alpha2) + 180, True)
    Else
        VincentyDirRevAzimuth = NormalizeAzimuth(ToDegrees(alpha2), True)

    End If

    Exit Function
error:
If Err.Number = Excel.xlErrNA Then
    VincentyDirRevAzimuth = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

' Calculates geodesic distance (in m) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvDistance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Variant
Attribute VincentyInvDistance.VB_Description = "Calculates geodesic distance in meters between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvDistance.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: p = VincentyInv(lat1, lon1, lat2, lon2)
    
    If Abs(p.s) < EPSILON16 Then
        VincentyInvDistance = CVErr(Excel.xlErrNA): Exit Function
    Else
        VincentyInvDistance = p.s
    End If
    
    Exit Function
error:
If Err.Number = Excel.xlErrDiv0 Then
    VincentyInvDistance = 0
ElseIf Err.Number = Excel.xlErrNA Then
    VincentyInvDistance = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

' Calculates geodesic azimuth (in degrees) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvFwdAzimuth(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Variant
Attribute VincentyInvFwdAzimuth.VB_Description = "Calculates geodesic forward azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvFwdAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: p = VincentyInv(lat1, lon1, lat2, lon2)
    
    If Abs(p.s) < EPSILON16 Then
        VincentyInvFwdAzimuth = CVErr(Excel.xlErrNA): Exit Function
    End If
    
    Dim fwdAz As Double
    
    If Abs(p.sinSqSigma) < EPSILON16 Then
        ' special handling of exactly antipodal points where sinSigma = 0
        fwdAz = 0
    Else
        fwdAz = Atan2(p.cosU2 * p.sinLambda, p.cosU1 * p.sinU2 - p.sinU1 * p.cosU2 * p.cosLambda)
    End If
    
    VincentyInvFwdAzimuth = NormalizeAzimuth(ToDegrees(fwdAz), True)
    Exit Function
error:
If Err.Number = Excel.xlErrDiv0 Then
    VincentyInvFwdAzimuth = CVErr(Excel.xlErrNull)
ElseIf Err.Number = Excel.xlErrNA Then
    VincentyInvFwdAzimuth = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

' Calculates geodesic reverse azimuth (in degrees) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvRevAzimuth(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double, Optional returnAzimuth As Boolean = False) As Variant
Attribute VincentyInvRevAzimuth.VB_Description = "Calculates geodesic reverse azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvRevAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: p = VincentyInv(lat1, lon1, lat2, lon2)
    
    If Abs(p.s) < EPSILON16 Then
        VincentyInvRevAzimuth = CVErr(Excel.xlErrNA): Exit Function
    End If
    
    Dim revAz As Double
    
    If Abs(p.sinSqSigma) < EPSILON16 Then
        ' special handling of exactly antipodal points where sinSigma = 0
        revAz = PI
    Else
        revAz = Atan2(p.cosU1 * p.sinLambda, -p.sinU1 * p.cosU2 + p.cosU1 * p.sinU2 * p.cosLambda)
    End If
    
    If returnAzimuth Then
        VincentyInvRevAzimuth = NormalizeAzimuth(ToDegrees(revAz) + 180, True)
    Else
        VincentyInvRevAzimuth = NormalizeAzimuth(ToDegrees(revAz), True)
    End If
    Exit Function
error:
If Err.Number = Excel.xlErrDiv0 Then
    VincentyInvRevAzimuth = CVErr(Excel.xlErrNull)
ElseIf Err.Number = Excel.xlErrNA Then
    VincentyInvRevAzimuth = CVErr(Excel.xlErrNA)
Else
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End If
End Function

Private Function VincentyDir(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As DirParams
    Dim p As DirParams
    
    Dim phi1 As Double: phi1 = ToRadians(lat)
    p.lambda1 = ToRadians(lon)
    Dim alpha1 As Double: alpha1 = ToRadians(azimuth)
    Dim s As Double: s = distance
    
    Dim fix1 As Double ' temp variable to prevent "formula too complex.." error
    Dim fix2 As Double ' temp variable to prevent "formula too complex.." error
    
    p.sinAlpha1 = Sin(alpha1)
    p.cosAlpha1 = Cos(alpha1)

    Dim tanU1 As Double: tanU1 = (1 - f) * Tan(phi1)
    p.cosU1 = 1 / Sqr((1 + tanU1 ^ 2))
    p.sinU1 = tanU1 * p.cosU1
    Dim sigma1 As Double: sigma1 = Atan2(tanU1, p.cosAlpha1) ' sigma1 = angular distance on the sphere from the equator to P1
    p.sinAlpha = p.cosU1 * p.sinAlpha1 ' Alpha = azimuth of the geodesic at the equator
    p.cosSqAlpha = 1 - p.sinAlpha ^ 2
    fix1 = low_a ^ 2 - low_b ^ 2
    Dim uSq As Double: uSq = p.cosSqAlpha * fix1 / (low_b ^ 2)
    fix1 = -768 + uSq * (320 - 175 * uSq)
    Dim A As Double: A = 1 + uSq / 16384 * (4096 + uSq * fix1)
    fix1 = -128 + uSq * (74 - 47 * uSq)
    Dim B As Double: B = uSq / 1024 * (256 + uSq * fix1)
    
    p.sigma = s / (low_b * A)
    Dim deltaSigma As Double
    Dim sigma2 As Double
    Dim iterationCount As Integer:  iterationCount = 0
    
    Do
        p.cos2sigmaM = Cos(2 * sigma1 + p.sigma)
        p.sinSigma = Sin(p.sigma)
        p.cosSigma = Cos(p.sigma)
        deltaSigma = B * p.sinSigma * (p.cos2sigmaM + B / 4 * (p.cosSigma * (-1 + 2 * p.cos2sigmaM * p.cos2sigmaM) - _
            B / 6 * p.cos2sigmaM * (-3 + 4 * p.sinSigma * p.sinSigma) * (-3 + 4 * p.cos2sigmaM * p.cos2sigmaM)))
        sigma2 = p.sigma
        p.sigma = s / (low_b * A) + deltaSigma
        iterationCount = iterationCount + 1
    Loop While Abs(p.sigma - sigma2) > EPSILON12 And iterationCount < MaxIterations
    
    If iterationCount >= MaxIterations Then
        ' failed to converge
        Err.Raise (Excel.xlErrNA): Exit Function
    End If
        
    VincentyDir = p
End Function

Private Function VincentyInv(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As InvParams
    Dim p As InvParams
    Dim sinSigma As Double
    Dim cosSigma As Double
    Dim sinAlpha As Double
    Dim cosSqAlpha As Double
    Dim cos2sigmaM As Double
    Dim C As Double
    Dim uSq As Double
    Dim upper_B As Double
    Dim fix2 As Double ' temp variable to prevent "formula too complex.." error
    Dim fix1 As Double ' temp variable to prevent "formula too complex.." error
    Dim iterationCount As Integer: iterationCount = 0
    
    lat1 = ToRadians(lat1)
    lat2 = ToRadians(lat2)
    
    Dim L As Double: L = ToRadians(lon2 - lon1)
    
    Dim tanU1 As Double: tanU1 = (1 - f) * Tan(lat1)
    p.cosU1 = 1 / Sqr(1 + (tanU1 ^ 2))
    p.sinU1 = tanU1 * p.cosU1
    
    Dim tanU2 As Double: tanU2 = (1 - f) * Tan(lat2)
    p.cosU2 = 1 / Sqr(1 + (tanU2 ^ 2))
    p.sinU2 = tanU2 * p.cosU2
    
    Dim antipodal As Boolean: antipodal = Abs(L) > PI / 2 Or Abs(lat2 - lat1) > PI / 2
    
    Dim lambda As Double: lambda = L
    p.sigma = IIf(antipodal, PI, 0)
    cosSigma = IIf(antipodal, -1, 1)
    cos2sigmaM = 1
    cosSqAlpha = 1
    Dim lambdaP As Double
        
    Do
        p.sinLambda = Sin(lambda)
        p.cosLambda = Cos(lambda)
        p.sinSqSigma = ((p.cosU2 * p.sinLambda) ^ 2) + ((p.cosU1 * p.sinU2 - p.sinU1 * p.cosU2 * p.cosLambda) ^ 2)
        If Abs(p.sinSqSigma) < EPSILON16 Then Exit Do  ' co-incident points/antipodal points
        sinSigma = Sqr(p.sinSqSigma)
        cosSigma = p.sinU1 * p.sinU2 + p.cosU1 * p.cosU2 * p.cosLambda
        p.sigma = Atan2(sinSigma, cosSigma)
        sinAlpha = p.cosU1 * p.cosU2 * p.sinLambda / sinSigma
        cosSqAlpha = 1 - (sinAlpha ^ 2)
        
        If cosSqAlpha <> 0 Then
            cos2sigmaM = cosSigma - 2 * p.sinU1 * p.sinU2 / cosSqAlpha
        Else
            cos2sigmaM = 0 ' // on equatorial line cosSqAlpha = 0 (par 6)
        End If

        C = f / 16 * cosSqAlpha * (4 + f * (4 - 3 * cosSqAlpha))
        lambdaP = lambda
        
        fix1 = cos2sigmaM + C * cosSigma * (-1 + 2 * (cos2sigmaM ^ 2))
        lambda = L + (1 - C) * f * sinAlpha * (p.sigma + C * sinSigma * fix1)
        
        Dim iterationCheck As Double: iterationCheck = IIf(antipodal, Abs(lambda) - PI, Abs(lambda))
        
        If iterationCheck > PI Then
            Err.Raise (Excel.xlErrNA): Exit Function
        End If
        
        iterationCount = iterationCount + 1
    Loop While Abs(lambda - lambdaP) > EPSILON12 And iterationCount < MaxIterations
    
    If iterationCount >= MaxIterations Then
        ' failed to converge
        Err.Raise (Excel.xlErrNA): Exit Function
    End If

    uSq = cosSqAlpha * (low_a ^ 2 - low_b ^ 2) / (low_b ^ 2)
    
    fix1 = -768 + uSq * (320 - 175 * uSq)
    p.upper_A = 1 + uSq / 16384 * (4096 + uSq * fix1)
    
    fix1 = -128 + uSq * (74 - 47 * uSq)
    upper_B = uSq / 1024 * (256 + uSq * fix1)
    
    fix1 = cosSigma * (-1 + 2 * cos2sigmaM ^ 2)
    fix2 = upper_B / 6 * cos2sigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2sigmaM ^ 2)
    
    p.deltaSigma = upper_B * sinSigma * (cos2sigmaM + upper_B / 4 * (fix1 - fix2))
    p.s = low_b * p.upper_A * (p.sigma - p.deltaSigma)
    
    VincentyInv = p
End Function

' Converts decimal latitude, longitude or azimuth value to degrees/minutes/seconds string format
Public Function ConvertDegrees(ByVal decimalDeg As Double, Optional isLongitude As Variant) As String
Attribute ConvertDegrees.VB_Description = "Converts latitude, longitude or azimuth in decimal degrees to string in degrees/minutes/seconds format"
Attribute ConvertDegrees.VB_ProcData.VB_Invoke_Func = " \n20"
    If Not IsMissing(isLongitude) And CBool(isLongitude) Then
        decimalDeg = NormalizeLon(decimalDeg)
    ElseIf Not IsMissing(isLongitude) And Not CBool(isLongitude) Then
        decimalDeg = NormalizeLat(decimalDeg)
    Else
        decimalDeg = NormalizeAzimuth(decimalDeg, False)
    End If
    
    Dim s As Integer: s = Sign(decimalDeg)
    decimalDeg = Abs(decimalDeg)
    Dim degrees As Integer: degrees = Fix(decimalDeg)
    Dim minutes As Integer: minutes = Fix((decimalDeg - degrees) * 60)
    Dim seconds As Double: seconds = Round((decimalDeg - degrees - (minutes / 60)) * 60 * 60, 4) ' 4 digit precision corresponds to ~3mm
            
    If Not IsMissing(isLongitude) And Not CBool(isLongitude) Then
        ConvertDegrees = Format$(degrees, "00") & "°" & Format$(minutes, "00") & "'" & Format$(seconds, "00.0000") + Chr(34)
    Else
        ConvertDegrees = Format$(degrees, "000") & "°" & Format$(minutes, "00") & "'" & Format$(seconds, "00.0000") + Chr(34)
    End If
    
    If decimalDeg = 0 Then
        ' do nothing
    ElseIf IsMissing(isLongitude) Then
        If s = -1 Then ConvertDegrees = "-" + ConvertDegrees
    ElseIf isLongitude Then
        If s = 1 Then
            ConvertDegrees = ConvertDegrees + "E"
        ElseIf s = -1 Then
            ConvertDegrees = ConvertDegrees + "W"
        End If
    Else
        If s = 1 Then
            ConvertDegrees = ConvertDegrees + "N"
        ElseIf s = -1 Then
            ConvertDegrees = ConvertDegrees + "S"
        End If
    End If
End Function

' Converts latitude, longitude or azimuth string in degrees/minutes/seconds format to decimal value
Public Function ConvertDecimal(degreeDeg As String) As Variant
Attribute ConvertDecimal.VB_Description = "Converts latitude, longitude or azimuth in degrees/minutes/seconds format to decimal value"
Attribute ConvertDecimal.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    degreeDeg = Replace$(degreeDeg, ChrW(8243), " ") ' double quote
    degreeDeg = Replace$(degreeDeg, ChrW(8242), " ") ' single quote
    degreeDeg = Replace$(degreeDeg, "''", " ") ' double quote
    degreeDeg = Replace$(degreeDeg, """", " ") ' double quote
    degreeDeg = Replace$(degreeDeg, "'", " ") ' single quote
    degreeDeg = Replace$(degreeDeg, "°", " ") ' ordinal indicator
    degreeDeg = Replace$(degreeDeg, Chr(248), " ") ' degree symbol
    degreeDeg = Replace$(degreeDeg, ":", " ")
    degreeDeg = Replace$(degreeDeg, "*", " ")
    degreeDeg = Trim$(degreeDeg)
    
    Dim lc As String: lc = Right$(degreeDeg, 1) ' the last character
    Dim fc As String: fc = Left$(degreeDeg, 1) ' the first character
    Dim s As Integer: s = 1  ' sign

    If Not IsNumeric(fc) And Not IsNumeric(lc) And fc <> "-" Then
        ConvertDecimal = CVErr(Excel.xlErrNA): Exit Function
    ElseIf Not IsNumeric(lc) Then
        degreeDeg = Left$(degreeDeg, Len(degreeDeg) - 1) ' trim the last char
        degreeDeg = Trim$(degreeDeg)
        
        Select Case lc
        Case "W", "w", "S", "s"
            s = -1
        Case "E", "e", "N", "n"
            s = 1
        Case Else
            ConvertDecimal = CVErr(Excel.xlErrNA): Exit Function
        End Select
    ElseIf Not IsNumeric(fc) And fc <> "-" Then
        degreeDeg = Right$(degreeDeg, Len(degreeDeg) - 1) ' trim the first char
        degreeDeg = Trim$(degreeDeg)
        
        Select Case fc
        Case "W", "w", "S", "s"
            s = -1
        Case "E", "e", "N", "n"
            s = 1
        Case Else
            ConvertDecimal = CVErr(Excel.xlErrNA): Exit Function
        End Select
    End If
    
    Dim temp As String
    
    ' remove multiple spaces
    Do
        temp = degreeDeg
        degreeDeg = Replace$(degreeDeg, Space(2), Space(1))
    Loop Until Len(temp) = Len(degreeDeg)
    
    Dim A() As String: A = Split(degreeDeg, " ")
    Dim L As Integer: L = UBound(A) ' length
    
    Dim degrees As Double: degrees = val(A(0))
    Dim minutes As Double: If L > 0 Then minutes = val(A(1)): minutes = minutes / 60
    Dim seconds As Double: If L > 1 Then seconds = val(A(2)): seconds = seconds / 3600
    
    ConvertDecimal = (degrees + (Sign(degrees) * minutes) + (Sign(degrees) * Sign(minutes) * seconds)) * s
    Exit Function
error:
    ConvertDecimal = CVErr(Excel.xlErrNA)
End Function

Private Function Sign(val As Double) As Integer
    Sign = IIf(val >= 0, 1, -1)
End Function

Private Function ToRadians(ByVal degrees As Double) As Double
    ToRadians = degrees * (PI / 180)
End Function

Private Function ToDegrees(ByVal radians As Double) As Double
    ToDegrees = (radians * 180) / PI
End Function

Private Function ModDouble(ByVal dividend As Double, ByVal divisor As Double, Optional sameSignAsDivisor As Boolean = False) As Double
    ' http://en.wikipedia.org/wiki/Modulo_operation
    If sameSignAsDivisor Then
        ModDouble = dividend - (divisor * Int(dividend / divisor))
    Else
        ModDouble = dividend - (divisor * Fix(dividend / divisor))
    End If
    
    ' this function can only be accurate when (a / b) is outside [-2.22E-16,+2.22E-16]
    ' without this correction, ModDouble(.66, .06) = 5.55111512312578E-17 when it should be 0
    ' http://en.wikipedia.org/wiki/Machine_epsilon
    If ModDouble >= -EPSILON16 And ModDouble <= EPSILON16 Then '+/- 2.22E-16
        ModDouble = 0
    End If
End Function

Public Function NormalizeLat(ByVal lat As Double) As Double
    NormalizeLat = Abs(ModDouble(lat - 90, 360, True) - 180) - 90
End Function

Public Function NormalizeLon(ByVal lon As Double) As Double
    NormalizeLon = 2 * ModDouble((lon / 2) + 90, 180, True) - 180
End Function

Public Function NormalizeAzimuth(ByVal azimuth As Double, Optional positiveOnly As Boolean = False) As Double
    NormalizeAzimuth = ModDouble(azimuth, 360, positiveOnly)
End Function

' source: http://en.wikibooks.org/wiki/Programming:Visual_Basic_Classic/Simple_Arithmetic#Trigonometrical_Functions
' note: x & y are in reverse order to match JavaScript Math.atan2() params order
Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    If y > 0 Then
        If x >= y Then
            Atan2 = Atn(y / x)
        ElseIf x <= -y Then
            Atan2 = Atn(y / x) + PI
        Else
            Atan2 = PI / 2 - Atn(x / y)
        End If
    Else
        If x >= -y Then
            Atan2 = Atn(y / x)
        ElseIf x <= y Then
            Atan2 = Atn(y / x) - PI
        Else
            Atan2 = -Atn(x / y) - PI / 2
        End If
    End If
End Function

Public Sub Workbook_Open()
    Application.MacroOptions Macro:="VincentyDirLat", Description:="Calculates geodesic latitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="VincentyDirLon", Description:="Calculates geodesic longitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="VincentyDirRevAzimuth", Description:="Calculates geodesic reverse azimuth (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids. Note: by default aziumuth from point 1 to point 2 at point 2 is returned. To obtain azimuth from point 2 to point 1 pass returnAzimuth = true.", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="VincentyInvDistance", Description:="Calculates geodesic distance in meters between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="VincentyInvFwdAzimuth", Description:="Calculates geodesic forward azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="VincentyInvRevAzimuth", Description:="Calculates geodesic reverse azimuth (in degrees) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids. Note: by default aziumuth from point 1 to point 2 at point 2 is returned. To obtain azimuth from point 2 to point 1 pass returnAzimuth = true.", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="ConvertDecimal", Description:="Converts latitude, longitude or azimuth in degrees/minutes/seconds format to decimal value", _
    ArgumentDescriptions:=Array("string in degrees/minutes/seconds format"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="ConvertDegrees", Description:="Converts latitude, longitude or azimuth string in degrees/minutes/seconds format to decimal value. This function has been designed to parse typical formats.", _
    ArgumentDescriptions:=Array("decimal degrees", "optional: longitude (true) or latitude (false)"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="NormalizeLat", Description:="Normalizes latitude to -90..+90 range", _
    Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="NormalizeLon", Description:="Normalizes longitude to -180..+180 range", _
    Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="NormalizeAzimuth", Description:="Normalizes azimuth to 0..360 range. Note: by default input and return values have the same sign. To obtain only positive values pass positiveOnly = true", _
    ArgumentDescriptions:=Array("azimuth", "optional: positive only , default true"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
End Sub
