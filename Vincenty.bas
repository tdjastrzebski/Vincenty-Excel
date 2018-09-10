Attribute VB_Name = "Vincenty"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Vincenty's Direct and Inverse Solution of Geodesics on the Ellipsoid
' algorithms by Thaddeus Vincenty (1975)
' https://en.wikipedia.org/wiki/Vincenty%27s_formulae
' https://www.ngs.noaa.gov/PUBS_LIB/inverse.pdf
' https://geographiclib.sourceforge.io/geodesic-papers/vincenty75b.pdf
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Ported to VBA by (c) Tomasz Jastrzebski 2018 MIT Licence
' https://github.com/tdjastrzebski/VincentyExcel
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Based on the implementation by Chris Veness
' https://www.movable-type.co.uk/scripts/latlong-vincenty.html
' https://github.com/chrisveness/geodesy
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const PI = 3.14159265358979
Private Const EPSILON As Double = 0.000000000001
' WGS-84 ellipsiod
Private Const low_a As Double = 6378137
Private Const low_b As Double = 6356752.3142
Private Const f As Double = 1 / 298.257223563
Private Const MaxIterations As Integer = 20

' Calculates geodesic latitude (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirLat(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As Variant
Attribute VincentyDirLat.VB_Description = "Calculates geodesic latitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirLat.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: Set p = VincentyDir(lat, lon, azimuth, distance)
    
    Dim x As Double: x = p.sinU1 * p.sinSigma - p.cosU1 * p.cosSigma * p.cosAlpha1
    Dim phi2 As Double: phi2 = Atan2((1 - f) * Sqr(p.sinAlpha * p.sinAlpha + x * x), p.sinU1 * p.cosSigma + p.cosU1 * p.sinSigma * p.cosAlpha1)
 
    VincentyDirLat = ToDegrees(phi2)
    Exit Function
error:
If err.number = Excel.xlErrNA Then
    VincentyDirLat = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

' Calculates geodesic longitude (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirLon(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As Variant
Attribute VincentyDirLon.VB_Description = "Calculates geodesic longitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirLon.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: Set p = VincentyDir(lat, lon, azimuth, distance)
    
    Dim lambda As Double: lambda = Atan2(p.cosU1 * p.cosSigma - p.sinU1 * p.sinSigma * p.cosAlpha1, p.sinSigma * p.sinAlpha1)
    Dim c As Double: c = f / 16 * p.cosSqAlpha * (4 + f * (4 - 3 * p.cosSqAlpha))
    Dim fix1 As Double: fix1 = p.cos2sigmaM + c * p.cosSigma * (-1 + 2 * p.cos2sigmaM * p.cos2sigmaM)
    Dim l As Double: l = lambda - (1 - c) * f * p.sinAlpha * (p.sigma + c * p.sinSigma * fix1)
    
    Dim lambda2 As Double: lambda2 = p.lambda1 + l
    
    If lambda2 = PI Then
        VincentyDirLon = 180
    Else
        lambda2 = ModDouble(lambda2 + 3 * PI, 2 * PI) - PI ' normalise to -180..+180
        VincentyDirLon = ToDegrees(lambda2)
    End If
    
    Exit Function
error:
If err.number = Excel.xlErrNA Then
    VincentyDirLon = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

' Calculates geodesic reverse azimuth (in degrees) based on one point, bearing (in degrees) and distance (in m) using Vincenty's direct formula for ellipsoids
Public Function VincentyDirRevAzimuth(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As Variant
Attribute VincentyDirRevAzimuth.VB_Description = "Calculates geodesic azimuth in degrees clockwise from north based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids"
Attribute VincentyDirRevAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As DirParams: Set p = VincentyDir(lat, lon, azimuth, distance)
    
    Dim x As Double: x = p.sinU1 * p.sinSigma - p.cosU1 * p.cosSigma * p.cosAlpha1
    Dim alpha2 As Double: alpha2 = Atan2(-x, p.sinAlpha)
    alpha2 = ModDouble(alpha2 + 2 * PI, 2 * PI) ' normalise to 0..360
    VincentyDirRevAzimuth = ToDegrees(alpha2)
    Exit Function
error:
If err.number = Excel.xlErrNA Then
    VincentyDirRevAzimuth = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

' Calculates geodesic distance (in m) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvDistance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Variant
Attribute VincentyInvDistance.VB_Description = "Calculates geodesic distance in meters between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvDistance.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: Set p = VincentyInv(lat1, lon1, lat2, lon2)
    VincentyInvDistance = low_b * p.upper_A * (p.sigma - p.deltaSigma)
    Exit Function
error:
If err.number = Excel.xlErrDiv0 Then
    VincentyInvDistance = 0
ElseIf err.number = Excel.xlErrNA Then
    VincentyInvDistance = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

' Calculates geodesic azimuth (in degrees) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvFwdAzimuth(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Variant
Attribute VincentyInvFwdAzimuth.VB_Description = "Calculates geodesic forward azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvFwdAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: Set p = VincentyInv(lat1, lon1, lat2, lon2)
    Dim fwdAz As Double: fwdAz = Atan2(p.cosU1 * p.sinU2 - p.sinU1 * p.cosU2 * p.cosLambda, p.cosU2 * p.sinLambda)
    VincentyInvFwdAzimuth = ToDegrees(fwdAz)
    If VincentyInvFwdAzimuth < 0 Then VincentyInvFwdAzimuth = VincentyInvFwdAzimuth + 360
    Exit Function
error:
If err.number = Excel.xlErrDiv0 Then
    VincentyInvFwdAzimuth = CVErr(Excel.xlErrNull)
ElseIf err.number = Excel.xlErrNA Then
    VincentyInvFwdAzimuth = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

' Calculates geodesic reverse azimuth (in degrees) between two points specified by latitude/longitude (in numeric degrees) using Vincenty's inverse formula for ellipsoids
Public Function VincentyInvRevAzimuth(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Variant
Attribute VincentyInvRevAzimuth.VB_Description = "Calculates geodesic reverse azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids"
Attribute VincentyInvRevAzimuth.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo error:
    Dim p As InvParams: Set p = VincentyInv(lat1, lon1, lat2, lon2)
    Dim revAz As Double: revAz = Atan2(-p.sinU1 * p.cosU2 + p.cosU1 * p.sinU2 * p.cosLambda, p.cosU1 * p.sinLambda)
    VincentyInvRevAzimuth = ToDegrees(revAz)
    If VincentyInvRevAzimuth < 0 Then VincentyInvRevAzimuth = VincentyInvRevAzimuth + 360
    Exit Function
error:
If err.number = Excel.xlErrDiv0 Then
    VincentyInvRevAzimuth = CVErr(Excel.xlErrNull)
ElseIf err.number = Excel.xlErrNA Then
    VincentyInvRevAzimuth = CVErr(Excel.xlErrNA)
Else
    err.Raise err.number, err.Source, err.Description, err.HelpFile, err.HelpContext
End If
End Function

Private Function VincentyDir(ByVal lat As Double, ByVal lon As Double, ByVal azimuth As Double, ByVal distance As Double) As DirParams
    Dim p As DirParams: Set p = New DirParams
    
    Dim phi1 As Double: phi1 = ToRadians(lat)
    p.lambda1 = ToRadians(lon)
    Dim alpha1 As Double: alpha1 = ToRadians(azimuth)
    Dim s As Double: s = distance
    
    Dim fix1 As Double ' temp variable to prevent "formula too complex.." error
    Dim fix2 As Double ' temp variable to prevent "formula too complex.." error
    
    p.sinAlpha1 = Sin(alpha1)
    p.cosAlpha1 = cos(alpha1)

    Dim tanU1 As Double: tanU1 = (1 - f) * Tan(phi1)
    p.cosU1 = 1 / Sqr((1 + tanU1 * tanU1))
    p.sinU1 = tanU1 * p.cosU1
    Dim sigma1 As Double: sigma1 = Atan2(p.cosAlpha1, tanU1)
    p.sinAlpha = p.cosU1 * p.sinAlpha1
    p.cosSqAlpha = 1 - p.sinAlpha * p.sinAlpha
    fix1 = low_a * low_a - low_b * low_b
    Dim uSq As Double: uSq = p.cosSqAlpha * fix1 / (low_b * low_b)
    fix1 = -768 + uSq * (320 - 175 * uSq)
    Dim a As Double: a = 1 + uSq / 16384 * (4096 + uSq * fix1)
    fix1 = -128 + uSq * (74 - 47 * uSq)
    Dim B As Double: B = uSq / 1024 * (256 + uSq * fix1)
    Dim deltaSigma As Double

    p.sigma = s / (low_b * a)
    Dim sigma2 As Double
    Dim iterationCount As Integer:  iterationCount = 0
    
    Do
        p.cos2sigmaM = cos(2 * sigma1 + p.sigma)
        p.sinSigma = Sin(p.sigma)
        p.cosSigma = cos(p.sigma)
        deltaSigma = B * p.sinSigma * (p.cos2sigmaM + B / 4 * (p.cosSigma * (-1 + 2 * p.cos2sigmaM * p.cos2sigmaM) - _
            B / 6 * p.cos2sigmaM * (-3 + 4 * p.sinSigma * p.sinSigma) * (-3 + 4 * p.cos2sigmaM * p.cos2sigmaM)))
        sigma2 = p.sigma
        p.sigma = s / (low_b * a) + deltaSigma
        iterationCount = iterationCount + 1
    Loop While Abs(p.sigma - sigma2) > EPSILON And iterationCount < MaxIterations
    
    If iterationCount >= MaxIterations Then
        ' failed to converge
        err.Raise (Excel.xlErrNA): Exit Function
    End If
        
    Set VincentyDir = p
End Function

Private Function VincentyInv(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As InvParams
    Dim p As InvParams: Set p = New InvParams

    Dim l As Double: l = ToRadians(lon2 - lon1)
    Dim U1 As Double: U1 = Atn((1 - f) * Tan(ToRadians(lat1)))
    Dim U2 As Double: U2 = Atn((1 - f) * Tan(ToRadians(lat2)))
    Dim lambda As Double: lambda = l
    Dim lambdaP As Double: lambdaP = 2 * PI
    Dim sinSigma As Double
    Dim cosSigma As Double
    Dim sinAlpha As Double
    Dim cosSqAlpha As Double
    Dim cos2sigmaM As Double
    Dim c As Double
    Dim uSq As Double
    Dim upper_B As Double
    Dim fix2 As Double ' temp variable to prevent "formula too complex.." error
    Dim fix1 As Double ' temp variable to prevent "formula too complex.." error
    Dim iterationCount As Integer: iterationCount = 0
    
    p.sinU1 = Sin(U1)
    p.sinU2 = Sin(U2)
    p.cosU1 = cos(U1)
    p.cosU2 = cos(U2)
        
    Do
        p.sinLambda = Sin(lambda)
        p.cosLambda = cos(lambda)
        sinSigma = Sqr(((p.cosU2 * p.sinLambda) ^ 2) + ((p.cosU1 * p.sinU2 - p.sinU1 * p.cosU2 * p.cosLambda) ^ 2))
        
        If sinSigma = 0 Then
            err.Raise (Excel.xlErrDiv0): Exit Function
        End If
        
        cosSigma = p.sinU1 * p.sinU2 + p.cosU1 * p.cosU2 * p.cosLambda
        p.sigma = Atan2(cosSigma, sinSigma)
        sinAlpha = p.cosU1 * p.cosU2 * p.sinLambda / sinSigma
        cosSqAlpha = 1 - sinAlpha * sinAlpha

        If cosSqAlpha = 0 Then ' check for a division by zero
            cos2sigmaM = 0     ' 2 points on the equator
        Else
            cos2sigmaM = cosSigma - 2 * p.sinU1 * p.sinU2 / cosSqAlpha
        End If
        
        c = f / 16 * cosSqAlpha * (4 + f * (4 - 3 * cosSqAlpha))
        lambdaP = lambda
        
        fix1 = cos2sigmaM + c * cosSigma * (-1 + 2 * (cos2sigmaM ^ 2))
        lambda = l + (1 - c) * f * sinAlpha * (p.sigma + c * sinSigma * fix1)
        
        iterationCount = iterationCount + 1
    Loop While Abs(lambda - lambdaP) > EPSILON And iterationCount < MaxIterations
    
    If iterationCount >= MaxIterations Then
        ' failed to converge
        err.Raise (Excel.xlErrNA): Exit Function
    End If

    uSq = cosSqAlpha * (low_a ^ 2 - low_b ^ 2) / (low_b ^ 2)
    
    fix1 = -768 + uSq * (320 - 175 * uSq)
    p.upper_A = 1 + uSq / 16384 * (4096 + uSq * fix1)
    
    fix1 = -128 + uSq * (74 - 47 * uSq)
    upper_B = uSq / 1024 * (256 + uSq * fix1)
    
    fix1 = cosSigma * (-1 + 2 * cos2sigmaM ^ 2)
    fix2 = 6 * cos2sigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2sigmaM ^ 2)
    
    p.deltaSigma = upper_B * sinSigma * (cos2sigmaM + upper_B / 4 * (fix1 - upper_B / fix2))
    
    Set VincentyInv = p
End Function

' Converts decimal latitude, longitude or azimuth value to degrees/minutes/seconds string format
Public Function ConvertDegrees(ByVal decimalDeg As Double) As String
Attribute ConvertDegrees.VB_Description = "Converts latitude, longitude or azimuth in decimal degrees to string in degrees/minutes/seconds format"
Attribute ConvertDegrees.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim s As Integer: s = Sign(decimalDeg)
    Dim degrees As Double: degrees = Fix(decimalDeg)
    Dim minutes As Double: minutes = Fix((decimalDeg - degrees) * 60)
    Dim seconds As Double: seconds = Round((decimalDeg - degrees - (minutes / 60)) * 60 * 60, 4) ' 4 digit precision corresponds to ~3mm
    
    If s < 0 Then
        If degrees <> 0 Then
            minutes = minutes * s: seconds = seconds * s
        ElseIf minutes <> 0 Then
            seconds = seconds * s
        End If
    End If
    
    ConvertDegrees = Trim(Str(degrees)) & "° " & Trim(Str(minutes)) & "' " & Format(seconds, "0.####") + Chr(34)
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
    degreeDeg = Replace$(degreeDeg, "°", " ")
    degreeDeg = Replace$(degreeDeg, "ø", " ")
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
        
        If lc = "W" Or lc = "w" Or lc = "S" Or lc = "s" Then
            s = -1
        ElseIf lc = "E" Or lc = "e" Or lc = "N" Or lc = "n" Then
            ' do nothing
        Else
            ConvertDecimal = CVErr(Excel.xlErrNA): Exit Function
        End If
    ElseIf Not IsNumeric(fc) And fc <> "-" Then
        degreeDeg = Right$(degreeDeg, Len(degreeDeg) - 1) ' trim the first char
        degreeDeg = Trim$(degreeDeg)
        
        If fc = "W" Or fc = "w" Or fc = "S" Or fc = "s" Then
            s = -1
        ElseIf fc = "E" Or fc = "e" Or fc = "N" Or lc = "n" Then
            ' do nothing
        Else
            ConvertDecimal = CVErr(Excel.xlErrNA): Exit Function
        End If
    End If
    
    Dim temp As String
    
    ' remove multiple spaces
    Do
        temp = degreeDeg
        degreeDeg = Replace$(degreeDeg, Space(2), Space(1))
    Loop Until Len(temp) = Len(degreeDeg)
    
    Dim a() As String: a = Split(degreeDeg, " ")
    Dim l As Integer: l = UBound(a)
    
    Dim degrees As Double: degrees = val(a(0))
    Dim minutes As Double: If l > 0 Then minutes = val(a(1)): minutes = minutes / 60
    Dim seconds As Double: If l > 1 Then seconds = val(a(2)): seconds = seconds / 3600
    
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

Private Function ModDouble(ByVal number As Double, ByVal divisor As Double) As Double
    ModDouble = number - (divisor * Fix(number / divisor))
End Function

' source: http://en.wikibooks.org/wiki/Programming:Visual_Basic_Classic/Simple_Arithmetic#Trigonometrical_Functions
Private Function Atan2(ByVal x As Double, ByVal Y As Double) As Double
    If Y > 0 Then
        If x >= Y Then
            Atan2 = Atn(Y / x)
        ElseIf x <= -Y Then
            Atan2 = Atn(Y / x) + PI
        Else
        Atan2 = PI / 2 - Atn(x / Y)
    End If
        Else
            If x >= -Y Then
            Atan2 = Atn(Y / x)
        ElseIf x <= Y Then
            Atan2 = Atn(Y / x) - PI
        Else
            Atan2 = -Atn(x / Y) - PI / 2
        End If
    End If
End Function

Public Sub Workbook_Open()
    Application.MacroOptions Macro:="VincentyDirLat", Description:="Calculates geodesic latitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="VincentyDirLon", Description:="Calculates geodesic longitude (in degrees) based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="VincentyDirRevAzimuth", Description:="Calculates geodesic azimuth in degrees clockwise from north based on one point, azimuth and distance using Vincenty's direct formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "azimuth in degrees", "distance in meters"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="VincentyInvDistance", Description:="Calculates geodesic distance in meters between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="VincentyInvFwdAzimuth", Description:="Calculates geodesic forward azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="VincentyInvRevAzimuth", Description:="Calculates geodesic reverse azimuth in degrees clockwise from north between two points specified by latitude/longitude using Vincenty's inverse formula for ellipsoids", _
    ArgumentDescriptions:=Array("latitude 1 in degrees", "longitude 1 in degrees", "latitude 2 in degrees", "longitude 2 in degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="ConvertDecimal", Description:="Converts latitude, longitude or azimuth in degrees/minutes/seconds format to decimal value", _
    ArgumentDescriptions:=Array("string in degrees/minutes/seconds format"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
    
    Application.MacroOptions Macro:="ConvertDegrees", Description:="Converts latitude, longitude or azimuth in decimal degrees to string in degrees/minutes/seconds format", _
    ArgumentDescriptions:=Array("decimal degrees"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/VincentyExcel"
End Sub

