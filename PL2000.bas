Attribute VB_Name = "PL2000"
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Functions to translate WGS84 coordinates to/from the Polish geodetic coordinate system PL-2000
' (based on on the Gauss-Kruger coordinate system and GRS 80 ellipsoid)
' https://pl.wikipedia.org/wiki/Uk%C5%82ad_wsp%C3%B3%C5%82rz%C4%99dnych_2000
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Ported to VBA by (c) Tomasz Jastrzebski 2020 MIT Licence
' Version: 2020-07-03
' Latest version available at:
' https://github.com/tdjastrzebski/Vincenty-Excel
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Based on the Excel Spreadsheet by Edward Zadorski
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' References:
' Instrukcja "Wytyczne techniczne G-1.10"
' Roman Kadaj, "Geodeta" nr 9-12/2000
' Janusz Jaworski, "Jak przeliczac", "Geodeta" nr 4/2000
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Option Explicit

Private Const m0 As Double = 0.999923 ' skale m0
Private Const R0 As Double = 6367449.14577 ' Lagrange radius
Private Const e As Double = 0.0818191910428
Private Const PI As Double = 3.14159265358979
Private Const a2 As Double = 8.377318247344E-04
Private Const a4 As Double = 7.608527788826E-07
Private Const a6 As Double = 1.197638019173E-09
Private Const a8 As Double = 2.4433762425E-12
Private Const b2 As Double = -8.377321681641E-04
Private Const b4 As Double = -5.905869626083E-08
Private Const b6 As Double = -1.673488904988E-10
Private Const b8 As Double = -2.167737805597E-13
Private Const c2 As Double = 0.003356551485597
Private Const c4 As Double = 6.571873148459E-06
Private Const c6 As Double = 1.764656426454E-08
Private Const c8 As Double = 5.40048218776E-11

Private Type To2000
    Xmer As Double
    Ymer As Double
End Type

Private Type From2000
    w As Double
    alpha As Double
End Type

' Calculates PL-2000 X coordinate based on geodesic latitude, longitude and target meridian.
Public Function To2000X(ByVal lat As Double, ByVal lon As Double, ByVal meridian As Integer) As Double
Attribute To2000X.VB_Description = "Calculates PL-2000 X coordinate based on geodesic latitude, longitude and target meridian."
Attribute To2000X.VB_ProcData.VB_Invoke_Func = " \n20"
Dim to2k As To2000: to2k = To2000(lat, lon, meridian)
Dim Xgk As Double: Xgk = R0 * (to2k.Xmer + (a2 * Sin(2 * to2k.Xmer) * CosH(2 * to2k.Ymer)) + (a4 * Sin(4 * to2k.Xmer) * CosH(4 * to2k.Ymer)) _
    + (a6 * Sin(6 * to2k.Xmer) * CosH(6 * to2k.Ymer)) + (a8 * Sin(8 * to2k.Xmer) * CosH(8 * to2k.Ymer)))
Dim x As Double: x = m0 * Xgk
To2000X = x
End Function

' Calculates PL-2000 Y coordinate based on geodesic latitude, longitude and target meridian.
Public Function To2000Y(ByVal lat As Double, ByVal lon As Double, ByVal meridian As Integer) As Double
Attribute To2000Y.VB_Description = "Calculates PL-2000 Y coordinate based on geodesic latitude, longitude and target meridian."
Attribute To2000Y.VB_ProcData.VB_Invoke_Func = " \n20"
Dim to2k As To2000: to2k = To2000(lat, lon, meridian)
Dim Ygk As Double: Ygk = R0 * (to2k.Ymer + (a2 * Cos(2 * to2k.Xmer) * SinH(2 * to2k.Ymer)) + (a4 * Cos(4 * to2k.Xmer) * SinH(4 * to2k.Ymer)) _
    + (a6 * Cos(6 * to2k.Xmer) * SinH(6 * to2k.Ymer)) + (a8 * Cos(8 * to2k.Xmer) * SinH(8 * to2k.Ymer)))
Dim y As Double: y = m0 * Ygk
Dim M As Integer: M = (meridian / 3)
y = y + M * 1000000 + 500000
To2000Y = y
End Function

Private Function To2000(ByVal lat As Double, ByVal lon As Double, ByVal meridian As Integer) As To2000
ValidateMeridian (meridian)
Dim radLatB As Double: radLatB = ToRadians(lat)
Dim radLonL As Double: radLonL = ToRadians(lon)
Dim L0 As Double: L0 = ToRadians(meridian)
Dim fi As Double: fi = 2 * (Atn((((1 - e * Sin(radLatB)) / (1 + e * Sin(radLatB))) ^ (e / 2)) * Tan(radLatB / 2 + PI / 4)) - PI / 4)
Dim to2k As To2000
to2k.Xmer = Atn(Sin(fi) / (Cos(fi) * Cos(radLonL - L0)))
to2k.Ymer = 0.5 * Log((1 + Cos(fi) * Sin(radLonL - L0)) / (1 - Cos(fi) * Sin(radLonL - L0)))
To2000 = to2k
End Function

' Calculates geodesic latitude (in degrees) based on PL-2000 X, Y coordinates and meridian.
Public Function From2000Lat(ByVal x As Double, ByVal y As Double, ByVal meridian As Integer) As Double
Attribute From2000Lat.VB_Description = "Calculates geodesic latitude (in degrees) based on PL-2000 X, Y coordinates and meridian."
Attribute From2000Lat.VB_ProcData.VB_Invoke_Func = " \n20"
Dim from2k As From2000: from2k = From2000(x, y, meridian)
Dim fi As Double: fi = ASin(Cos(from2k.w) * Sin(from2k.alpha))
Dim radB As Double: radB = fi + c2 * Sin(2 * fi) + c4 * Sin(4 * fi) + c6 * Sin(6 * fi) + c8 * Sin(8 * fi) + 0.0000000008
Dim B As Double: B = ToDegrees(radB)
From2000Lat = B
End Function

' Calculates geodesic longitude (in degrees) based on PL-2000 X, Y coordinates and meridian.
Public Function From2000Lon(ByVal x As Double, ByVal y As Double, ByVal meridian As Integer) As Double
Attribute From2000Lon.VB_Description = "Calculates geodesic longitude (in degrees) based on PL-2000 X, Y coordinates and meridian."
Attribute From2000Lon.VB_ProcData.VB_Invoke_Func = " \n20"
Dim from2k As From2000: from2k = From2000(x, y, meridian)
Dim dl As Double: dl = Atn((Tan(from2k.w)) / Cos(from2k.alpha))
Dim L As Double: L = meridian + ToDegrees(dl)
From2000Lon = L
End Function

Private Function From2000(ByVal x As Double, ByVal y As Double, ByVal meridian As Integer) As From2000
ValidateMeridian (meridian)
Dim Xgk As Double: Xgk = x / m0
Dim M As Integer: M = (meridian / 3)
Dim Ygk As Double: Ygk = (y - M * 1000000 - 500000) / m0
Dim u As Double: u = Xgk / R0
Dim v As Double: v = Ygk / R0
Dim from2k As From2000
from2k.alpha = u + (b2 * Sin(2 * u) * CosH(2 * v) + b4 * Sin(4 * u) * CosH(4 * v) + b6 * Sin(6 * u) * CosH(6 * v) + b8 * Sin(8 * u) * CosH(8 * v))
Dim beta As Double: beta = v + (b2 * Cos(2 * u) * SinH(2 * v) + b4 * Cos(4 * u) * SinH(4 * v) + b6 * Cos(6 * u) * SinH(6 * v) + b8 * Cos(8 * u) * SinH(8 * v))
from2k.w = 2 * Atn(Exp(beta)) - PI / 2
From2000 = from2k
End Function

Private Sub ValidateMeridian(ByVal meridian As Integer)
Select Case meridian
Case 15, 18, 21, 24
    ' do nothing
Case Else
    Err.Raise (Excel.xlErrNA)
End Select
End Sub

Private Function ToRadians(ByVal degrees As Double) As Double
ToRadians = degrees * (PI / 180)
End Function

Private Function ToDegrees(ByVal radians As Double) As Double
ToDegrees = (radians * 180) / PI
End Function

Private Function SinH(ByVal radians As Double) As Double
SinH = (Exp(radians) - Exp(-radians)) / 2
End Function

Private Function CosH(ByVal radians As Double) As Double
CosH = (Exp(radians) + Exp(-radians)) / 2
End Function

Private Function ASin(x As Double) As Double
If Abs(x) = 1 Then
    ASin = Sgn(x) * PI / 2
Else
    ASin = Atn(x / Sqr(1 - x ^ 2))
End If
End Function

Private Sub Workbook_Open()
    ' note: Description is 255 chars max
    Application.MacroOptions Macro:="From2000Lat", Description:="Calculates geodesic latitude (in degrees) based on PL-2000 X, Y coordinates and meridian.", _
    ArgumentDescriptions:=Array("X coordinate", "Y coordinate", "meridian - accepted values are 15, 18, 21, 24"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="From2000Lon", Description:="Calculates geodesic longitude (in degrees) based on PL-2000 X, Y coordinates and meridian.", _
    ArgumentDescriptions:=Array("X coordinate", "Y coordinate", "meridian - accepted values are 15, 18, 21, 24"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="To2000X", Description:="Calculates PL-2000 X coordinate based on geodesic latitude, longitude and target meridian.", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "target meridian - accepted values are 15, 18, 21, 24"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
    
    Application.MacroOptions Macro:="To2000Y", Description:="Calculates PL-2000 Y coordinate based on geodesic latitude, longitude and target meridian.", _
    ArgumentDescriptions:=Array("latitude in degrees", "longitude in degrees", "target meridian - accepted values are 15, 18, 21, 24"), Category:="Geodesic", HelpFile:="https://github.com/tdjastrzebski/Vincenty-Excel"
End Sub

