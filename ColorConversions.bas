Attribute VB_Name = "ColorConversions"
'Formulas from http://www.easyrgb.com/en/math.php

'-------------------------------------
'|    Conversion LAB <--> XYZ        |
'-------------------------------------

Sub LAB2XYZ(ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    
    var_Y = (LabL + 16) / 116
    var_X = LabA / 500 + var_Y
    var_Z = var_Y - LabB / 200
    
    If (var_Y ^ 3 > 0.008856) Then var_Y = var_Y ^ 3 _
                              Else var_Y = (var_Y - 16 / 116) / 7.787
    If (var_X ^ 3 > 0.008856) Then var_X = var_X ^ 3 _
                              Else var_X = (var_X - 16 / 116) / 7.787
    If (var_Z ^ 3 > 0.008856) Then var_Z = var_Z ^ 3 _
                              Else var_Z = (var_Z - 16 / 116) / 7.787

    
    'Ref D50
    x = var_X * 96.422 ' Ref_X D50
    y = var_Y * 100#   ' Ref_Y D50
    z = var_Z * 82.521 ' Ref_Z D50
End Sub

Sub XYZ2LAB(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)
    'Reference-X, Y and Z refer to specific illuminants and observers.
    'Common reference values are available below in this same page.
    
    'Ref D50
    var_X = x / 96.422
    var_Y = y / 100#
    var_Z = z / 82.521
    
    If (var_X > 0.008856) Then var_X = var_X ^ (1 / 3) _
                          Else var_X = (7.787 * var_X) + (16 / 116)
    If (var_Y > 0.008856) Then var_Y = var_Y ^ (1 / 3) _
                          Else var_Y = (7.787 * var_Y) + (16 / 116)
    If (var_Z > 0.008856) Then var_Z = var_Z ^ (1 / 3) _
                          Else var_Z = (7.787 * var_Z) + (16 / 116)
    
    LabL = (116 * var_Y) - 16
    LabA = 500 * (var_X - var_Y)
    LabB = 200 * (var_Y - var_Z)
End Sub





'-------------------------------------
'|    Conversion XYZ <--> sRGB       |
'-------------------------------------

Sub XYZ2sRGB(ByRef x As Double, ByRef y As Double, ByRef z As Double, ByRef sR As Double, ByRef sG As Double, ByRef sB As Double)
    'X, Y and Z input refer to a D65/2° standard illuminant.
    'sR, sG and sB (standard RGB) output range = 0 ÷ 255
    
    var_X = x / 100
    var_Y = y / 100
    var_Z = z / 100
    
    
    'Bradford XYZ D50 to sRGB D65
    var_R = var_X * 3.1338561 + var_Y * (-1.6168667) + var_Z * (-0.4906146)
    var_G = var_X * (-0.9787684) + var_Y * 1.9161415 + var_Z * 0.033454
    var_B = var_X * 0.0719453 + var_Y * (-0.2289914) + var_Z * 1.4052427
    
    'Companding
    If (var_R < 0) Then var_R = 0 Else If (var_R > 0.0031308) Then var_R = 1.055 * (var_R ^ (1 / 2.4)) - 0.055 _
                                       Else var_R = 12.92 * var_R
    If (var_G < 0) Then var_G = 0 Else If (var_G > 0.0031308) Then var_G = 1.055 * (var_G ^ (1 / 2.4)) - 0.055 _
                                       Else var_G = 12.92 * var_G
    If (var_B < 0) Then var_B = 0 Else If (var_B > 0.0031308) Then var_B = 1.055 * (var_B ^ (1 / 2.4)) - 0.055 _
                                       Else var_B = 12.92 * var_B

    sR = Min(var_R * 255, 255)
    sG = Min(var_G * 255, 255)
    sB = Min(var_B * 255, 255)
End Sub

Sub sRGB2XYZ(ByVal sR As Double, ByVal sG As Double, ByVal sB As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    'sR, sG and sB (Standard RGB) input range = 0 ÷ 255
    'X, Y and Z output refer to a D50/2° standard illuminant.
    
    var_R = (sR / 255#)
    var_G = (sG / 255#)
    var_B = (sB / 255#)
    
    If (var_R > 0.04045) Then var_R = ((var_R + 0.055) / 1.055) ^ 2.4 _
                         Else var_R = var_R / 12.92
    If (var_G > 0.04045) Then var_G = ((var_G + 0.055) / 1.055) ^ 2.4 _
                         Else var_G = var_G / 12.92
    If (var_B > 0.04045) Then var_B = ((var_B + 0.055) / 1.055) ^ 2.4 _
                         Else var_B = var_B / 12.92
    
    
    Dim M(1 To 3, 1 To 3) As Double
    
    'Bradford sRGB D65 to XYZ D50
    M(1, 1) = 0.4360747
    M(1, 2) = 0.3850649
    M(1, 3) = 0.1430804
    
    M(2, 1) = 0.2225045
    M(2, 2) = 0.7168786
    M(2, 3) = 0.0606169
    
    M(3, 1) = 0.0139322
    M(3, 2) = 0.0971045
    M(3, 3) = 0.7141733
    
    
    var_X = var_R * M(1, 1) + var_G * M(1, 2) + var_B * M(1, 3)
    var_Y = var_R * M(2, 1) + var_G * M(2, 2) + var_B * M(2, 3)
    var_Z = var_R * M(3, 1) + var_G * M(3, 2) + var_B * M(3, 3)
    
    
    'XYZ in [0, 100]
    x = var_X * 100#
    y = var_Y * 100#
    z = var_Z * 100#
End Sub






'-------------------------------------
'|    Conversion XYZ <--> aRGB       |
'-------------------------------------



Sub XYZ2aRGB(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef aR As Double, ByRef aG As Double, ByRef aB As Double)
    'X, Y and Z input refer to a D50/2° standard illuminant.
    'aR, aG and aB (RGB Adobe 1998) output range = 0 ÷ 255
    var_X = x / 100#
    var_Y = y / 100#
    var_Z = z / 100#
    
    Dim M(1 To 3, 1 To 3) As Double
    
    'Bradford XYZ D50 to aRGB D65
    M(1, 1) = 1.9624274
    M(1, 2) = -0.6105343
    M(1, 3) = -0.3413404
    
    M(2, 1) = -0.9787684
    M(2, 2) = 1.9161415
    M(2, 3) = 0.033454
    
    M(3, 1) = 0.0286869
    M(3, 2) = -0.1406752
    M(3, 3) = 1.3487655
    
    var_R = var_X * M(1, 1) + var_Y * M(1, 2) + var_Z * M(1, 3)
    var_G = var_X * M(2, 1) + var_Y * M(2, 2) + var_Z * M(2, 3)
    var_B = var_X * M(3, 1) + var_Y * M(3, 2) + var_Z * M(3, 3)

    If var_R < 0 Then var_R = 0 Else var_R = var_R ^ (1 / 2.19921875)
    If var_G < 0 Then var_G = 0 Else var_G = var_G ^ (1 / 2.19921875)
    If var_B < 0 Then var_B = 0 Else var_B = var_B ^ (1 / 2.19921875)
    
    aR = var_R * 255#
    aG = var_G * 255#
    aB = var_B * 255#
End Sub

Sub aRGB2XYZ(ByVal aR As Double, ByVal aG As Double, ByVal aB As Double, ByRef x As Double, ByRef y As Double, ByRef z As Double)
    'aR, aG and aB (RGB Adobe 1998) input range = 0 ÷ 255
    'X, Y and Z output refer to a D50/2° standard illuminant.
    
    var_R = (aR / 255#)
    var_G = (aG / 255#)
    var_B = (aB / 255#)
    
    var_R = var_R ^ 2.19921875
    var_G = var_G ^ 2.19921875
    var_B = var_B ^ 2.19921875
    
    var_R = var_R * 100#
    var_G = var_G * 100#
    var_B = var_B * 100#
    
    Dim M(1 To 3, 1 To 3) As Double
    'Bradford XYZ D50 to aRGB D65
    M(1, 1) = 0.6097559
    M(1, 2) = 0.2052401
    M(1, 3) = 0.149224
    
    M(2, 1) = 0.3111242
    M(2, 2) = 0.625656
    M(2, 3) = 0.0632197
    
    M(3, 1) = 0.0194811
    M(3, 2) = 0.0608902
    M(3, 3) = 0.7448387
    
    
    x = var_R * M(1, 1) + var_G * M(1, 2) + var_B * M(1, 3)
    y = var_R * M(2, 1) + var_G * M(2, 2) + var_B * M(2, 3)
    z = var_R * M(3, 1) + var_G * M(3, 2) + var_B * M(3, 3)
End Sub









'Wrapper functions
Public Sub LAB2sRGB(ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    LAB2XYZ LabL, LabA, LabB, x, y, z
    XYZ2sRGB x, y, z, r, g, b

End Sub
Public Sub sRGB2LAB(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    sRGB2XYZ r, g, b, x, y, z
    XYZ2LAB x, y, z, LabL, LabA, LabB

End Sub
Public Sub LAB2aRGB(ByVal LabL As Double, ByVal LabA As Double, ByVal LabB As Double, ByRef r As Double, ByRef g As Double, ByRef b As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    LAB2XYZ LabL, LabA, LabB, x, y, z
    XYZ2aRGB x, y, z, r, g, b

End Sub
Public Sub aRGB2LAB(ByVal r As Double, ByVal g As Double, ByVal b As Double, ByRef LabL As Double, ByRef LabA As Double, ByRef LabB As Double)

    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    aRGB2XYZ r, g, b, x, y, z
    XYZ2LAB x, y, z, LabL, LabA, LabB
End Sub






Sub CalcLAB()
Attribute CalcLAB.VB_ProcData.VB_Invoke_Func = "a\n14"
Dim r As Double
Dim g As Double
Dim b As Double
Dim LabL As Double
Dim LabA As Double
Dim LabB As Double

    r = Selection.Cells(1, 1).Value
    g = Selection.Cells(1, 2).Value
    b = Selection.Cells(1, 3).Value
    
    sRGB2LAB r, g, b, LabL, LabA, LabB
    Selection.Cells(1, 3).offset(0, 1).Cells.Value = LabL
    Selection.Cells(1, 3).offset(0, 2).Cells.Value = LabA
    Selection.Cells(1, 3).offset(0, 3).Cells.Value = LabB
    
End Sub

Sub CalcRGB()
Attribute CalcRGB.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r As Double
Dim g As Double
Dim b As Double
Dim LabL As Double
Dim LabA As Double
Dim LabB As Double

    LabL = Selection.Cells(1, 1).Value
    LabA = Selection.Cells(1, 2).Value
    LabB = Selection.Cells(1, 3).Value
    
    LAB2aRGB LabL, LabA, LabB, r, g, b
    Selection.Cells(1, 3).offset(0, 1).Cells.Value = r
    Selection.Cells(1, 3).offset(0, 2).Cells.Value = g
    Selection.Cells(1, 3).offset(0, 3).Cells.Value = b
    
End Sub

Private Function Min(ParamArray values() As Variant) As Variant
   Dim minValue, Value As Variant
   minValue = values(0)
   For Each Value In values
       If Value < minValue Then minValue = Value
   Next
   Min = minValue
End Function

Private Function Max(ParamArray values() As Variant) As Variant
   Dim maxValue, Value As Variant
   maxValue = values(0)
   For Each Value In values
       If Value > maxValue Then maxValue = Value
   Next
   Max = maxValue
End Function
