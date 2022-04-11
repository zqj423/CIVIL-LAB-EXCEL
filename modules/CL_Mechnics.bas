Attribute VB_Name = "CL_Mechnics"
'#############################################################################################################
' Mechanics Function Module
'
'This Module contains functions of general mechanics for the excel
'Equations related to Geometry structral analysis ect are included.
'
'The function name should be all started with 'CL_M_'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC

'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit        ' all Parameters must be defined first

'############################################ Cross-section Properties ###############################################################

Function CL_M_RectData(h As Double, B As Double, info As String) As Variant
    
    ' check and return rectangle geometry informations'
    Dim CN          As Integer
    Dim res         As Double
    
    Select Case info
        
        Case "A"
            res = h * B
            
        Case "Iy"
            res = 1 / 12 * B * h ^ 3
            
        Case "Iz"
            res = 1 / 12 * h * B ^ 3
            
        Case "Wply"
            res = 1 / 4 * B * h ^ 2
            
        Case "Wplz"
            res = 1 / 4 * h * B ^ 2
            
        Case "Wely"
            res = 1 / 6 * B * h ^ 2
            
        Case "Welz"
            res = 1 / 6 * h * B ^ 2
            
        Case "U"
            res = 2 * (h + B)
            
        Case Else
            
            MsgBox "error -info Type please use A, Iy,Iz, Wply, Wplz, Wely, Welz, U"
            
    End Select
    
    CL_M_RectData = res
    
End Function



Function CL_M_Zel(ParamArray list1()) As Double
    
    Dim n           As Integer        'the number of shapes
    Dim nl          As Integer        ' the number of arguments in the list
    Dim i           As Integer        ' midvalues
    Dim z_na        As Double        'the location of neutral axis
    Dim EA_tot      As Double        ' the total area of the shape
    Dim EA_totz     As Double        ' the total area multiply by z
    Dim info()      As Double        'the new array for the information (1 to 4, 1 to n)
    Dim list      As Variant
    
    Dim a_tot, a_totz As Double
    
    Dim list1_()    As Variant
    
    'resort the data
    
    list1_ = list1
    
    list = Reg2List(list1_)
    
    nl = UBound(list)
    
    ' checking the number of arguments
    If Not (nl Mod 2 = 0) Then
        
        MsgBox "wrong numbers of arguments, check again"
        Exit Function
    End If
    
    n = nl / 2
    
    ReDim info(1 To 2, 1 To n)
    ' read date to info
    ': info(1:i) : A;
    ': info(2:i) : z;

    
    a_tot = 0
    a_totz = 0
    
    'calculate the neutral axis
    For i = 1 To n
        info(1, i) = list(2 * (i - 1) + 1)
        info(2, i) = list(2 * (i - 1) + 2)
        
        EA_tot = EA_tot + info(1, i)
        EA_totz = EA_totz + info(1, i) * info(2, i)
        
    Next i
    
    z_na = EA_totz / EA_tot
    
    'calcualte the effective stiffness

    
    CL_M_Zel = z_na
    
End Function


'***********************************************************************************
'
'
'checking data of self defined double symmetrical I beam
'

'
'***********************************************************************************


Function CL_M_HProfileData(h As Double, B As Double, tf As Double, tw As Double, info As String) As Double

Dim res As Double

Select Case info

Case "Iy"

res = 1 / 12 * tw * (h - tf * 2) ^ 3 + 2 * (1 / 12 * B * tf ^ 3 + B * tf * (h / 2 - tf / 2) ^ 2)

Case "Iz"

res = 1 / 12 * tf * B ^ 3 * 2 + 1 / 12 * (h - 2 * tf) * tw ^ 3

Case "A"

res = h * B - (h - tf * 2) * (B - tw)

Case "Wely"
res = 1 / 12 * tw * (h - tf * 2) ^ 3 + 2 * (1 / 12 * B * tf ^ 3 + B * tf * (h / 2 - tf / 2) ^ 2)
res = res / h * 2

Case "Welz"

res = 1 / 12 * tf * B ^ 3 * 2 + 1 / 12 * (h - 2 * tf) * tw ^ 3
res = res / B * 2

Case "Wply"

res = 1 / 4 * B * h ^ 2 - 1 / 4 * (B - tw) * (h - tf * 2) ^ 2

Case "Wplz"

res = 1 / 4 * tf * 2 * B ^ 2 + 1 / 4 * (h - 2 * tf) * tw ^ 2

Case Else
res = 0
MsgBox "parameter wrong, plase use any of Iy Iz A Wely Welz Wply Wplz"

End Select

CL_M_HProfileData = res

End Function


Function CL_M_TwoWaySlabAlpha(lratio As Double, Optional ResultType As Integer = 0, Optional supporttype As Integer = 0) As Double
Dim tablename As String
Dim lxlys As String
Dim axs As String
Dim ays As String
Dim ax0s As String
Dim ay0s As String

If supporttype = 0 Then
lxlys = "CLT_TwoWaySlabA0[lx/ly]"
axs = "CLT_TwoWaySlabA0[ax]"
ays = "CLT_TwoWaySlabA0[ay]"
ax0s = "CLT_TwoWaySlabA0[ax0]"
ay0s = "CLT_TwoWaySlabA0[ay0]"


ElseIf supporttype = 1 Then
lxlys = "CLT_TwoWaySlabA1[lx/ly]"
axs = "CLT_TwoWaySlabA1[ax]"
ays = "CLT_TwoWaySlabA1[ay]"
ax0s = "CLT_TwoWaySlabA1[ax0]"
ay0s = "CLT_TwoWaySlabA1[ay0]"

ElseIf supporttype = 2 Then
lxlys = "CLT_TwoWaySlabA2[lx/ly]"
axs = "CLT_TwoWaySlabA2[ax]"
ays = "CLT_TwoWaySlabA2[ay]"
ax0s = "CLT_TwoWaySlabA2[ax0]"
ay0s = "CLT_TwoWaySlabA2[ay0]"

ElseIf supporttype = 3 Then
lxlys = "CLT_TwoWaySlabA3[lx/ly]"
axs = "CLT_TwoWaySlabA3[ax]"
ays = "CLT_TwoWaySlabA3[ay]"
ax0s = "CLT_TwoWaySlabA3[ax0]"
ay0s = "CLT_TwoWaySlabA3[ay0]"

ElseIf supporttype = 4 Then
lxlys = "CLT_TwoWaySlabA4[lx/ly]"
axs = "CLT_TwoWaySlabA4[ax]"
ays = "CLT_TwoWaySlabA4[ay]"
ax0s = "CLT_TwoWaySlabA4[ax0]"
ay0s = "CLT_TwoWaySlabA4[ay0]"

ElseIf supporttype = 5 Then
lxlys = "CLT_TwoWaySlabA5[lx/ly]"
axs = "CLT_TwoWaySlabA5[ax]"
ays = "CLT_TwoWaySlabA5[ay]"
ax0s = "CLT_TwoWaySlabA5[ax0]"
ay0s = "CLT_TwoWaySlabA5[ay0]"

ElseIf supporttype = 6 Then
lxlys = "CLT_TwoWaySlabA6[lx/ly]"
axs = "CLT_TwoWaySlabA6[ax]"
ays = "CLT_TwoWaySlabA6[ay]"
ax0s = "CLT_TwoWaySlabA6[ax0]"
ay0s = "CLT_TwoWaySlabA6[ay0]"

ElseIf supporttype = 7 Then
lxlys = "CLT_TwoWaySlabA7[lx/ly]"
axs = "CLT_TwoWaySlabA7[ax]"
ays = "CLT_TwoWaySlabA7[ay]"
ax0s = "CLT_TwoWaySlabA7[ax0]"
ay0s = "CLT_TwoWaySlabA7[ay0]"

ElseIf supporttype = 9 Then
lxlys = "CLT_TwoWaySlabA9[lx/ly]"
axs = "CLT_TwoWaySlabA9[ax]"
ays = "CLT_TwoWaySlabA9[ay]"
ax0s = "CLT_TwoWaySlabA9[ax0]"
ay0s = "CLT_TwoWaySlabA9[ay0]"

ElseIf supporttype = 8 Then
lxlys = "CLT_TwoWaySlabA8[lx/ly]"
axs = "CLT_TwoWaySlabA8[ax]"
ays = "CLT_TwoWaySlabA8[ay]"
ax0s = "CLT_TwoWaySlabA8[ax0]"
ay0s = "CLT_TwoWaySlabA8[ay0]"


End If
If ResultType = 0 Then
CL_G_TwoWaySlabAlpha = CL_G_TabelInterpolation(Range(lxlys), Range(axs), lratio)
ElseIf ResultType = 1 Then
CL_G_TwoWaySlabAlpha = CL_G_TabelInterpolation(Range(lxlys), Range(ays), lratio)
ElseIf ResultType = 2 Then
CL_G_TwoWaySlabAlpha = CL_G_TabelInterpolation(Range(lxlys), Range(ax0s), lratio)
ElseIf ResultType = 3 Then
CL_G_TwoWaySlabAlpha = CL_G_TabelInterpolation(Range(lxlys), Range(ay0s), lratio)
End If
End Function

Function CL_M_Rebar_A(d As Double, n As Double) As Double
    ' find the area of rebar
    CL_G_Rebar_A = d * d / 4 * 3.14159267 * n
    
End Function

Function CL_M_Rebars_A(ParamArray list1()) As Double
    
    Dim n           As Integer        'the number of shapes
    Dim nl          As Integer        ' the number of arguments in the list
    Dim i           As Integer        ' midvalues
    Dim z_na        As Double        'the location of neutral axis

    Dim info()      As Double        'the new array for the information (1 to 4, 1 to n)
    Dim list      As Variant
    
    Dim a_tot As Double
    
    Dim list1_()    As Variant
    
    'resort the data
    
    list1_ = list1
    
    list = Reg2List(list1_)
    
    nl = UBound(list)
    
    ' checking the number of arguments
    If Not (nl Mod 2 = 0) Then
        
        MsgBox "wrong numbers of arguments, check again"
        Exit Function
    End If
    
    n = nl / 2
    
    ReDim info(1 To 2, 1 To n)
    ' read date to info
    ': info(1:i) : d;
    ': info(2:i) : n;

    
    a_tot = 0

    
    'calculate the neutral axis
    For i = 1 To n
        info(1, i) = list(2 * (i - 1) + 1)
        info(2, i) = list(2 * (i - 1) + 2)
    
        a_tot = a_tot + info(1, i) * info(1, i) / 4 * 3.14159267 * info(2, i)
        
    Next i
    

    CL_M_Rebars_A = a_tot
    
End Function
