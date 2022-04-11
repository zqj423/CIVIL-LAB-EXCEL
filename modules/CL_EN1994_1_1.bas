Attribute VB_Name = "CL_EN1994_1_1"
'#############################################################################################################
' Eurocode 4 function Module
'
'This Module contains functions from or related to EN1994-1-1 for composite structures
'
'
'The function name should be all started with 'CL_EC4_'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC

'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first


'################################################ SECTION 3 ##################################################



'################################################ SECTION 4 ##################################################

'################################################ SECTION 5 ##################################################

' EFFECTIVE WIDTH   REF: EN1994-1-1 S5.4.1.2 PAGE 29 ---------------------------------------------------------
Function CL_EC4_beff(b_0 As Variant, b_1 As Variant, b_2 As Variant, L_e As Variant, is_Middle_Span As Boolean) As Double
    
    If IsMissing(is_Middle_Span) Then
        is_Middle_Span = True
    End If
    
    If is_Middle_Span Then
        
       CL_EC4_beff = b_0 + WorksheetFunction.Min(L_e / 8, b_1) + WorksheetFunction.Min(L_e / 8, b_2)
        
    Else
        Dim beta_1  As Double
        Dim beta_2  As Double
        Dim bc_1    As Double
        Dim bc_2    As Double
        
        bc_1 = WorksheetFunction.Min(L_e / 8, b_1)
        bc_2 = WorksheetFunction.Min(L_e / 8, b_2)
        
        beta_1 = WorksheetFunction.Min(1, 0.55 + 0.025 * L_e / bc_1)
        beta_2 = WorksheetFunction.Min(1, 0.55 + 0.025 * L_e / bc_2)
        
        CL_EC4_beff = b_0 + bc_1 * beta_1 + bc_2 * beta_2
        
    End If
    
End Function

' EFFECTIVE lENGTH   REF: EN1994-1-1 S5.4.2.1 PAGE 31 ---------------------------------------------------------
Function CL_EC4_Le(location As Integer, L1 As Variant, L2 As Variant) As Double
    
    If location = 1 Then
        CL_EC4_Le = 0.85 * L1
    ElseIf location = 2 Then
        CL_EC4_Le = 0.25 * (L1 + L2)
    ElseIf location = 3 Then
        CL_EC4_Le = 0.7 * L2
    ElseIf location = 4 Then
        CL_EC4_Le = 2 * L2
    ElseIf location = 5 Then
        CL_EC4_Le = L1
        
    End If
    
End Function




'################################################ SECTION 6 ##################################################
Function CL_EC4_EIeff(ParamArray list1()) As Double
    ' function to calculate the effective stiffness of cross-section contains many parts.
    ' the inputs should follow "E1,I1,A1,z1,E2,I2,A2,z2 ... Ei,Ii,Ai,zi"
    
    Dim n           As Integer        'the number of shapes
    Dim nl          As Integer        ' the number of arguments in the list
    Dim i           As Integer        ' midvalues
    Dim z_na        As Double        'the location of neutral axis
    Dim EA_tot      As Double        ' the total area of the shape
    Dim EA_totz     As Double        ' the total area multiply by z
    Dim info()      As Double        'the new array for the information (1 to 4, 1 to n)
    Dim EIeff       As Double        ' the stiffness effective
    Dim list      As Variant
    
    Dim a_tot, a_totz As Double
    
    Dim list1_()    As Variant
    
    'resort the data
    
    list1_ = list1
    
    list = Reg2List(list1_)
    
    nl = UBound(list)
    
    ' checking the number of arguments
    If Not (nl Mod 4 = 0) Then
        
        MsgBox "wrong numbers of arguments, check again"
        Exit Function
    End If
    
    n = nl / 4
    
    ReDim info(1 To 4, 1 To n)
    ' read date to info
    ': info(1:i) : Elastic modulus;
    ': info(2:i) : moment of inertia;
    ': info(3:i) : area;
    ': info(4:i) : distance of gravity center to reference point;
    
    a_tot = 0
    a_totz = 0
    
    'calculate the neutral axis
    For i = 1 To n
        info(1, i) = list(4 * (i - 1) + 1)
        info(2, i) = list(4 * (i - 1) + 2)
        info(3, i) = list(4 * (i - 1) + 3)
        info(4, i) = list(4 * (i - 1) + 4)
        
        EA_tot = EA_tot + info(3, i) * info(1, i)
        EA_totz = EA_totz + info(3, i) * info(4, i) * info(1, i)
        
    Next i
    
    z_na = EA_totz / EA_tot
    
    'calcualte the effective stiffness
    EIeff = 0
    
    For i = 1 To n
        
        EIeff = EIeff + info(1, i) * (info(2, i) + info(3, i) * (info(4, i) - z_na) ^ 2)
        
    Next i
    
    CL_EC4_EIeff = EIeff
    
End Function

' function to calculate the neutral axisof cross-section contains many parts.
' the inputs should follow "E1,I1,A1,z1,E2,I2,A2,z2 ... Ei,Ii,Ai,zi"

Function CL_EC4_Zel(ParamArray list1()) As Double
    
    Dim n           As Integer        'the number of shapes
    Dim nl          As Integer        ' the number of arguments in the list
    Dim i           As Integer        ' midvalues
    Dim z_na        As Double        'the location of neutral axis
    Dim EA_tot      As Double        ' the total area of the shape
    Dim EA_totz     As Double        ' the total area multiply by z
    Dim info()      As Double        'the new array for the information (1 to 4, 1 to n)
    Dim EIeff       As Double        ' the stiffness effective
    Dim list()      As Variant
    
    Dim a_tot, a_totz As Double
    
    Dim list1_()    As Variant
    
    'resort the data
    
    list1_ = list1
    
    list = Reg2List(list1_)
    
    nl = UBound(list)
    
    ' checking the number of arguments
    If Not (nl Mod 4 = 0) Then
        
        MsgBox "wrong numbers of arguments, check again"
        Exit Function
    End If
    
    n = nl / 4
    
    ReDim info(1 To 4, 1 To n)
    ' read date to info
    ': info(1:i) : Elastic modulus;
    ': info(2:i) : moment of inertia;
    ': info(3:i) : area;
    ': info(4:i) : distance of gravity center to reference point;
    
    a_tot = 0
    a_totz = 0
    
    'calculate the neutral axis
    For i = 1 To n
        info(1, i) = list(4 * (i - 1) + 1)
        info(2, i) = list(4 * (i - 1) + 2)
        info(3, i) = list(4 * (i - 1) + 3)
        info(4, i) = list(4 * (i - 1) + 4)
        
        EA_tot = EA_tot + info(3, i) * info(1, i)
        EA_totz = EA_totz + info(3, i) * info(4, i) * info(1, i)
        
    Next i
    
    z_na = EA_totz / EA_tot
    
    'calcualte the effective stiffness
    EIeff = 0
    
    For i = 1 To n
        
        EIeff = EIeff + info(1, i) * (info(2, i) + info(3, i) * (info(4, i) - z_na) ^ 2)
        
    Next i
    
    CL_EC4_Zel = z_na
    
End Function


'################################################ SECTION 7 ##################################################

'################################################ SECTION 8 ##################################################


'################################################ SECTION 9 ##################################################
