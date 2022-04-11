Attribute VB_Name = "CL_EN1992_1_1"
'#############################################################################################################
' Eurocode 2 function Module
'
'This Module contains functions from or related to EN1992 for concrete structures
'
'The function name should be all started with 'CL_EC2_'
'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC

'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first

'################################################ SECTION 3 ##################################################



'################################################ SECTION 4 ##################################################

'################################################ SECTION 5 ##################################################

'################################################ SECTION 6 ##################################################

'################################################ SECTION 7 ##################################################

'################################################ SECTION 8 ##################################################


'################################################ SECTION 9 ##################################################


'################################################ Annex ##################################################


' calcualte the creep coefficent in Eurocode2 based on Annex B---------------------------------------------------------
Function CL_EC2_CreepCo(t As Double, t_0T As Variant, RH As Variant, Ac As Variant, u As Variant, f_cm As Variant, cement_type As Variant)

' calcualte the creep coefficent in Eurocode2 based on Annex B
' cement_type as -1 for class S 0 for class N 1 for class R
' RH should be in 0 to 100
' fcm Mpa Ac mm2 u mm

Dim alpha_1 As Variant
Dim alpha_2 As Variant
Dim alpha_3 As Variant
Dim beta_fcm As Variant
Dim h_0 As Variant
Dim phi_RH As Variant
Dim beta_t0 As Variant
Dim phi_0 As Variant
Dim beta_c_tt0 As Variant
Dim beta_h As Variant
Dim t_0 As Variant

If t = 0 Then
t = 1E+22
End If

'check cerment
Select Case cement_type

Case "S"
cement_type = -1
Case "N"
cement_type = 0
Case "R"
cement_type = 1

End Select

'calcualtion of alpha values B.8C

alpha_1 = (35 / f_cm) ^ 0.7
alpha_2 = (35 / f_cm) ^ 0.2
alpha_3 = (35 / f_cm) ^ 0.5

'calcualtion of t_0
t_0 = Application.WorksheetFunction.Max(t_0T * (9 / (2 + t_0T ^ 1.2) + 1) ^ cement_type, 0.5)
'calcualtion of beta fcm B.4
beta_fcm = 16.8 / (f_cm ^ 0.5)

'calcualtion of phi_RH
h_0 = 2 * Ac / u

    If f_cm > 35 Then
     phi_RH = (1 + (1 - RH / 100) / (0.1 * h_0 ^ (1 / 3)) * alpha_1) * alpha_2
     beta_h = Application.WorksheetFunction.Min(1.5 * (1 + (0.012 * RH) ^ 18) * h_0 + 250 * alpha_3, 1500 * alpha_3)
    Else
     phi_RH = (1 + (1 - RH / 100) / (0.1 * h_0 ^ (1 / 3)))
        beta_h = Application.WorksheetFunction.Min(1.5 * (1 + (0.012 * RH) ^ 18) * h_0 + 250, 1500)
    End If

'calculation of beta_t0
beta_t0 = 1 / (0.1 + t_0 ^ 0.2)

'calculation of phi_0
phi_0 = phi_RH * beta_fcm * beta_t0

' calculation of beta_c_tt0


beta_c_tt0 = ((t - t_0) / (beta_h + t - t_0)) ^ 0.3

CL_EC2_CreepCo = phi_0 * beta_c_tt0

End Function

