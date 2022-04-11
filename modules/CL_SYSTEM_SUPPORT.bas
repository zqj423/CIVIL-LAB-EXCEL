Attribute VB_Name = "CL_SYSTEM_SUPPORT"
'#############################################################################################################
' system Function Module
'
'This Module contains internal system functions for the excel, which is private use, can not call from excel
'The functions mainly for the support of execl to work well
'
'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC
'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first
Option Private Module ' private use





'************************************************************************************************************

'Register the help documents for UDFs, will runafter open the workbook




'***********************************************************************************************************

Sub RegHelpDocument()

'CL_G_ArcData

Application.MacroOptions Macro:="CL_G_ArcData", _
                         Category:="CivilLAB", _
                         HelpFile:=ActiveWorkbook.Path & "\Civil-Lab help.chm", _
                         HelpContextID:=44, _
                         Description:="calculates a result based on provided inputs", _
                         ArgumentDescriptions:=Array( _
            "Radius of the arc", _
            "distance of cutting line to the top surface", _
            "the infomation needed can be one of:  A, z, z1  Iy Iz, Wply, Wplz, Wely, Welz, U ,U1, U2")



End Sub

