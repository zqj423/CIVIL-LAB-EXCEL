Attribute VB_Name = "CL_Ribbon_controls"
Private Declare PtrSafe Function ShellExecute _
  Lib "shell32.dll" Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal Operation As String, _
  ByVal Filename As String, _
  Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

Public Sub OpenCLhelp(control As IRibbonControl)
    Dim lSuccess As Long
    lSuccess = ShellExecute(0, "Open", "http://excel.civil-lab.lu")
End Sub

'Callback for main onAction
Sub openMeinMenu(control As IRibbonControl)
Main.Show vbModeless
End Sub

'Callback for scripts onAction
Sub opSSScriptM(control As IRibbonControl)
UF_SS_script.Show vbModeless
End Sub

'Callback for uniteditor onAction
Sub openunitM(control As IRibbonControl)
UF_unit.Show vbModeless
End Sub

'Callback for procopypaste onAction
Sub ProcopypastM(control As IRibbonControl)
UF_CopyPaste.Show 'vbModeless
End Sub

'Callback for dr_rec onAction
Sub DrawRecM(control As IRibbonControl)
UF_drRect.Show
End Sub

'Callback for dr_cir onAction
Sub DrawCirM(control As IRibbonControl)
UF_drCircle.Show vbModeless
End Sub

'Callback for dr_prof onAction
Sub DrawProfM(control As IRibbonControl)
UF_drProfile.Show vbModeless
End Sub

'Callback for generatelist onAction
Sub GenerateListM(control As IRibbonControl)
UF_GL.Show vbModeless
End Sub

'Callback for check
Sub CheckM(control As IRibbonControl)
UF_check.Show vbModeless
End Sub
'Callback for EC tools
Sub opECToolM(control As IRibbonControl)
UF_ECtools.Show vbModeless
End Sub

Sub opLevelSum(control As IRibbonControl)
UF_LevelSum.Show 'vbModeless
End Sub
Sub opColorGroup(control As IRibbonControl)
UF_ColorGroup.Show 'vbModeless
End Sub

Sub ColorIt(control As IRibbonControl)
UF_Colorit.Show vbModeless
End Sub

