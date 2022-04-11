Attribute VB_Name = "CL_Generalfuctions"
'#############################################################################################################
' General Function Module
'
'This Module contains general functions for the excel, which is not related to Civil Engineering Design aspect
'
'
'The function name should be all started with 'CL_G_'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC
'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit        ' all Parameters must be defined first





' Function: TableInterpolation ---------------------------------------------------------------------------------

Function CL_G_TabelInterpolation(A1 As Range, A2 As Range, input_ As Double) As Double
    
    'Automatic find value x in the first list (List 1) and calculate the corresponding value
    'in second list (List 2) based on linear interpolation (no direction limits)
    ' notice the list1 should be in a certain order,only increase or decrease order
    
    Dim A_in        As Variant        ' the list of value in list 1
    Dim A_ck        As Variant        ' the list of vaule in list 2
    Dim i           As Integer        ' internal use parameter
    Dim l_in        As Integer        ' the number of list 1
    Dim i_1, i_2, O_1, O_2 As Double
    
    ' Reg2List : internal function
    ' change the different input parameters to a one row list
    
    A_in = Reg2List(A1)
    A_ck = Reg2List(A2)
    l_in = UBound(A_in)
    
    For i = 1 To l_in - 1
        
        If (A_in(i) - input_) * (A_in(i + 1) - input_) <= 0 Then
            ' function do the interpolation
            i_1 = A_in(i)
            i_2 = A_in(i + 1)
            
            O_1 = A_ck(i)
            O_2 = A_ck(i + 1)
            
            CL_G_TabelInterpolation = O_1 + (O_2 - O_1) / (i_2 - i_1) * (input_ - i_1)
        End If
        
    Next i
    
End Function

Function CL_G_TableFind(table As Range, Lookup_value_row As String, lookup_value_column As String) As Variant
    ' find the value inside a Table which is based on mathing the first row and colomn
    
    Dim rg          As Range
    Dim x           As Integer
    Dim y           As Integer
    
    Set rg = table
    
    x = WorksheetFunction.Match(Lookup_value_row, rg.Columns(1).Cells, 0)
    y = WorksheetFunction.Match(lookup_value_column, rg.Rows(1), 0)
    
    CL_G_TableFind = rg.Cells(x, y).value
    
End Function

Function CL_G_ListTableFind(list As Range, Lookup_value_row As String, lookup_value_column As String) As Variant
    ' find the value in a list table as first row and colum used for inputs.
    ' it is useful  as now the talbe can be flexible to be expended or moved.
    Dim namerow, rg As Range
    Dim x           As Integer
    Dim y           As Integer
    x = 0
    Set namerow = list.ListObject.HeaderRowRange
    Set rg = list.Columns(1)
    x = WorksheetFunction.Match(lookup_value_column, namerow, 0)
    y = WorksheetFunction.Match(Lookup_value_row, rg, 0)
    CL_G_ListTableFind = list.Cells(y, x).value
    
End Function



' Table picker to generate a dropdown list of tablelist row or columns
Function CL_G_CreateListTablePicker(Tablelist As Range, Optional direction As String = "row", Optional offset As Integer = 1) As Variant
    ' First input is the Listtable range
    ' Direction if empty then by row name otherwise by column
    ' offset is which column default first column
    
    Dim Ad_, items  As String
    Dim fistvalue   As Variant
    
    Ad_ = ActiveCell.address
    
    'Create list items
    If direction = "column" Or direction = "c" Or direction = "1" Then
        items = RangeToString(Tablelist.Columns(offset).Cells)
        fistvalue = Tablelist.Columns(offset).Cells(1).value
    Else
        items = RangeToString(Tablelist.ListObject.HeaderRowRange.Cells)
        fistvalue = Tablelist.ListObject.HeaderRowRange.Cells(1).value
    End If
    
    With Range(Ad_).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
             xlBetween, Formula1:=items
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    '  ActiveCell.Formula = ""
    
    CL_G_CreateListTablePicker = fistvalue
    
End Function

'*************************************************************

'To reuse a calculationsheet or part of it as a Function! With it, it can recalute the table by replace the parameters and
'copy out the results needed
' the parameters are unlimited, but should only be range, the first part and second part are divided by "#"
' format CL_G_TableFunction(inputLocation1, input1, inputLocation2, input 2,..., "#" outputLocation1, output1,outputLocation2,output2...)
' each inputs (the second value) will replace the value of inputlocation value.
' each new outputlocation values will then copied to the output cells.

'**************************************************************

' overall control function, by changing the global variables and change formula

Private Function CL_G_TableFunction(ParamArray list1()) As String
    
    Dim list1_()    As Variant
    Dim list()      As Variant
    Dim i           As Integer

    'resort the data
    
    
    TBF_doit = True
    
    tbf_str = Selection.formula
    tbf_str = Replace(tbf_str, "=", "")
    tbf_list = list1
    
    'TBF_done = True
    
    TBF_address = ActiveCell.Parent.Name & "!" & ActiveCell.address
    
End Function

'**************************************************************

' procedure to carry out the work

Public Sub TBFunction()
    
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim j1          As Integer
    
    Dim pr_         As Variant
    Dim pr1_        As Variant
    Dim storage()   As String
    
    Dim list()      As Range
    Dim list1()     As Variant
    Dim reg1        As Range
    Dim reg2        As Range
    ' change the formula incase of repeating loops
   ' ThisWorkbook.ActiveSheet.Range(TBF_address).Interior.ColorIndex = 37
    
    
    Range(TBF_address).formula = tbf_str
    TBF_doit = False
    list1 = tbf_list
    
    ' checking input pairs
    i = 0
    For Each pr_ In list1
        If VarType(pr_) = 8204 Then
            i = i + 1
        Else
            If Not pr_ = "#" Then
                i = i + 1
            Else
                Exit For
            End If
        End If
    Next pr_
    
    If Not i Mod 2 = 0 Then
        MsgBox "input data Not in pair "
        Exit Sub
    End If
    
    If Not ((UBound(list1) - i) Mod 2 = 0) Then
        MsgBox " output data Not in pair"
        Exit Sub
    End If
    
    ReDim storage(0 To i, 0 To 100, 0 To 100)
    
    'replace the inputs
    For j = 1 To i / 2
        Set reg1 = list1(2 * j - 1)
        Set reg2 = list1(2 * j - 2)
        
        storage = Range2Arr(storage, j, reg2)
        procopy reg1, reg2, False
        
    Next j
    'copy out the outputs
    j = j * 2 - 1
    
    For k = 1 To (UBound(list1) - i) / 2
        Set reg1 = list1(2 * k - 2 + j)
        Set reg2 = list1(2 * k - 1 + j)
        procopy reg1, reg2, False
        
    Next k
    'copy back the origional inputs
    
    For j1 = 1 To i / 2
        Set reg1 = list1(2 * j1 - 2)
        
        Arr2Range storage, j1, reg1
    Next j1
    'select back the string
    
    Range(TBF_address).Select
End Sub

Private Function CL_G_CountSting(str As String, Optional aimstr As String = ".")

CL_G_CountSting = Len(str) - Len(Replace(str, aimstr, ""))
End Function

Private Sub CL_G_Addformula(reg0 As Range, reg1 As Range)
If reg0.formula = "" Then
reg0.formula = "=" + reg1.address
Else
reg0.formula = reg0.formula + "+" + reg1.address
End If


End Sub

Public Sub CL_G_LevelSum(reg0 As Range, reg1 As Range, Optional ifformat As Boolean = True, Optional indicator As String = ".")

If reg0.Count = reg1.Count Then
Dim nrows As Integer 'the total rows
nrows = reg0.Count
Dim levels(10000) As Integer
Dim Reseted(10000) As Integer
Dim i, j, levelmin As Integer
'to calcualte the levels
levelmin = 100
For i = 1 To nrows
levels(i) = CL_G_CountSting(reg0.Cells(i), indicator)
Reseted(i) = 1
If levels(i) < levelmin Then
levelmin = levels(i)

End If

Next i
'to add sum values

Dim p As Double
Dim plim As Integer


plim = nrows

For i = 1 To nrows
    p = levels(i)

    If i = plim Then
    plim = nrows
    End If
    
    For j = i + 1 To plim
    
    If levels(j) <= p Then
      plim = j
      Exit For
    ElseIf levels(j) = p + 1 Then
    'remove the orginal formula
        If Reseted(i) = 1 Then
        reg1.Cells(i).formula = ""
        Reseted(i) = 0
        
        End If
        CL_G_Addformula reg1.Cells(i), reg1.Cells(j)
    
    End If
    
    Next j
    
    'formatting the cell rows
    If ifformat Then
    
    If p = levelmin Then
        reg1.Cells(i).Font.Size = 14
        reg1.Cells(i).Font.Bold = True
        reg1.Cells(i).Interior.ThemeColor = xlThemeColorDark2
        
        reg0.Cells(i).Font.Size = 14
        reg0.Cells(i).Font.Bold = True
        reg0.Cells(i).Interior.ThemeColor = xlThemeColorDark2
        
     ElseIf p = levelmin + 1 Then
        reg1.Cells(i).Font.Size = 12
        reg1.Cells(i).Font.Bold = True
        reg0.Cells(i).Font.Size = 12
        reg0.Cells(i).Font.Bold = True
     ElseIf p = levelmin + 2 Then
     
        reg1.Cells(i).Font.Size = 11
        reg1.Cells(i).Font.Bold = False
        reg0.Cells(i).Font.Size = 11
        reg0.Cells(i).Font.Bold = False
     
     ElseIf p = levelmin + 3 Then
     
            reg1.Cells(i).Font.Size = 9
            reg1.Cells(i).Font.Bold = False
            reg1.Cells(i).Font.Italic = True

            reg0.Cells(i).Font.Size = 9
            reg0.Cells(i).Font.Bold = False
            reg0.Cells(i).Font.Italic = True
    End If
    
    End If
    




Next i


Else
'to report error
End If



End Sub

Public Sub CL_G_GroupColor(grouplist As Range, ref As Range, Optional refcol As Integer = 1)

Dim index0, rown, coln, i As Integer

rown = ref.Rows.Count

For i = 1 To rown

On Error Resume Next

index0 = WorksheetFunction.Match(ref.Cells(i, refcol), grouplist, 0)
ref.Rows(i).Interior.Color = grouplist(index0).Interior.Color
   
Next


End Sub


Public Function CL_G_VlookupTable(lookupvalue As Range, aimtable As Range, Optional aimname As String = "", Optional refname As String = "")
'Do advanced Vlookup in a formatted list table, based on item names instead of the column number
'lookup_value: the value to be looked up
'Aim_table: the table contains the data to be looked up, must be a formatted table (Home> Styles > Format as table)
'Aim_name : [optional] the column name where the return values are defined. The function takes the column name of the current cell location
'Ref_name : [optional] the column name where the reference values are defined. by default takes the column name of the lookup value.

Dim col1, col2, coli, ref1, ref2, aim1, aim2 As Integer
Dim refcolumn, aimcolumn As Range

'get the reference row name
If refname = "" Then

col1 = lookupvalue.Column - lookupvalue.ListObject.Range.Column + 1
refname = lookupvalue.ListObject.HeaderRowRange.Cells(col1).Text
End If

If aimname = "" Then
'get the aim row name
col2 = Application.Caller.Column - lookupvalue.ListObject.Range.Column + 1
aimname = lookupvalue.ListObject.HeaderRowRange.Cells(col2).Text
End If

'get the index of the ref
ref1 = WorksheetFunction.Match(refname, aimtable.ListObject.HeaderRowRange, 0)
aim1 = WorksheetFunction.Match(lookupvalue, aimtable.Columns(ref1), 0)

ref2 = WorksheetFunction.Match(aimname, aimtable.ListObject.HeaderRowRange, 0)
CL_G_VlookupTable = aimtable.Cells(aim1, ref2)
End Function


Public Sub CL_G_ColorItHex(rg As Range, Optional codetype As Integer = 1, Optional location_ As Integer = 1)

'codetype: 1 hex #00ff00 2 rgb 128,128,128
'location : 1 :background 2: forground

Dim i, n, R, G, B, c1, c2 As Integer
Dim value As String

n = rg.Count

For i = 1 To n

value = rg.Cells(i)

On Error Resume Next
If codetype = 1 Then 'hextype of data

R = WorksheetFunction.Hex2Dec(Mid(value, 2, 2))
G = WorksheetFunction.Hex2Dec(Mid(value, 4, 2))
B = WorksheetFunction.Hex2Dec(Mid(value, 6, 2))

ElseIf codetype = 2 Then 'RGBnumbers

c1 = WorksheetFunction.Find(",", value) 'first comma
c2 = WorksheetFunction.Find(",", value, c1 + 1) 'second comma

R = Mid(value, 1, c1 - 1)
G = Mid(value, c1 + 1, c2 - c1 - 1)
B = Mid(value, c2 + 1, 3)

End If


If location_ = 1 Then
rg.Cells(i).Interior.Color = RGB(R, G, B)
ElseIf location_ = 2 Then
rg.Cells(i).Font.Color = RGB(R, G, B)
End If

Next
End Sub


