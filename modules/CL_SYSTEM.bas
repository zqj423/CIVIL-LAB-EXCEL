Attribute VB_Name = "CL_SYSTEM"
'#############################################################################################################
' General Function Module
'
'This Module contains internal system functions for the excel, which is private use, can not call from excel
'
'
'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC
'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first
Option Private Module ' private use




Function Reg2List(Region_1) As Variant
' Reg2List : internal function
' change the different input parameters to a one row list
    Dim i As Integer
    Dim pr_ As Variant
    Dim pr1_ As Variant
    Dim list() As Variant

    i = 0
    For Each pr_ In Region_1

        If Not VarType(pr_) = 8204 Then          ' check if is a range or a parameter list
            i = i + 1
            ReDim Preserve list(1 To i)
            list(i) = pr_
    
        Else
    
            For Each pr1_ In pr_
                i = i + 1
                ReDim Preserve list(1 To i)
                list(i) = pr1_
            Next pr1_

        End If

    Next pr_

    Reg2List = list


End Function




Public Sub CL_Dr_ChartResize(chartname As String, Optional pr_ As Double = 1)
' Resize a select chart with maximum values -x,x;  -y, y ;x=0
' pr_ is the zoom size

    Dim shp As Shape
    Dim Chart As Chart
    Dim srs As Series
    Dim bd_ As Double

    Dim i As Integer
    Dim j As Integer
    Dim i_ As Integer
    Dim j_ As Integer

    Dim xvalue() As Double
    Dim yvalue() As Double

    ActiveSheet.ChartObjects(chartname).Activate
    Set Chart = ActiveChart

    bd_ = 0
    i_ = Chart.SeriesCollection.Count
    For i = 1 To i_
        Set srs = Chart.SeriesCollection(i)
        j_ = UBound(srs.XValues)
        x = srs.XValues(1)
        For j = 1 To j_
            bd_ = WorksheetFunction.Max(Abs(srs.XValues(j)), bd_, Abs(srs.Values(j)))
        Next j
    Next i
    Chart.Axes(xlValue).MinimumScale = -bd_
    Chart.Axes(xlValue).MaximumScale = bd_
    Chart.Axes(xlCategory).MinimumScale = -bd_
    Chart.Axes(xlCategory).MaximumScale = bd_

End Sub






' GET THE CELL ADDRESS
Private Sub celladdress()
    Dim strAddress  As String
    
    strAddress = ActiveCell.address
    
    MsgBox strAddress
    
End Sub



' Transfer a range to cells


Function RangeToString(ByVal myRange As Range) As String

    If Not myRange Is Nothing Then
        Dim myCell As Range
        Dim str As String
        
        For Each myCell In myRange
        str = CStr(myCell.value)
            RangeToString = RangeToString & "," & str
        Next myCell
        'Remove extra comma
        RangeToString = Right(RangeToString, Len(RangeToString) - 1)
    End If
End Function





' get the greek char index

Function Greekchar(index_ As String) As String

Greekchar = CL_G_Table_find(Range("CLT_SYS_Greek"), index_, "Greek")

End Function




' procopy (rg1 , rg2 , CK_format )
'Copy the reg1 to rg2, can also choose if copy cell format or not

Sub procopy(rg1 As Range, rg2 As Range, CK_format As Boolean)

Dim formula As String
Dim address As String

Dim x As Integer
Dim y As Integer

Dim i As Integer
Dim j As Integer

y = rg1.Rows.Count
x = rg1.Columns.Count

rg2 = rg2.Resize(y, x)

'paste the format
If CK_format Then

rg1.Copy
'Rg2.PasteSpecial Paste:=xlPasteAllUsingSourceTheme
 rg2.PasteSpecial Paste:=xlPasteAll

Else

For i = 1 To y

    For j = 1 To x

     rg2.Cells(i, j).value = rg1.Cells(i, j).value
    
    Next j
Next i

End If
Application.CutCopyMode = False


End Sub



'store a range in the tempworksheet at line number*100

Function RangeStorage(page As Worksheet, number As Integer, rg1 As Range) As Range


Dim x As Integer
Dim y As Integer

Dim i As Integer
Dim j As Integer

Dim rg2 As Range

'Set rg2 = Worksheets("CL_Temp").Range(Cells(number * 100, 1), Cells(number * 100, 2))

page.Visible = True

y = rg1.Rows.Count
x = rg1.Columns.Count
Set rg2 = page.Cells((number - 1) * 100 + 1)
rg2 = rg2.Resize(y, x)


For i = 1 To y

    For j = 1 To x

     rg2.Cells(i, j).value = rg1.Cells(i, j).value
    
    Next j
Next i

Set RangeStorage = rg2

End Function


'store a range in the 3 diamtional array, return the array

' the array(z,y,x) has a few preset values as follow.
'       value z : the Range number z>=1
'       value y,x : the value of Range z
' the cell (z,1,0) is total column numbers of range z
' the cell (z,2,0) is total row numbers of range z
' the array (0,y,x) stores the overall paramters
' the array (0,0,0) is total range number
' the array (z,0,0) is address of range z


Function Range2Arr(list As Variant, number As Integer, rg1 As Range) As Variant


Dim x As Integer
Dim y As Integer
Dim z As Integer

Dim xu As Integer
Dim yu As Integer
Dim zu As Integer
Dim i As Integer
Dim j As Integer

Dim listin() As String

listin = list



z = number
y = rg1.Rows.Count
x = rg1.Columns.Count

zu = WorksheetFunction.Max(UBound(listin, 1), z)
xu = WorksheetFunction.Max(UBound(listin, 2), x, 3)
yu = WorksheetFunction.Max(UBound(listin, 3), y)



listin(0, 0, 0) = zu
listin(z, 0, 0) = rg1.Parent.Name & "!" & rg1.address
listin(z, 1, 0) = x
listin(z, 2, 0) = y


For i = 1 To y

    For j = 1 To x

     listin(z, i, j) = rg1.Cells(i, j).value
    
    Next j
Next i

Range2Arr = listin

End Function



' to read the data in a formated 3D array and write the the provided ranges

' the array(z,y,x) has a few preset values as follow.
'       value z : the Range number starts from 1
'       value y,x : the value of Range z, starts from 1

' the cell (z,1,0) is total column numbers of range z
' the cell (z,2,0) is total row numbers of range z

' the array (0,y,x) stores the overall paramters
' the array (0,0,0) is total range number
' the array (z,0,0) is address of range z


'***** this version only allow to use the 2D part!*****

Sub Arr2Range(list As Variant, number As Integer, ByRef rg1 As Range)


Dim x As Integer
Dim y As Integer
Dim z As Integer

Dim xu As Integer
Dim yu As Integer
Dim zu As Integer
Dim i As Integer
Dim j As Integer

Dim listin() As String

listin = list



z = number
y = listin(z, 2, 0)
x = listin(z, 1, 0)


For i = 1 To y

    For j = 1 To x

      rg1.Cells(i, j).value = listin(z, i, j)
    
    Next j
Next i



End Sub





