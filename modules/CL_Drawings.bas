Attribute VB_Name = "CL_Drawings"
'#############################################################################################################
'Drawing Function Module
'
'This Module help to create drawings supporting Automatic drawnings
'
'
'The function name should be all started with 'CL_Dr_'
'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC

'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first




Private Function CL_Dr_drawrecdata(h As Double, B As Double, x As Double, y As Double) As Variant

Dim xydata(1 To 3, 1 To 5) As Double

'writing the location information of points:

'point1
xydata(1, 1) = x - B / 2
xydata(2, 1) = y + h / 2
'point2
xydata(1, 2) = x - B / 2
xydata(2, 2) = y - h / 2
'point3
xydata(1, 3) = x + B / 2
xydata(2, 3) = y - h / 2
'point4
xydata(1, 4) = x + B / 2
xydata(2, 4) = y + h / 2
 'point5
xydata(1, 5) = x - B / 2
xydata(2, 5) = y + h / 2
' maximum dot number
'xydata(3, 1) = WorksheetFunction.Max(Abs(h / 2 + y), Abs(y - h / 2), Abs(x + b / 2), Abs(x - b / 2))

CL_Dr_drawrecdata = xydata
End Function


Private Function CL_Dr_drawProfileWdata(h As Double, B As Double, tf As Double, tw As Double, x As Double, y As Double) As Variant

Dim xydata(1 To 2, 1 To 13) As Double
Dim xi(1 To 2) As Double
Dim yi(1 To 2) As Double

'writing the location information of points:
xi(1) = B / 2
xi(2) = tw / 2
yi(1) = h / 2
yi(2) = h / 2 - tf
'point1
xydata(1, 1) = x - xi(1)
xydata(2, 1) = y + yi(1)
'point2
xydata(1, 2) = x - xi(1)
xydata(2, 2) = y + yi(2)
'point3
xydata(1, 3) = x - xi(2)
xydata(2, 3) = y + yi(2)
'point4
xydata(1, 4) = x - xi(2)
xydata(2, 4) = y - yi(2)
 'point5
xydata(1, 5) = x - xi(1)
xydata(2, 5) = y - yi(2)
 'point6
xydata(1, 6) = x - xi(1)
xydata(2, 6) = y - yi(1)
 'point7
xydata(1, 7) = x + xi(1)
xydata(2, 7) = y - yi(1)
 'point8
xydata(1, 8) = x + xi(1)
xydata(2, 8) = y - yi(2)
 'point9
xydata(1, 9) = x + xi(2)
xydata(2, 9) = y - yi(2)
 'point10
xydata(1, 10) = x + xi(2)
xydata(2, 10) = y + yi(2)
 'point11
xydata(1, 11) = x + xi(1)
xydata(2, 11) = y + yi(2)
'point12
xydata(1, 12) = x + xi(1)
xydata(2, 12) = y + yi(1)
'point13
xydata(1, 13) = x - xi(1)
xydata(2, 13) = y + yi(1)
' maximum dot number
'xydata(3, 1) = WorksheetFunction.Max(Abs(h / 2 + y), Abs(y - h / 2), Abs(x + b / 2), Abs(x - b / 2))

CL_Dr_drawProfileWdata = xydata
End Function


Private Function CL_Dr_drawCircledata(R As Double, dv_ As Integer, x As Double, y As Double) As Variant

Dim xydata() As Double
ReDim xydata(1 To 3, 1 To dv_ + 1)

Dim i As Integer

Dim delta As Double
delta = 3.14159265358979 * 2 / dv_

dv_ = WorksheetFunction.Max(3, dv_)


'writing the location information of points:
For i = 1 To dv_ + 1

    xydata(1, i) = Sin(i * delta) * R + x
    xydata(2, i) = Cos(i * delta) * R + y
Next i

' maximum dot number
'xydata(3, 1) = WorksheetFunction.Max(Abs(h / 2 + y), Abs(y - h / 2), Abs(x + b / 2), Abs(x - b / 2))

CL_Dr_drawCircledata = xydata
End Function


Private Function CL_Dr_Readdate(ByRef chartname As String, data As Variant, name_ As String) As Integer
Dim shp As Shape
Dim Chart As Chart
Dim datax() As Double
Dim datay() As Double
Dim srs As Series
Dim id_ As Integer

Dim chartL As Integer
Dim lsp As Integer
Dim Sp, nc, i As Integer

'process with data

nc = UBound(data, 2) - 1


ReDim datax(nc)
ReDim datay(nc)

For i = 0 To nc
    datax(i) = data(1, i + 1)
    datay(i) = data(2, i + 1)
Next i
'checking if need to create new drawing
If chartname = "" Then
    Set shp = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, 10, 10, 200, 200)
    Set Chart = shp.Chart
    
    'formatting the chart

        Chart.SetElement (msoElementPrimaryCategoryAxisNone)
        Chart.SetElement (msoElementPrimaryValueAxisNone)
        Chart.SetElement (msoElementPrimaryValueGridLinesNone)
        Chart.SetElement (msoElementPrimaryCategoryGridLinesNone)
        Chart.SetElement (msoElementChartTitleNone)

        shp.LockAspectRatio = msoTrue
  ' get the chart name
        chartname = Chart.Name
        lsp = InStr(chartname, " ")
        chartL = Len(chartname)
        Sp = chartL - lsp
        chartname = Right(chartname, Sp)

Else
    
    Set Chart = ActiveSheet.ChartObjects(chartname)

End If
id_ = Chart.SeriesCollection.Count + 1
Set srs = Chart.SeriesCollection.NewSeries

With srs
.Name = "ID " & id_ & name_
.XValues = datax
.Values = datay

End With

CL_Dr_Readdate = id_
CL_Dr_ChartResize (chartname)
End Function


Private Function CL_Dr_Readdate3(ByRef chartname As String, data As Variant, name_ As Variant) As Integer
Dim shp As Shape
Dim chart_ As ChartObject

Dim srs As Series
Dim id_ As Integer

Dim chartL As Integer
Dim lsp As Integer
Dim Sp, nc1, nc2, nc3, i, j As Integer

'process with data


nc3 = UBound(data, 3) - 1




If chartname = "" Then
    Set shp = ActiveSheet.Shapes.AddChart2(240, xlXYScatter, 10, 10, 200, 200)
    Set chart_ = shp.Chart
    
    'formatting the chart

        chart_.SetElement (msoElementPrimaryCategoryAxisNone)
        chart_.SetElement (msoElementPrimaryValueAxisNone)
        chart_.SetElement (msoElementPrimaryValueGridLinesNone)
        chart_.SetElement (msoElementPrimaryCategoryGridLinesNone)
        chart_.SetElement (msoElementChartTitleNone)

        shp.LockAspectRatio = msoTrue
  ' get the chart name
        chartname = chart_.Name
        lsp = InStr(chartname, " ")
        chartL = Len(chartname)
        Sp = chartL - lsp
        chartname = Right(chartname, Sp)

Else
Set chart_ = ActiveSheet.ChartObjects(chartname)
      Dim s As Series
      For Each s In chart_.Chart.SeriesCollection
        s.Delete
       Next s
  
End If


Dim datax() As Double
Dim datay() As Double

For j = 0 To nc3
nc2 = WorksheetFunction.Max(data(3, 1, j + 1) - 2, 0)


ReDim datax(0 To nc2)
ReDim datay(0 To nc2)
    For i = 0 To nc2
        datax(i) = data(1, i + 1, j + 1)
        datay(i) = data(2, i + 1, j + 1)
    Next i
    'checking if need to create new drawing
    
    id_ = chart_.Chart.SeriesCollection.Count + 1
    Set srs = chart_.Chart.SeriesCollection.NewSeries
    
    With srs
    .Name = "ID " & id_ & "-" & name_(j + 1, 1)
    .XValues = datax
    .Values = datay
    .MarkerBackgroundColor = data(3, 2, j + 1)
    .MarkerForegroundColor = data(3, 2, j + 1)
    .DataLabels(1).Format.TextFrame2.TextRange. _
        Characters.Text = name_(j + 1, 1)
    End With
    
Erase datax
Erase datay
Next j

CL_Dr_Readdate3 = id_
'CL_Dr_ChartResize (chartname)
End Function



Private Function GetArrLength(a As Variant) As Long
   If IsEmpty(a) Then
      GetArrLength = 0
   Else
      GetArrLength = UBound(a) - LBound(a) + 1
   End If
End Function


Function CL_Dr_ScatterbyGroup(Xs As Range, Ys As Range, groups As Range, Optional chartname As String = "")
Dim xygdata() As Double 'the data grouped
Dim gi() As Double
Dim i, j, k, ii, ng, nx, ny As Integer
Dim uG As Variant
Dim cg() As Double




'get the unique group
uG = WorksheetFunction.Unique(groups)
ng = GetArrLength(uG)
nx = Xs.Count

ReDim xygdata(1 To 3, 1 To nx, 1 To ng)

' 1 to 3 : 1 x 2 y 3 for count
' set the inital data
For i = 1 To ng
xygdata(3, 1, i) = 1
Next

'collect data and format info
For i = 1 To nx
k = WorksheetFunction.Match(groups.Cells(i), uG, 0)
ii = xygdata(3, 1, k)
xygdata(1, ii, k) = Xs.Cells(i).value
xygdata(2, ii, k) = Ys.Cells(i).value

xygdata(3, 1, k) = xygdata(3, 1, k) + 1 'go to next
xygdata(3, 2, k) = groups.Cells(i).Interior.Color
Next

CL_Dr_ScatterbyGroup = CL_Dr_Readdate3(chartname, xygdata, uG)


End Function




















Private Sub CL_Dr_updatedate(chartname As String, data As Variant, name_ As String, id_ As Integer)
Dim shp As Shape
Dim Chart As Chart
Dim datax() As Double
Dim datay() As Double
Dim srs As Series
Dim nc As Integer
Dim i As Integer
Dim bd_ As Double


'process with data
nc = UBound(data, 2) - 1
ReDim datax(nc)
ReDim datay(nc)

For i = 0 To nc
    datax(i) = data(1, i + 1)
   datay(i) = data(2, i + 1)
Next i


' target the chart
ActiveSheet.ChartObjects(chartname).Activate
Set Chart = ActiveChart

bd_ = Chart.Axes(xlValue).MaximumScale


Set srs = Chart.SeriesCollection(id_)
With srs
.Name = "ID " & id_ & name_
.XValues = datax
.Values = datay


End With
End Sub


Private Sub CL_Dr_ChartResize(chartname As String, Optional pr_ As Double = 1)
Dim shp As Shape
Dim Chart As Chart
Dim srs As Series
Dim bd_ As Double

Dim i As Integer
Dim j As Integer
Dim i_ As Integer
Dim j_ As Integer

Dim x As Double

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



















'***********************************************************************************
'
'
'create new shapes
'

'
'***********************************************************************************

Public Function CL_Dr_Rec(h As Double, B As Double, Optional x As Double = 0, Optional y As Double = 0, Optional chartname As String = "") As String
Dim name_ As String
Dim data_() As Double
Dim id_ As String
   
    data_ = CL_Dr_drawrecdata(h, B, x, y)
    name_ = " :Rectangle " & h & " x " & B & " @( " & x & " , " & y & ")"
    id_ = CL_Dr_Readdate(chartname, data_, name_)
    
    CL_Dr_Rec = "=CL_Dr_updateRec(" & h & "," & B & "," & x & "," & y & "," & Chr(34) & chartname & Chr(34) & "," & id_ & ") [paste string to modify]"
    'ActiveCell.Formula = ""




End Function

Public Function CL_Dr_Profile_Weld(h As Double, B As Double, tf As Double, tw As Double, Optional x As Double = 0, Optional y As Double = 0, Optional chartname As String = "") As String
Dim name_ As String
Dim data_() As Double
Dim id_ As Integer

data_ = CL_Dr_drawProfileWdata(h, B, tf, tw, x, y)
name_ = " :ProfileW " & h & " x " & B & " @( " & x & " , " & y & ")"
id_ = CL_Dr_Readdate(chartname, data_, name_)

CL_Dr_Profile_Weld = "=CL_Dr_updateProfile_weld(" & h & "," & B & "," & tf & "," & tw & "," & x & "," & y & "," & Chr(34) & chartname & Chr(34) & "," & id_ & ") [paste string to modify]"
End Function
Public Function CL_Dr_Circle(R As Double, Optional divisions As Integer = 32, Optional x As Double = 0, Optional y As Double = 0, Optional chartname As String = "") As String
Dim name_ As String
Dim data_() As Double
Dim id_ As Integer

data_ = CL_Dr_Circledata(R, divisions, x, y)
name_ = " :Circle " & R & " @( " & x & " , " & y & ")"
id_ = CL_Dr_Readdate(chartname, data_, name_)

CL_Dr_drCircle = "=CL_Dr_updateCircle(" & R & "," & divisions & "," & x & "," & y & "," & Chr(34) & chartname & Chr(34) & "," & id_ & ") [paste string to modify]"
End Function
























'***********************************************************************************
'
'
'Autoupdate shapes
'

'
'***********************************************************************************

Public Function CL_Dr_updateRec(h As Double, B As Double, x As Double, y As Double, chartname As String, id_ As Integer) As Variant
Dim name_ As String
Dim data_() As Double

data_ = CL_Dr_drawrecdata(h, B, x, y)
name_ = " :Rectangle " & h & " x " & B & " @( " & x & " , " & y & ")"
CL_Dr_updatedate chartname, data_, name_, id_
CL_Dr_ChartResize (chartname)
CL_Dr_updateRec = chartname & " :ID: " & id_ & " <UPDATE>"

End Function



Public Function CL_Dr_updateCircle(R As Double, divisions As Integer, x As Double, y As Double, chartname As String, id_ As Integer) As Variant
Dim name_ As String
Dim data_() As Double

data_ = CL_Dr_drawCircledata(R, divisions, x, y)
name_ = " :Circle " & R & " @( " & x & " , " & y & ")"
CL_Dr_updatedate chartname, data_, name_, id_
CL_Dr_ChartResize (chartname)
CL_Dr_updateCircle = chartname & " :ID: " & id_ & " <UPDATE>"

End Function

Public Function CL_Dr_updateProfile_weld(h As Double, B As Double, tf As Double, tw As Double, x As Double, y As Double, chartname As String, id_ As Integer) As Variant
Dim name_ As String
Dim data_() As Double

data_ = CL_Dr_drawProfileWdata(h, B, tf, tw, x, y)
name_ = " :ProfileW " & h & " x " & B & " @( " & x & " , " & y & ")"
CL_Dr_updatedate chartname, data_, name_, id_
CL_Dr_ChartResize (chartname)
CL_Dr_updateProfile_weld = chartname & " :ID: " & id_ & " <UPDATE>"
SendKeys ESCAPE
End Function

Public Function CL_Dr_updateProfile(SC As String, x As Double, y As Double, chartname As String, id_ As Integer) As Variant
Dim h As Double
Dim B As Double
Dim tf As Double
Dim tw As Double

Dim name_ As String
Dim data_() As Double

    h = CL_S_ProfiledSteelData(SC, "h")
    B = CL_S_ProfiledSteelData(SC, "b")
    tf = CL_S_ProfiledSteelData(SC, "tf")
    tw = CL_S_ProfiledSteelData(SC, "tw")

data_ = CL_Dr_drawProfileWdata(h, B, tf, tw, x, y)
name_ = " :Profile " & SC & " @( " & x & " , " & y & ")"
CL_Dr_updatedate chartname, data_, name_, id_
CL_Dr_ChartResize (chartname)
CL_Dr_updateProfile = chartname & " :ID: " & id_ & " <UPDATE>"



End Function



'





























