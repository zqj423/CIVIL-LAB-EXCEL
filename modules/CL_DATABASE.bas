Attribute VB_Name = "CL_DATABASE"
'#############################################################################################################
' DATABASE Function Module
'
'This Module contains DATABASE functions for the excel, which is not related to Civil Engineering Design aspect
'
'The database is stored inside the Excel page CL_Tables please do not delect it!
'The function name should be all started with 'CL_D_'
'
'
'Copyright : Qingjie, www.qingjiezhang.com ,General Copyright of CC label CC-BY-SA-NC

'#############################################################################################################

' OVERALL ENVIORONMENT SETTING ---------------------------------------------------------------------------------
Option Explicit  ' all Parameters must be defined first


'  SUPPORTING FUNCTIONS  ---------------------------------------------------------------------------------
' the function CL_G_ListTableFind is used, it the generalfunction module is not included, please uncomment the following
'code

   ' Function CL_G_ListTableFind(List As Range, Lookup_value_row As String, lookup_value_column As String) As Variant
   '     ' find the value in a list table as first row and colum used for inputs.
   '     ' it is useful as now the talbe can be flexible to be expended or moved.
   '     Dim namerow, rg As Range
   '     Dim x As Integer
   '    Dim y As Integer
   '    x = 0
   '     Set namerow = List.ListObject.HeaderRowRange
   '     Set rg = List.Columns(1)
   '     x = WorksheetFunction.Match(lookup_value_column, namerow, 0)
   '     y = WorksheetFunction.Match(Lookup_value_row, rg, 0)
   '     CL_G_ListTableFind = List.Cells(y, x).Value
   ' End Function

' MAIN FUNCTIONS---------------------------------------------------------------------------------
' to add your own function, please add table in CL_Tables
' and write a checking function you can copy the sample codes and replace the name inside <>

'# SAMPLE CODE
'Function CL_D_<Dataname>(<Rowname> As String, <Columnname> As String) As Double
''
'Dim TableName As String
'TableName = "<Tablename>"
'CL_D_<Dataname> = CL_G_ListTableFind(Range(TableName), <Rowname>, <Columnname>)

'End Function


'######################################## GET DATA ###################################################################



Function CL_D_SteelData(SteelGrade As String, datatype As String) As Double
'Checking steel material properites

Dim tablename As String
tablename = "CLT_Steel_EC3"
CL_D_SteelData = CL_G_ListTableFind(Range(tablename), SteelGrade, datatype)

End Function


Function CL_D_ConcreteData(ConcreteClass As String, datatype As String) As Double
'Checking concrete material properites

Dim tablename As String
tablename = "CLT_Con_EC2"
CL_D_ConcreteData = CL_G_ListTableFind(Range(tablename), ConcreteClass, datatype)

End Function


Function CL_D_TimberData(TimberClass As String, datatype As String) As Double
'Checking concrete material properites

Dim tablename As String
tablename = "CLT_Timber_M"
CL_D_TimberData = CL_G_ListTableFind(Range(tablename), TimberClass, datatype)

End Function



Function CL_D_EUSteelProfilesData(ProfileName As String, datatype As String) As Double
'Checking concrete material properites

Dim tablename As String
tablename = "CLT_SteelProfiles_EU"
CL_D_EUSteelProfilesData = CL_G_ListTableFind(Range(tablename), ProfileName, datatype)

End Function




'#################################### GET ITEM LIST #######################################################################

'  CL_G_CreateListTablePicker(Tablelist As Range, Optional Direction As String = "row", Optional offset As Integer = 1) As Variant
'  CL_G_CreatelistTablePicker is required.

Function CL_D_SteelItemsList()  'change
Dim tablename As String
tablename = "CLT_Steel_EC3"  ' change
CL_D_SteelItemsList = CL_G_CreateListTablePicker(Range(tablename)) 'change
End Function

Function CL_D_SteelList()
Dim tablename As String
tablename = "CLT_Steel_EC3"
CL_D_SteelList = CL_G_CreateListTablePicker(Range(tablename), "column")
End Function


Function CL_D_RebarList()
Dim tablename As String
tablename = "CLT_Rebars"
CL_D_RebarList = CL_G_CreateListTablePicker(Range(tablename), "column")
End Function


Function CL_D_TimberList()
Dim tablename As String
tablename = "CLT_Timber_M"
CL_D_TimberList = CL_G_CreateListTablePicker(Range(tablename), "column")
End Function

Function CL_D_TimberItemsList()
Dim tablename As String
tablename = "CLT_Timber_M"
CL_D_TimberItemsList = CL_G_CreateListTablePicker(Range(tablename))
End Function


Function CL_D_ConcreteItemsList()
Dim tablename As String
tablename = "CLT_Con_EC2"
CL_D_ConcreteItemsList = CL_G_CreateListTablePicker(Range(tablename))
End Function

Function CL_D_ConcreteList()
Dim tablename As String
tablename = "CLT_Con_EC2"
CL_D_ConcreteList = CL_G_CreateListTablePicker(Range(tablename), "column")
End Function



Function CL_D_EUSteelProfilesItemsList()
Dim tablename As String
tablename = "CLT_SteelProfiles_EU"
CL_D_EUSteelProfilesItemsList = CL_G_CreateListTablePicker(Range(tablename))
End Function

Function CL_D_EUSteelProfilesList()
Dim tablename As String
tablename = "CLT_SteelProfiles_EU"
CL_D_EUSteelProfilesList = CL_G_CreateListTablePicker(Range(tablename), "column")
End Function

