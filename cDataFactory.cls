VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' Class Name:   cDataFactory
' Description:  Used to Get data to and From cDataTables
'               Ability to Load and Save from Excel
'
' Created by:   Brendon Raw
' History:      v1.00 06 February 2012 - Created
'               v1.01 09 February 2012 - Enable ExcludeHeader Row
' References:   cDataTable v1.11
'
'
' Instructions: (GetDataTableFromRange) Creates a datatable from a range
'               (InsertDataTableToCell) Copys a DataTable with the top left corner in specified cell
'
' TODO:         put and retrieve from Recordsets
'*************************************************************************

Public Function GetDataTableFromRange(ByVal rngIN As Object) As cDataTable

    On Error GoTo Catch
    Dim obj As cDataTable
    Dim varArray As Variant
    If rngIN.Rows.Count > 1 Then
        Set obj = New cDataTable
        varArray = rngIN
        obj.LoadArray varArray
    End If
    GoTo Finally:
Catch:

Finally:
Set GetDataTableFromRange = obj
End Function

Public Function InsertDataTableToCell(ByVal rngOUT As Object, objTable As cDataTable, Optional blnExcludeHeader As Variant) As Object
    On Error GoTo Catch
    Dim rngResized As Object
    Set rngResized = rngOUT.Resize(objTable.RecordCount + 1, objTable.DefaultRecord(True).Fields.Count)
    If IsMissing(blnExcludeHeader) Then
        rngResized = objTable.AsArray
    Else
        rngResized = objTable.AsArray(blnExcludeHeader)
    End If
Catch:

Finally:
Set InsertDataTableToCell = rngResized
End Function




