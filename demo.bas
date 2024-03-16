Attribute VB_Name = "demo"
Sub Demo()
Attribute Demo.VB_ProcData.VB_Invoke_Func = " \n14"
Dim wb As Workbook
Dim ws As Worksheet
Dim d As cDataTable
Dim f As cDataFactory
Dim rng As Range

Set f = New cDataFactory

Set wb = Workbooks.Add

Set ws = ActiveSheet
ws.Name = "DemoAsArray"
Set d = CreateRecords1
f.InsertDataTableToCell Selection, d


Set ws = wb.Worksheets.Add
ws.Name = "DemoSelectFields"
Set d = DemoSelectFields
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoWhereField"
Set d = DemoWhereField
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoWhereFieldIn"
Set d = DemoWhereFieldIn
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoWhereFieldNotIn"
Set d = DemoWhereFieldNotIn
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoWhereFieldBetween"
Set d = DemoWhereFieldBetween
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoAggregateSUM"
Set d = DemoAggregateSUM
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoAggregateCOUNT"
Set d = DemoAggregateCOUNT
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoAggregateMIN"
Set d = DemoAggregateMIN
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoAggregateMAX"
Set d = DemoAggregateMAX
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoSort"
Set d = DemoSort
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoBOFEOF"
Set d = DemoBOFEOF
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoDataFactoryExcel"
Set d = DemoDataFactoryExcel
Set objrange = f.InsertDataTableToCell(Selection, d)
Set d = f.GetDataTableFromRange(objrange)
Set objrange = objrange.Offset(0, 10)
f.InsertDataTableToCell objrange, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoDuplicateField"
Set d = DemoDuplicateField
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoMoveOperations"
Set d = DemoMoveOperations
f.InsertDataTableToCell Selection, d

Set ws = wb.Worksheets.Add
ws.Name = "DemoExcludeHeader"
Set d = DemoExcludeHeader
f.InsertDataTableToCell Selection, d, True

Set ws = wb.Worksheets.Add
ws.Name = "DemoSetFieldOrder"
Set d = DemoSetFieldOrder
f.InsertDataTableToCell Selection, d

End Sub

Function CreateRecords1() As cDataTable
Attribute CreateRecords1.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As New cDataTable
Dim r As New cRecord
r.AddField "ID", 1
r.AddField "Column1", "USD"
r.AddField "Column2", 6
r.AddField "Column3", "Orange"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 2
r.AddField "Column1", "GBP"
r.AddField "Column2", 14.6
r.AddField "Column3", "Red"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 3
r.AddField "Column1", "GBP"
r.AddField "Column2", 14.3
r.AddField "Column3", "Orange"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 4
r.AddField "Column1", "GBP"
r.AddField "Column2", 7
r.AddField "Column3", "Orange"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 5
r.AddField "Column1", "USD"
r.AddField "Column2", 9
r.AddField "Column3", "Orange"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 6
r.AddField "Column1", "USD"
r.AddField "Column2", 1.3
r.AddField "Column3", "Green"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 7
r.AddField "Column1", "USD"
r.AddField "Column2", 9
r.AddField "Column3", "Green"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 8
r.AddField "Column1", "USD"
r.AddField "Column2", 80
r.AddField "Column3", "Green"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 9
r.AddField "Column1", "USD"
r.AddField "Column2", 90
r.AddField "Column3", "Red"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 10
r.AddField "Column1", "GBP"
r.AddField "Column2", 4.7
r.AddField "Column3", "Green"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 11
r.AddField "Column1", "GBP"
r.AddField "Column2", 19
r.AddField "Column3", "Orange"
d.AddRecord r
Set r = New cRecord
r.AddField "ID", 12
r.AddField "Column1", "USD"
r.AddField "Column2", 10
r.AddField "Column3", "Green"
d.AddRecord r
Set CreateRecords1 = d
End Function

Function DemoSelectFields() As cDataTable
Attribute DemoSelectFields.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Set d = CreateRecords1
Set d = d.SelectFields("ID", "Column2")
Set DemoSelectFields = d
End Function

Function DemoWhereField() As cDataTable
Attribute DemoWhereField.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Set d = CreateRecords1
Set d = d.WhereField("Column2", 14.3)
Set DemoWhereField = d
End Function

Function DemoWhereFieldIn() As cDataTable
Attribute DemoWhereFieldIn.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Set d = CreateRecords1
Set d = d.WhereFieldIn("Column2", 14.3, 80, 90)
Set DemoWhereFieldIn = d
End Function

Function DemoWhereFieldNotIn() As cDataTable
Attribute DemoWhereFieldNotIn.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Set d = CreateRecords1
Set d = d.WhereFieldNotIn("Column2", 14.3, 80, 90)
Set DemoWhereFieldNotIn = d
End Function

Function DemoWhereFieldBetween() As cDataTable
Attribute DemoWhereFieldBetween.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Set d = CreateRecords1
Set d = d.WhereFieldBetween("Column2", 7, 15)
Set DemoWhereFieldBetween = d
End Function

Function DemoAggregateSUM() As cDataTable
Attribute DemoAggregateSUM.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set d = d.SelectFields("Column1", "Column2", "Column3")
Set r = d.DefaultRecord(False)
r.UpdateField "Column2", "SUM"
Set d = d.Aggregate(r)
Set DemoAggregateSUM = d
End Function

Function DemoAggregateCOUNT() As cDataTable
Attribute DemoAggregateCOUNT.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set d = d.SelectFields("Column1", "Column2", "Column3")
Set r = d.DefaultRecord(False)
r.UpdateField "Column2", "COUNT"
Set d = d.Aggregate(r)
Set DemoAggregateCOUNT = d
End Function

Function DemoAggregateMIN() As cDataTable
Attribute DemoAggregateMIN.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set d = d.SelectFields("Column1", "Column2", "Column3")
Set r = d.DefaultRecord(False)
r.UpdateField "Column2", "MIN"
Set d = d.Aggregate(r)
Set DemoAggregateMIN = d
End Function

Function DemoAggregateMAX() As cDataTable
Attribute DemoAggregateMAX.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set d = d.SelectFields("Column1", "Column2", "Column3")
Set r = d.DefaultRecord(False)
r.UpdateField "Column2", "MAX"
Set d = d.Aggregate(r)
Set DemoAggregateMAX = d
End Function

Function DemoSort() As cDataTable
Attribute DemoSort.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set r = d.DefaultRecord(False)
r.UpdateField "Column1", "2"
r.UpdateField "Column2", "-3"
r.UpdateField "Column3", "1"
Set d = d.Sort(r)
Set DemoSort = d
End Function

Function DemoBOFEOF() As cDataTable
Attribute DemoBOFEOF.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set r = d.MoveFirst
Do Until d.EOF
    r.UpdateField "Column1", d.EOF
    r.UpdateField "Column2", d.BOF
    r.UpdateField "Column3", d.RecordNumber
    Set r = d.MoveNext
Loop
Set DemoBOFEOF = d
End Function

Function DemoDataFactoryExcel() As cDataTable
Attribute DemoDataFactoryExcel.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set DemoDataFactoryExcel = d
End Function

Function DemoDuplicateField() As cDataTable
Attribute DemoDuplicateField.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
d.AddField "Column1", "test"

Set DemoDuplicateField = d
End Function

Function DemoMoveOperations() As cDataTable
Attribute DemoMoveOperations.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
d.DeleteRecord
d.MovePrevious
d.DeleteRecord
d.MoveFirst
d.DeleteRecord
d.MoveNext
d.DeleteRecord
d.NewRecord
d.UpdateFieldValue "ID", 13
d.UpdateFieldValue "Column1", "ZAR"
d.UpdateFieldValue "Column2", 56
d.UpdateFieldValue "Column3", "Blue"
d.DeleteRecord "5"
Set DemoMoveOperations = d
End Function

Function DemoExcludeHeader() As cDataTable
Attribute DemoExcludeHeader.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
Set DemoExcludeHeader = d
End Function

Function DemoSetFieldOrder() As cDataTable
Dim d As cDataTable
Dim r As cRecord
Set d = CreateRecords1
d.SetFieldOrder "Column3", 0
d.SetFieldOrder "ID", 12
d.SetFieldOrder "Column1", 2
d.AddField "Column4", "New", 0
d.AddField "Column0", "New", 4
Set DemoSetFieldOrder = d
End Function
