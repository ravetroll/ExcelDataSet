VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' Class Name:   cDataTable
' Description:  Used to store a table of data in memory.
'               Quickly accessible to and from Arrays
'               Ability to Load and Save from XML
'               Data filtering capability
'               No data typing required
' Created by:   Brendon Raw
' History:      v1.00 04 October 2011 - Created
'               v1.01 05 October 2011 - return values for functions, move functions
'               v1.02 10 October 2011 - Aggregate function. Renamed to cDataTable from cDataset
'               v1.03 12 October 2011 - Added SelectFields
'               v1.04 05 January 2012 - Added Sort function
'               v1.05 12 Janaury 2012 - Added AppendDataTable function
'               v1.06 17 January 2012 - Added BOF and EOF properties
'               v1.07 06 February 2012 - Added WhereFieldIn and WhereFieldNotIn
'               v1.08 08 February 2012 - Added DeleteRecord, CurrentRecord, Thoroughly Reworked indexing
'               v1.09 09 February 2012 - Added ExcludeHeader to AsArray
'               v1.10 15 February 2012 - Fixed situation where CopyField keeps adding to Default Record
'               v1.11 16 March 2012 - Added SetOrderField, added order to AddField
' References:   cRecord v1.09
'               cField v1.01
'               Microsoft XML v6
'               Microsoft Scripting Runtime
'
' Instructions: (LoadArray, AsArray) to convert to and from array
'               (BOF, EOF) Return true if we hit the begin or end of the table
'               (AddField, DeleteField, CopyField) Add/Remove/Copy named field from table
'               (SetFieldOrder) Sets the order position of the named field
'               (UpdateFieldValue) Updates field in current record
'               (CurrentRecord) Returns the record the current record pointer is on
'               (MoveFirst, MoveLast, MoveNext, MovePrevious) Move the current record pointer
'               (NewRecord) Creates a new empty current record at bottom of table
'               (CopyRecord) Copys the value of an existing field into a new field
'               (AddRecord) Adds a new cRecord object to the dictionary of records
'               (AppendDataTable) appends cDataTable to calling object
'               (IsValueInField) Returns true if a value is in a specified field
'               (FieldOrderNumber) returns the Order number of a field
'               (FieldNameFromOrder) returns field name from Order number
'               (WhereFieldBetween) returns a subset of records matching the between criteria
'               (WhereField) returns a subset of records matching the Where criteria
'               (WhereFieldIn) returns a subset of records matching the Where Array criteria
'               (WhereFieldNotIn) returns a subset of records not matching the Where Array criteria
'               (Aggregate) performs a GROUP BY, along with MIN/MAX/SUM aggregation
'               (Sort) performs ascending or descending sort on specified fields
'               (SelectFields) selects a subset of fields from the Table
'               (Copy) returns a copy of this instance
'               (Dispose) performs cleanup tasks
'               (XML) returns/loads the object as an XML object
'               (LoadXML, SaveAsXML) loads and saves object in XML file
'*************************************************************************
Public Records As Dictionary
Private colRecords As Collection
Public TableName As String
Private recDefault As cRecord
Private lngRecordCount As Long
Private lngCurrentRecord As Long
Private isDisposing As Boolean
Private isBOF As Boolean
Private isEOF As Boolean
Private intNextID As Long

Private Sub Class_Initialize()
Set Records = New Dictionary
Set colRecords = New Collection
isBOF = True
isEOF = True
intNextID = 1
End Sub

Public Property Get BOF() As Boolean
BOF = isBOF
End Property

Public Property Get EOF() As Boolean
EOF = isEOF
End Property

Private Function NextID() As Long
NextID = intNextID
intNextID = intNextID + 1
End Function

Public Function LoadArray(ar As Variant) As Boolean
Dim lngHeader As Long
Dim lngRow As Long
Dim recTemp As cRecord
For lngRow = LBound(ar, 1) + 1 To UBound(ar, 1)
    Set recTemp = New cRecord
    
    For lngHeader = LBound(ar, 2) To UBound(ar, 2)
        recTemp.AddField ar(1, lngHeader), ar(lngRow, lngHeader)
    Next lngHeader
    
    AddRecord recTemp
    
Next lngRow
Set recTemp = Records(CStr(lngCurrentRecord))
Set recDefault = recTemp.NewRecord
LoadArray = True
End Function

Public Function AsArray(Optional blnExcludeHeader As Variant) As Variant
Dim ar As Variant
Dim lngHeader As Long
Dim lngRow As Long
Dim strRec As Variant
Dim strField As Variant
Dim tmpRecord As cRecord
Dim tmpField As cField

ar = SizeArray(Records.Count + 1, colRecords(lngCurrentRecord).Fields.Count)
If IsMissing(blnExcludeHeader) Then
    Set tmpRecord = colRecords(1)
    For Each strField In tmpRecord.Fields
        Set tmpField = tmpRecord.Fields(strField)
        ar(lngRow, lngHeader) = tmpField.Name
        lngHeader = lngHeader + 1
    Next strField
    lngRow = lngRow + 1
Else
    If Not blnExcludeHeader Then
        Set tmpRecord = colRecords(1)
        For Each strField In tmpRecord.Fields
            Set tmpField = tmpRecord.Fields(strField)
            ar(lngRow, lngHeader) = tmpField.Name
            lngHeader = lngHeader + 1
        Next strField
        lngRow = lngRow + 1
    
    End If
End If
lngHeader = 0
For Each strRec In Records
    Set tmpRecord = Records(strRec)
    For Each strField In tmpRecord.Fields
        Set tmpField = tmpRecord.Fields(strField)
        ar(lngRow, lngHeader) = tmpField.Value
        lngHeader = lngHeader + 1
    Next strField
    lngRow = lngRow + 1
    lngHeader = 0
Next strRec
AsArray = ar
End Function

Private Function SizeArray(lngRow, lngCol) As Variant
Dim ar As Variant
ReDim ar(0 To lngRow - 1, 0 To lngCol - 1)
SizeArray = ar
End Function

Public Property Get RecordCount() As Long
RecordCount = lngRecordCount
End Property

Public Property Get RecordNumber() As Long
RecordNumber = lngCurrentRecord
End Property

Public Function AddField(ByVal fieldName As String, ByVal FieldValue As Variant, Optional FieldOrder As Variant) As Boolean
Dim strRec As Variant
Dim rec As cRecord
Dim fld As cField
For Each strRec In Records
    Set fld = Nothing
    Set rec = Records(strRec)
    If IsMissing(FieldOrder) Then
        Set fld = rec.AddField(fieldName, FieldValue)
    Else
        If IsNumeric(FieldOrder) Then
            Set fld = rec.AddField(fieldName, FieldValue, CLng(FieldOrder))
        End If
    End If
    If Not fld Is Nothing Then
        AddField = True
        If Not recDefault Is Nothing Then
            If Not recDefault.Fields.Exists(fld.Name) Then
                If IsMissing(FieldOrder) Then
                    recDefault.AddField fieldName, FieldValue
                Else
                    If IsNumeric(FieldOrder) Then
                        recDefault.AddField fieldName, FieldValue, CLng(FieldOrder)
                    End If
                End If
            End If
        End If
    Else
        AddField = False
        Exit Function
    End If
Next strRec
End Function

Public Function SetFieldOrder(ByVal fieldName As String, ByVal FieldOrder As Long) As Long
Dim strRec As Variant
Dim rec As cRecord
Dim lngOrder As Long
For Each strRec In Records
    Set fld = Nothing
    Set rec = Records(strRec)
    lngOrder = rec.SetFieldOrder(fieldName, FieldOrder)
    If lngOrder >= 0 Then
        SetFieldOrder = lngOrder
        If Not recDefault Is Nothing Then
            recDefault.SetFieldOrder fieldName, FieldOrder
        End If
    Else
        SetFieldOrder = lngOrder
        Exit Function
    End If
Next strRec
End Function


Public Function CopyField(ByVal sourceFieldName As String, ByVal newFieldName As String)
Dim strRec As Variant
Dim rec As cRecord
For Each strRec In Records
    Set rec = Records(strRec)
    If rec.CopyField(sourceFieldName, newFieldName) Then
        CopyField = True
        If Not recDefault Is Nothing Then
            If Not recDefault.Fields.Exists(newFieldName) Then
                recDefault.CopyField sourceFieldName, newFieldName
            End If
        End If
    Else
        CopyField = False
        Exit Function
    End If
Next strRec
End Function

Public Function DeleteField(ByVal fieldName As String) As Boolean
Dim strRec As Variant
Dim rec As cRecord
For Each strRec In Records
    Set rec = Records(strRec)
    If rec.DeleteField(fieldName) Then
        DeleteField = True
        If Not recDefault Is Nothing Then
            recDefault.DeleteField fieldName
        End If
    Else
        DeleteField = False
        Exit Function
    End If
Next strRec
End Function

Public Property Get CurrentRecord() As cRecord
If Records.Count > 0 Then
    Set CurrentRecord = colRecords(lngCurrentRecord)
End If
End Property

Public Function MoveFirst() As cRecord
Dim recCurrent As cRecord
If Records.Count >= 1 Then
    lngCurrentRecord = 1
    Set recCurrent = colRecords(lngCurrentRecord)
    
    isBOF = False
    isEOF = False
End If
Set MoveFirst = recCurrent
End Function

Public Function MoveLast() As cRecord
Dim recCurrent As cRecord
If Records.Count >= 1 Then
    lngCurrentRecord = lngRecordCount
    Set recCurrent = colRecords(lngCurrentRecord)
    isBOF = False
    isEOF = False
End If
Set MoveLast = recCurrent
End Function

Public Function MoveNext() As cRecord
Dim recCurrent As cRecord
If Records.Count >= 1 Then
    If lngCurrentRecord + 1 <= lngRecordCount Then
        lngCurrentRecord = lngCurrentRecord + 1
        Set recCurrent = colRecords(lngCurrentRecord)
        isBOF = False
    Else
        isEOF = True
    End If
Else
    isBOF = True
    isEOF = True
End If
Set MoveNext = recCurrent
End Function

Public Function MovePrevious() As cRecord
Dim recCurrent As cRecord
If Records.Count >= 1 Then
    If lngCurrentRecord - 1 >= 1 Then
        lngCurrentRecord = lngCurrentRecord - 1
        Set recCurrent = colRecords(lngCurrentRecord)
        isEOF = False
    Else
        isBOF = True
    End If
End If
Set MovePrevious = recCurrent
End Function

Public Function UpdateFieldValue(ByVal fieldName As String, ByVal FieldValue As Variant) As Boolean
Dim rec As cRecord
Set rec = colRecords(lngCurrentRecord)
If Not rec Is Nothing Then
    If rec.UpdateField(fieldName, FieldValue) Then
        UpdateFieldValue = True
    Else
        UpdateFieldValue = False
        Exit Function
    End If
End If
End Function

Public Function NewRecord() As cRecord
Dim recCurrent As cRecord
Dim recNew As cRecord
If Records.Count >= 1 Then
    If recDefault Is Nothing Then
        Set recCurrent = colRecords(lngCurrentRecord)
        Set recNew = recCurrent.NewRecord
    Else
        Set recNew = recDefault.Copy
    End If
    lngRecordCount = lngRecordCount + 1
    recNew.ID = NextID
    Records.Add CStr(recNew.ID), recNew
    colRecords.Add recNew, CStr(recNew.ID)
    lngCurrentRecord = lngRecordCount
    Set NewRecord = recNew
    isBOF = False
    isEOF = False
End If
End Function

Private Function AddFirstRecord(rec As cRecord) As Boolean
If Records.Count = 0 Then
    If recDefault Is Nothing Then
        lngRecordCount = 1
        rec.ID = NextID
        Records.Add CStr(rec.ID), rec
        colRecords.Add rec, CStr(rec.ID)
        lngCurrentRecord = lngRecordCount
        Set recDefault = rec.NewRecord
        AddFirstRecord = True
        isBOF = False
        isEOF = False
    Else
        If rec.Compare(recDefault, True) Then
            lngRecordCount = 1
            rec.ID = NextID
            Records.Add CStr(rec.ID), rec
            colRecords.Add rec, CStr(rec.ID)
            lngCurrentRecord = lngRecordCount
            AddFirstRecord = True
            isBOF = False
            isEOF = False
        End If
    End If
    
End If
End Function

Public Function DeleteRecord(Optional RecordID As Variant) As Boolean
Dim rec As cRecord
Dim lngDeleteRecord As Long


If colRecords.Count > 0 Then
    If IsMissing(RecordID) Then
        lngDeleteRecord = lngCurrentRecord
        Set rec = colRecords(lngDeleteRecord)
        colRecords.Remove lngDeleteRecord
        Records.Remove CStr(rec.ID)
    Else
        If Records.Exists(CStr(RecordID)) Then
            Set rec = colRecords(CStr(RecordID))
            colRecords.Remove CStr(RecordID)
            Records.Remove CStr(rec.ID)
        Else
            Exit Function
        End If
    End If
    
    
    
    lngRecordCount = lngRecordCount - 1
    If lngRecordCount = 0 Then
        isBOF = True
        isEOF = True
        lngCurrentRecord = 0
    End If
    If lngCurrentRecord > lngRecordCount Then
        lngCurrentRecord = lngRecordCount
        isEOF = True
    End If
    
    Set rec = Nothing
    DeleteRecord = True
End If

End Function

Public Function AddRecord(rec As cRecord) As Boolean
Dim varRec As Variant
If Records.Count > 0 Then
    Set varRec = Records(CStr(lngCurrentRecord))
    If varRec.Fields.Count = rec.Fields.Count Then
        If recDefault Is Nothing Then
            lngRecordCount = lngRecordCount + 1
            rec.ID = NextID
            Records.Add CStr(rec.ID), rec
            colRecords.Add rec, CStr(rec.ID)
            lngCurrentRecord = lngRecordCount
            AddRecord = True
            isBOF = False
            isEOF = False
        Else
            If recDefault.Compare(rec, True) Then
                lngRecordCount = lngRecordCount + 1
                rec.ID = NextID
                Records.Add CStr(rec.ID), rec
                colRecords.Add rec, CStr(rec.ID)
                lngCurrentRecord = lngRecordCount
                AddRecord = True
                isBOF = False
                isEOF = False
            End If
        End If
    End If
Else
    AddRecord = AddFirstRecord(rec)
End If
End Function

Public Function AppendDataTable(tbl As cDataTable) As Boolean
Dim varRec As Variant
Dim blnError As Boolean
For Each varRec In tbl.Records
    If Not AddRecord(tbl.Records(varRec)) Then
        AppendDataTable = False
        Exit Function
    End If
Next varRec
AppendDataTable = True
End Function

Public Function IsValueInField(ByVal fieldName As String, ByVal FieldValue As Variant) As Boolean
Dim strRec As Variant
Dim rec As cRecord
IsValueInField = False
For Each strRec In Records
    Set rec = Records(strRec)
    If rec.IsFieldValue(fieldName, FieldValue) Then IsValueInField = True
Next strRec
End Function

Public Function FieldOrderNumber(ByVal fieldName As String) As Long
Dim rec As cRecord
Set rec = Records(lngCurrentRecord)
If rec.Fields.Exists(fieldName) Then
    FieldOrderNumber = rec.FieldOrderNumber(fieldName)
End If
End Function

Public Function FieldNameFromOrder(ByVal FieldOrder As Long) As String
Dim rec As cRecord
Set rec = Records(CStr(lngCurrentRecord))
FieldNameFromOrder = rec.FieldNameFromOrder(FieldOrder)
End Function

Public Function WhereFieldBetween(ByVal fieldName As String, ByVal smallValue As Variant, ByVal largeValue As Variant) As cDataTable
Dim strRec As Variant
Dim rec As cRecord
Dim tmpDS As cDataTable
Set tmpDS = New cDataTable
For Each strRec In Records
    Set rec = Records(strRec)
    If rec.IsFieldValueBetween(fieldName, smallValue, largeValue) Then
        tmpDS.AddRecord rec.Copy
    End If
Next strRec
Set WhereFieldBetween = tmpDS
End Function

Public Function WhereField(ByVal fieldName As String, ByVal FieldValue As Variant, Optional compareRule As Variant) As cDataTable
Dim ds As cDataTable
Dim rec As cRecord
Dim strCompareRule As String

Set WhereField = Nothing
If IsMissing(compareRule) Then
    strCompareRule = "="
Else
    strCompareRule = compareRule
End If
Set ds = New cDataTable
For Each strRec In Records
    Set rec = Records(strRec)
    If rec.IsFieldValue(fieldName, FieldValue, strCompareRule) Then
        ds.AddRecord rec.Copy
    End If
Next strRec
Set WhereField = ds
End Function

Public Function WhereFieldIn(ByVal fieldName As String, ParamArray fieldValues() As Variant) As cDataTable
Dim ds As cDataTable
Dim rec As cRecord
Dim strCompareRule As String
Dim lngCount As Long
Set WhereFieldIn = Nothing

If Not IsEmpty(fieldValues) Then
    strCompareRule = "="
    
    Set ds = New cDataTable
    For Each strRec In Records
        Set rec = Records(strRec)
        For lngCount = LBound(fieldValues) To UBound(fieldValues)
            If rec.IsFieldValue(fieldName, fieldValues(lngCount), strCompareRule) Then
                ds.AddRecord rec.Copy
                Exit For
            End If
        Next lngCount
    Next strRec
End If
Set WhereFieldIn = ds
End Function

Public Function WhereFieldNotIn(ByVal fieldName As String, ParamArray fieldValues() As Variant) As cDataTable
Dim ds As cDataTable
Dim rec As cRecord
Dim strCompareRule As String
Dim lngCount As Long
Dim blnFound As Boolean
Set WhereFieldNotIn = Nothing

If Not IsEmpty(fieldValues) Then
    strCompareRule = "="
    
    Set ds = New cDataTable
    For Each strRec In Records
        Set rec = Records(strRec)
        blnFound = False
        For lngCount = LBound(fieldValues) To UBound(fieldValues)
            If rec.IsFieldValue(fieldName, fieldValues(lngCount), strCompareRule) Then
                blnFound = True
                Exit For
                
            End If
        Next lngCount
        If Not blnFound Then
            ds.AddRecord rec.Copy
        End If
    Next strRec
Else
    Set ds = Me.Copy
End If
Set WhereFieldNotIn = ds
End Function

Public Function DefaultRecord(ByVal blnDefaults As Boolean) As cRecord
Dim recGrp As cRecord
If Not recDefault Is Nothing Then
    If blnDefaults Then
        Set recGrp = recDefault.Copy
    Else
        Set recGrp = recDefault.NewRecord
    End If
End If
Set DefaultRecord = recGrp
End Function

Public Function Aggregate(ByVal rec As cRecord) As cDataTable
Dim varRec As Variant
Dim varDat As Variant
Dim recRec As cRecord
Dim recDat As cRecord
Dim varFld As Variant
Dim fldFld As cField
Dim varAgg As Variant
Dim dat As cDataTable
Dim blnFieldsMatch As Boolean
Dim blnRecordsMatch As Boolean
Dim varRecNumber As Variant
Dim varDatNumber As Variant

Set dat = New cDataTable

For Each varRec In Records
    blnRecordsMatch = False
    Set recRec = Records(varRec)
    If dat.RecordCount = 0 Then
        Set recDat = recRec.Copy
        For Each varFld In rec.Fields
            Set fldFld = rec.Fields(varFld)
            Select Case fldFld.Value
            Case "SUM"
                If Not IsNumeric(recDat.FieldValue(fldFld.Name)) Then
                    recDat.UpdateField fldFld.Name, 0#
                    
                End If
            Case "MIN", "MAX"
                If Not IsNumeric(recDat.FieldValue(fldFld.Name)) Then
                    recDat.UpdateField fldFld.Name, Null
                    
                End If
            Case "COUNT"
                recDat.UpdateField fldFld.Name, 1
            End Select
        Next varFld
        dat.AddRecord recDat
    Else
        For Each varDat In dat.Records
            Set recDat = dat.Records(varDat)
            If rec.Compare(recDat, True) And rec.Compare(recRec, True) Then
                blnFieldsMatch = True
                
                For Each varFld In rec.Fields
                    Set fldFld = rec.Fields(varFld)
                    If fldFld.Value = "GROUP BY" Or IsNull(fldFld.Value) Then
                        
                        If recRec.FieldValue(fldFld.Name) <> recDat.FieldValue(fldFld.Name) Then
                            blnFieldsMatch = False
                            
                        End If
                    End If
                Next varFld
                If blnFieldsMatch Then
                    blnRecordsMatch = True
                    For Each varFld In rec.Fields
                        Set fldFld = rec.Fields(varFld)
                        varRecNumber = recRec.FieldValue(fldFld.Name)
                        varDatNumber = recDat.FieldValue(fldFld.Name)
                        If IsNumeric(varRecNumber) Then varRecNumber = CDbl(varRecNumber)
                        If IsNumeric(varDatNumber) Then varDatNumber = CDbl(varDatNumber)
                        Select Case fldFld.Value
                        Case "SUM"
                            If IsNumeric(varRecNumber) And IsNumeric(varDatNumber) Then
                                recDat.UpdateField fldFld.Name, varDatNumber + varRecNumber
                            End If
                        Case "MIN"
                            If IsNumeric(varRecNumber) And IsNumeric(varDatNumber) Then
                                If varDatNumber > varRecNumber Then
                                    recDat.UpdateField fldFld.Name, varRecNumber
                                End If
                            End If
                        Case "MAX"
                            If IsNumeric(varRecNumber) And IsNumeric(varDatNumber) Then
                                If varDatNumber < varRecNumber Then
                                    recDat.UpdateField fldFld.Name, varRecNumber
                                End If
                            End If
                        Case "COUNT"
                            recDat.UpdateField fldFld.Name, recDat.FieldValue(fldFld.Name) + 1
                        Case Else
                        
                        End Select
                    Next varFld
                    Exit For ' Because the record was found we can now exit the loop
                Else
                    
                End If
            Else
                Set Aggregate = Nothing
                Exit Function
            End If
        Next varDat
        
        If Not blnRecordsMatch Then
            Set recDat = recRec.Copy
            For Each varFld In rec.Fields
                Set fldFld = rec.Fields(varFld)
                Select Case fldFld.Value
                Case "SUM"
                    If Not IsNumeric(recDat.FieldValue(fldFld.Name)) Then
                        recDat.UpdateField fldFld.Name, 0#
                        
                    End If
                Case "MIN", "MAX"
                    If Not IsNumeric(recDat.FieldValue(fldFld.Name)) Then
                        recDat.UpdateField fldFld.Name, Null
                        
                    End If
                Case "COUNT"
                    recDat.UpdateField fldFld.Name, 1
                End Select
            Next varFld
            dat.AddRecord recDat
        
        
        End If
    End If
Next varRec
Set Aggregate = dat
End Function

Public Function Sort(ByVal recSort As cRecord) As cDataTable
Dim datSorted As cDataTable
Dim colSorting As Collection
Dim varRec As Variant
Dim recRec As cRecord
Dim varCol As Variant
Dim recCol As cRecord
Dim lngCounter As Long
Dim lngCounterStore As Long
Dim lngCurrent As Long
Dim lngPrevious As Long
Dim lngStartCountDown As Long
Dim lngStartCountDownPrevious As Long
Set colSorting = New Collection
For Each varRec In Records
    Set recRec = Records(varRec)
    Set recRec = recRec.Copy
    If colSorting.Count = 0 Then
        colSorting.Add recRec, varRec
    Else
        lngPrevious = 0
        lngCurrent = 0
        lngStartCountDown = colSorting.Count
        lngStartCountDownPrevious = colSorting.Count
        
        If colSorting.Count > 2000 Then
            
            For lngCounter = lngStartCountDown To 1 Step -1000
                lngCounterStore = lngCounter
                Set varRec = colSorting(lngCounter)
                Set recCol = varRec
                If recRec.GreaterOrSmallerThan(recCol, recSort) = -1 Then
                    'Debug.Print ("Smaller " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = -1
                ElseIf recRec.GreaterOrSmallerThan(recCol, recSort) = 1 Then
                    'Debug.Print ("Larger " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 1
                Else
                     'Debug.Print ("Equal " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 0
                End If
                lngStartCountDown = lngCounter
                If lngPrevious = 1 And lngCurrent <= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent = 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = -1 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                lngStartCountDownPrevious = lngStartCountDown
            Next lngCounter
        Else
            lngStartCountDown = colSorting.Count
        End If
        lngPrevious = 0
        lngCurrent = 0
        If colSorting.Count > 200 Then
            
            For lngCounter = lngStartCountDown To 1 Step -100
                lngCounterStore = lngCounter
                Set varRec = colSorting(lngCounter)
                Set recCol = varRec
                If recRec.GreaterOrSmallerThan(recCol, recSort) = -1 Then
                    'Debug.Print ("Smaller " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = -1
                ElseIf recRec.GreaterOrSmallerThan(recCol, recSort) = 1 Then
                    'Debug.Print ("Larger " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 1
                Else
                     'Debug.Print ("Equal " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 0
                End If
                lngStartCountDown = lngCounter
                If lngPrevious = 1 And lngCurrent <= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent = 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = -1 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                lngStartCountDownPrevious = lngStartCountDown
            Next lngCounter
        Else
            lngStartCountDown = colSorting.Count
        End If
        lngPrevious = 0
        lngCurrent = 0
        If colSorting.Count > 20 Then
            
            For lngCounter = lngStartCountDown To 1 Step -10
                lngCounterStore = lngCounter
                Set varRec = colSorting(lngCounter)
                Set recCol = varRec
                If recRec.GreaterOrSmallerThan(recCol, recSort) = -1 Then
                    'Debug.Print ("Smaller " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = -1
                ElseIf recRec.GreaterOrSmallerThan(recCol, recSort) = 1 Then
                    'Debug.Print ("Larger " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 1
                Else
                     'Debug.Print ("Equal " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                    lngPrevious = lngCurrent
                    lngCurrent = 0
                End If
                lngStartCountDown = lngCounter
                If lngPrevious = 1 And lngCurrent <= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent = 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = -1 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                If lngPrevious = 0 And lngCurrent >= 0 Then lngStartCountDown = lngStartCountDownPrevious: Exit For
                lngStartCountDownPrevious = lngStartCountDown
            Next lngCounter
        Else
            lngStartCountDown = colSorting.Count
        End If
        lngPrevious = 0
        lngCurrent = 0
        For lngCounter = lngStartCountDown To 1 Step -1
            lngCounterStore = lngCounter
            Set varRec = colSorting(lngCounter)
            Set recCol = varRec
            If recRec.GreaterOrSmallerThan(recCol, recSort) = -1 Then
                'Debug.Print ("Smaller " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                lngPrevious = lngCurrent
                lngCurrent = -1
            ElseIf recRec.GreaterOrSmallerThan(recCol, recSort) = 1 Then
                'Debug.Print ("Larger " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                lngPrevious = lngCurrent
                lngCurrent = 1
            Else
                 'Debug.Print ("Equal " & recRec.FieldValue("Column1") & ":" & recCol.FieldValue("Column1"))
                lngPrevious = lngCurrent
                lngCurrent = 0
            End If
            If lngPrevious = 1 And lngCurrent <= 0 Then Exit For
            If lngPrevious = 0 And lngCurrent = 0 Then Exit For
            If lngPrevious = -1 And lngCurrent >= 0 Then Exit For
            If lngPrevious = 0 And lngCurrent >= 0 Then Exit For
        Next lngCounter
'        colSorting.Add recRec, CStr(colSorting.Count + 1)
        If lngPrevious = 0 And lngCurrent = 0 Then
            colSorting.Add recRec, CStr(colSorting.Count + 1)
        ElseIf lngPrevious = 1 And lngCurrent <= 0 Then
            colSorting.Add recRec, CStr(colSorting.Count + 1), lngCounterStore
        ElseIf lngPrevious = 0 And lngCurrent >= 0 Then
            colSorting.Add recRec, CStr(colSorting.Count + 1)
        ElseIf lngPrevious = -1 And lngCurrent >= 0 Then
            colSorting.Add recRec, CStr(colSorting.Count + 1), , lngCounterStore
        ElseIf lngCounterStore = 1 Then
            colSorting.Add recRec, CStr(colSorting.Count + 1), lngCounterStore
        End If
    End If
Next varRec
Set datSorted = New cDataTable
For Each varRec In colSorting
    Set recCol = varRec
    datSorted.AddRecord recCol
    
Next varRec
Set Sort = datSorted
End Function

Public Function SelectFields(ParamArray fieldList() As Variant) As cDataTable
Dim varTemp As Variant
Dim varArray() As Variant
Dim recTemp As cRecord
Dim recAdd As cRecord
Dim datReturn As cDataTable
varArray = fieldList
Set datReturn = New cDataTable
For Each varTemp In Records
    Set recTemp = Records(varTemp)
    Set recAdd = recTemp.SelectFields(varArray)
    datReturn.AddRecord recAdd
Next varTemp
Set SelectFields = datReturn
End Function

Public Function Copy() As cDataTable
Dim ds As cDataTable
Dim rec As cRecord
Set ds = New cDataTable
For Each strRec In Records
    Set rec = Records(strRec)
    ds.AddRecord rec.Copy
Next strRec
Set Copy = ds
End Function

Public Function Dispose()
Dim strRec As Variant
Dim rec As cRecord
Dim fld As New cField
isDisposing = True
For Each strRec In Records
    Set rec = Records(strRec)
    rec.Dispose
Next strRec
End Function

Public Property Get XML() As MSXML2.IXMLDOMElement
Dim strRec As Variant
Dim rec As cRecord
Dim doc As MSXML2.DOMDocument60
Dim baseElement As MSXML2.IXMLDOMElement
Dim tempElement As MSXML2.IXMLDOMElement
Dim nameElement As MSXML2.IXMLDOMElement
If Not isDisposing Then
    Set doc = New MSXML2.DOMDocument60
    Set baseElement = doc.createElement("cDataTable")
    Set nameElement = doc.createElement("TableName")
    nameElement.nodeTypedValue = TableName
    Set tempElement = doc.createElement("Records")
    For Each strRec In Records
        Set rec = Records(strRec)
        tempElement.appendChild rec.XML
    Next strRec
    baseElement.appendChild nameElement
    baseElement.appendChild tempElement
    Set XML = baseElement
End If
End Property

Public Property Set XML(xmlIN As MSXML2.IXMLDOMElement)
Dim xmlRecords As MSXML2.IXMLDOMNodeList
Dim xmlTableName As MSXML2.IXMLDOMNodeList
Dim tempElement As MSXML2.IXMLDOMElement
Dim tempRecord As cRecord
Set Records = New Dictionary
lngRecordCount = 0
lngCurrentRecord = 0
If xmlIN.nodeName = "cDataTable" Then
    Set xmlTableName = xmlIN.getElementsByTagName("TableName")
    TableName = xmlTableName.Item(0).nodeTypedValue
    Set xmlRecords = xmlIN.getElementsByTagName("cRecord")
    For Each tempElement In xmlRecords
        Set tempRecord = New cRecord
        Set tempRecord.XML = tempElement
        AddRecord tempRecord
    Next tempElement
End If
End Property

Public Function SaveAsXML(ByVal strFileName As String) As Boolean
Dim doc As MSXML2.DOMDocument60
Dim proc As MSXML2.IXMLDOMProcessingInstruction
Set doc = New MSXML2.DOMDocument60
Set proc = doc.createProcessingInstruction("xml", "version='1.0'")
doc.appendChild proc
doc.async = False
doc.appendChild XML
doc.Save strFileName
SaveAsXML = True
End Function

Public Function LoadXML(ByVal strFileName As String) As Boolean
Dim doc As MSXML2.DOMDocument60
Dim tempNode As MSXML2.IXMLDOMNodeList
Dim tempElement As MSXML2.IXMLDOMElement
Set doc = New MSXML2.DOMDocument60
doc.async = False
doc.Load strFileName
lngCurrentRecord = 0
lngRecordCount = 0
Set Records = New Dictionary
Set tempNode = doc.getElementsByTagName("cDataTable")
Set tempElement = tempNode.Item(0)
Set Me.XML = tempElement
LoadXML = True
End Function
