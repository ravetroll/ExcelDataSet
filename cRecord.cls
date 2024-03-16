VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' Class Name:  cRecord
' Description:  Used to store a record of data in memory.
'               Ability to represent in XML
'               No data typing required
' Created by:   Brendon Raw
' History:      v1.00 04 October 2011
'               v1.01 05 October 2011
'               v1.02 10 October 2011 - added Compare and ContainsSubsetOf
'               v1.03 12 October 2011 - added SelectFields
'               v1.04 13 October 2011 - added GreaterOrSmallerThan
'               v1.05 24 October 2011 - added isDisposing to notify de-allocation
'               v1.06 06 February 2012 - added <> to IsFieldValue
'               v1.07 07 February 2012 - AddField will now put an incrementing number on the end of the fieldname if it finds duplicate names
'               v1.08 08 February 2012 - Added a Record ID field for identification in DataTable
'               v1.09 16 March 2012 - Added SetFieldOrder method, added order to AddField
' References:   cField v1.02
'               Microsoft XML v6
'               Microsoft Scripting Runtime
'
' Instructions: (AddField, DeleteField) Add/Remove named field from record
'               (SetFieldOrder) Sets the order number of a field item
'               (UpdateField) Updates field in record
'               (NewRecord) Creates a new empty record
'               (CopyRecord) Copys the value of an existing field into a new field
'               (FieldValue) Returns value of named field
'               (FieldOrderNumber) returns the Order number of a field
'               (FieldNameFromOrder) returns field name from Order number
'               (IsFieldValueBetween) returns true if a field value is between specified values
'               (IsFieldValue) returns true if a field has a specified value
'               (Compare) compares record to a different record, comparing structure and values
'               (GreaterOrSmallerThan) compares the current record to a different record based upon sort criteria
'               (ContainsSubsetOf) if a record contains without difference a subset of another record
'               (SelectFields) returns a specified set of fields from the record
'               (Copy) returns a copy of this instance
'               (Dispose) performs cleanup tasks
'               (XML) returns/loads the object as an XML object
'*************************************************************************
Public Fields As Dictionary
Private isDisposing As Boolean
Public lngID As Long

Public Property Get ID() As Long
ID = lngID
End Property

Public Property Let ID(lngIn As Long)
If lngID = -1 Then lngID = lngIn
End Property

Private Sub Class_Initialize()
Set Fields = New Dictionary
lngID = -1
End Sub

Public Function AddField(ByVal fieldName As String, ByVal FieldValue As Variant, Optional FieldOrder As Variant) As cField
Dim tmpField As cField
Dim tempFieldName As String
Dim lngCounter As Long
lngCounter = 1
Set tmpField = New cField
Set AddField = Nothing
tmpField.Name = fieldName
tmpField.Value = FieldValue
tmpField.Order = Fields.Count

If Not Fields.Exists(fieldName) Then
    Fields.Add fieldName, tmpField
Else
    tempFieldName = fieldName
    Do Until Not Fields.Exists(tempFieldName)
        tempFieldName = fieldName & CStr(lngCounter)
        lngCounter = lngCounter + 1
    Loop
    tmpField.Name = tempFieldName
    Fields.Add tempFieldName, tmpField
    lngCounter = 1
End If
If Not IsMissing(FieldOrder) Then
    If IsNumeric(FieldOrder) Then
        SetFieldOrder tmpField.Name, FieldOrder
        Set tmpField = Fields(tmpField.Name)
    End If
End If
Set AddField = tmpField
End Function

Public Function SetFieldOrder(ByVal fieldName As String, ByVal FieldOrder As Long) As Long
'Returns -1 if error, otherwise the Order Number of the Ordered Field
Dim lngOrder As Long
Dim dicNewFields As Dictionary
Dim strField As Variant
Dim rField As cField
Dim theField As cField
Dim lngCount As Long
Dim blnInserted As Boolean

If Fields.Exists(fieldName) Then
    Set theField = Fields(fieldName)
    Set theField = theField.Copy
    
    If FieldOrder > Fields.Count - 1 Then
        lngOrder = Fields.Count - 1
    Else
        If FieldOrder < 0 Then
            lngOrder = 0
        Else
            lngOrder = FieldOrder
        End If
    End If
    theField.Order = lngOrder
    Set dicNewFields = New Dictionary
    lngCount = 0
    For Each strField In Fields
        Set rField = Fields(strField)
        Set rField = rField.Copy
        
        If lngCount = theField.Order Then
            If Not dicNewFields.Exists(theField.Name) Then
                dicNewFields.Add theField.Name, theField
                
            End If
            
        End If
        If Not dicNewFields.Exists(rField.Name) And rField.Name <> theField.Name Then
            rField.Order = lngCount
            dicNewFields.Add rField.Name, rField
        Else
            lngCount = lngCount - 1
        End If
        lngCount = lngCount + 1
        
    Next strField
    If Not dicNewFields.Exists(theField.Name) Then
        dicNewFields.Add theField.Name, theField
    End If
    Set Fields = dicNewFields
    SetFieldOrder = theField.Order
Else
    SetFieldOrder = -1
End If
End Function

Public Function CopyField(ByVal sourceFieldName As String, ByVal newFieldName) As Boolean
Dim NewField As cField
Dim existingField As cField
If sourceFieldName <> newFieldName Then
    Set NewField = New cField
    If Fields.Exists(sourceFieldName) Then
        Set existingField = Fields(sourceFieldName)
        If Not AddField(newFieldName, existingField.Value) Is Nothing Then CopyField = True
    End If
End If
End Function

Public Function DeleteField(ByVal fieldName As String) As Boolean
Dim fldDelete As cField
Dim strField As Variant
Dim rField As cField
If Fields.Exists(fieldName) Then
    Set fldDelete = Fields(fieldName)
    Fields.Remove fieldName
    For Each strField In Fields
        Set rField = Fields(strField)
        If rField.Order > fldDelete.Order Then
            rField.Order = rField.Order - 1
        End If
    Next strField
    DeleteField = True
End If
End Function

Public Function UpdateField(ByVal fieldName As String, ByVal FieldValue As Variant) As Boolean
Dim tmpField As cField
UpdateField = False
If Fields.Exists(fieldName) Then
    Set tmpField = Fields(fieldName)
    tmpField.Value = FieldValue
    UpdateField = True
End If
End Function

Public Function SelectFields(fieldList As Variant) As cRecord
Dim rCopy As cRecord
Dim strField As Variant
Dim rField As cField
Dim CopyField As cField
Dim lngCounter As Long
Dim blnFieldFound As Boolean
Set rCopy = New cRecord
For Each strField In Fields
    Set rField = Fields(strField)
    blnFieldFound = False
    For lngCounter = LBound(fieldList) To UBound(fieldList)
        If UCase(rField.Name) = UCase(fieldList(lngCounter)) Then
            blnFieldFound = True
            Exit For
        End If
    Next lngCounter
    If blnFieldFound Then
        
        rCopy.AddField rField.Name, rField.Value
    End If
Next strField
Set SelectFields = rCopy
End Function

Public Function Copy() As cRecord
Dim rCopy As cRecord
Dim strField As Variant
Dim rField As cField
Dim CopyField As cField
Set rCopy = New cRecord
For Each strField In Fields
    Set rField = Fields(strField)
    Set CopyField = rField.Copy
    rCopy.Fields.Add CopyField.Name, CopyField
Next strField
Set Copy = rCopy
End Function

Public Function NewRecord() As cRecord
Dim rCopy As cRecord
Dim strField As Variant
Dim rField As cField
Dim CopyField As cField
Set rCopy = New cRecord
For Each strField In Fields
    Set rField = Fields(strField)
    Set CopyField = rField.NewField
    rCopy.Fields.Add CopyField.Name, CopyField
Next strField
Set NewRecord = rCopy
End Function

Public Function IsFieldValueBetween(ByVal fieldName As String, ByVal smallValue As Variant, ByVal largeValue As Variant) As Boolean
If IsFieldValue(fieldName, smallValue, ">=") Then
    If IsFieldValue(fieldName, largeValue, "<=") Then
        IsFieldValueBetween = True
    End If
End If
End Function

Public Function IsFieldValue(ByVal fieldName As String, ByVal FieldValue As Variant, Optional compareRule As Variant) As Boolean
Dim rField As cField
Dim strCompareRule As String
Dim varField As Variant
Dim varValue As Variant
IsFieldValue = False
If IsMissing(compareRule) Then
    strCompareRule = "="
Else
    strCompareRule = compareRule
End If
If Fields.Exists(fieldName) Then
    Set rField = Fields(fieldName)
    If IsDate(FieldValue) And IsDate(rField.Value) Then
        varField = CDate(rField.Value)
        varValue = CDate(FieldValue)
    ElseIf IsNumeric(FieldValue) And IsNumeric(rField.Value) Then
        varField = CDbl(rField.Value)
        varValue = CDbl(FieldValue)
    Else
        If IsNull(rField.Value) Then
            varField = Null
        Else
            varField = CStr(rField.Value)
        End If
        If IsNull(FieldValue) Then
            varValue = Null
        Else
            varValue = CStr(FieldValue)
        End If
        
    End If
    Select Case strCompareRule
    Case "="
        
        If IsNull(varValue) Then
            If IsNull(varField) Then IsFieldValue = True
        Else
            If varField = varValue Then IsFieldValue = True
        End If
    Case "ISNULL"
        If IsNull(varField) Then IsFieldValue = True
    Case "NOTNULL"
        If Not IsNull(varField) Then IsFieldValue = True
    Case Else
        
        If Not IsNull(varValue) And Not IsNull(varField) Then
            Select Case strCompareRule
            Case ">="
                If varField >= varValue Then IsFieldValue = True
            Case "<="
                If varField <= varValue Then IsFieldValue = True
            Case ">"
                If varField > varValue Then IsFieldValue = True
            Case "<"
                If varField < varValue Then IsFieldValue = True
            Case "<>"
                If varField <> varValue Then IsFieldValue = True
            End Select
        End If
        
    End Select
End If


End Function

Public Function Compare(ByVal rec As cRecord, ByVal structureOnly As Boolean) As Boolean
Dim varThis As Variant
Dim fldThis As cField
Compare = True
If Fields.Count = rec.Fields.Count Then
    For Each varThis In Fields.Items
        Set fldThis = varThis
        If rec.Fields.Exists(fldThis.Name) Then
            If rec.Fields(fldThis.Name).Order = fldThis.Order Then
                If Not structureOnly Then
                    If Not rec.Fields(fldThis.Name).Value = fldThis.Value Then
                        Compare = False
                        Exit Function
                    End If
                End If
            Else
                Compare = False
                Exit Function
            End If
        Else
            Compare = False
            Exit Function
        End If
    Next varThis
Else
    Compare = False
End If
End Function

Public Function GreaterOrSmallerThan(ByVal recCompare As cRecord, ByVal recSetting As cRecord) As Long
Const clngSmallest = -2 ^ 31 + 1
Dim recSortFields As cRecord
Dim recSortFieldsOrdered As cRecord
Dim varFld As Variant
Dim fldFld As cField
Dim lngSmallest As Long
Dim strSmallestName As String
Dim lngAccepted As Long
Set recSortFields = New cRecord
Set recSortFieldsOrdered = New cRecord
GreaterOrSmallerThan = 0
For Each varFld In recSetting.Fields
    Set fldFld = recSetting.Fields(varFld)
    If Not IsNull(fldFld.Value) And IsNumeric(fldFld.Value) Then
        recSortFields.AddField fldFld.Name, CLng(fldFld.Value)
    End If
Next varFld
lngSmallest = clngSmallest
lngAccepted = -1
Do Until recSortFieldsOrdered.Fields.Count = recSortFields.Fields.Count
    
    
    For Each varFld In recSortFields.Fields
        
        Set fldFld = recSortFields.Fields(varFld)
        If lngSmallest = clngSmallest And lngAccepted = -1 Then
            lngSmallest = fldFld.Value
            strSmallestName = fldFld.Name
        End If
        If lngSmallest = clngSmallest And lngAccepted > -1 Then
            If lngAccepted < Abs(fldFld.Value) Then
                lngSmallest = fldFld.Value
                strSmallestName = fldFld.Name
            End If
        End If
        If Abs(fldFld.Value) >= lngAccepted And Not recSortFieldsOrdered.Fields.Exists(fldFld.Name) Then
            If Abs(lngSmallest) > Abs(fldFld.Value) Then
                lngSmallest = fldFld.Value
                strSmallestName = fldFld.Name
            End If
        End If
    Next varFld
    
    If Not recSortFieldsOrdered.AddField(strSmallestName, lngSmallest) Is Nothing Then
        lngAccepted = Abs(lngSmallest)
        lngSmallest = clngSmallest
    End If
Loop
For Each varFld In recSortFieldsOrdered.Fields
    Set fldFld = recSortFieldsOrdered.Fields(varFld)
    If fldFld.Value > 0 Then
        If Me.FieldValue(fldFld.Name) > recCompare.FieldValue(fldFld.Name) Then
            GreaterOrSmallerThan = 1
        ElseIf Me.FieldValue(fldFld.Name) < recCompare.FieldValue(fldFld.Name) Then
            GreaterOrSmallerThan = -1
        Else
            GreaterOrSmallerThan = 0
        End If
        
    ElseIf fldFld.Value < 0 Then
        If Me.FieldValue(fldFld.Name) > recCompare.FieldValue(fldFld.Name) Then
            GreaterOrSmallerThan = -1
        ElseIf Me.FieldValue(fldFld.Name) < recCompare.FieldValue(fldFld.Name) Then
            GreaterOrSmallerThan = 1
        Else
            GreaterOrSmallerThan = 0
        End If
    Else
        GreaterOrSmallerThan = 0
    End If
    If GreaterOrSmallerThan <> 0 Then Exit For
Next varFld
End Function

Public Function ContainsSubsetOf(ByVal rec As cRecord) As Boolean
Dim varThis As Variant
Dim fldThis As cField
ContainsSubsetOf = True

    For Each varThis In Fields.Items
        Set fldThis = varThis
        If rec.Fields.Exists(fldThis.Name) Then
            If rec.Fields(fldThis.Name).Order = fldThis.Order Then
                If rec.Fields(fldThis.Name).Value = fldThis.Value Then
                
                Else
                    ContainsSubsetOf = False
                    Exit Function
                End If
            Else
                ContainsSubsetOf = False
                Exit Function
            End If
        Else
            ContainsSubsetOf = False
            Exit Function
        End If
    Next varThis

End Function


Public Function FieldOrderNumber(ByVal fieldName As String) As Long
Dim rField As cField
If Fields.Exists(fieldName) Then
    Set rField = Fields(fieldName)
    FieldOrderNumber = rField.Order
End If
End Function

Public Function FieldNameFromOrder(ByVal FieldOrder As Long) As String
Dim strField As Variant
Dim rField As cField
For Each strField In Fields
    Set rField = Fields(strField)
    If rField.Order = FieldOrder Then
        FieldNameFromOrder = rField.Name
        Exit Function
    End If
Next strField
End Function

Public Function FieldValue(ByVal fieldName As String) As Variant
Dim rField As cField
If Fields.Exists(fieldName) Then
    Set rField = Fields(fieldName)
    FieldValue = rField.Value
End If
End Function


Public Sub Dispose(Optional fldBlank As Variant)
Dim strField As Variant
Dim rField As cField
Dim lngCounter As Long
isDisposing = True
If IsMissing(fldBlank) Then
    For Each strField In Fields
        Set rField = Fields(strField)
        rField.Dispose
    Next strField
Else
    For Each strField In Fields
        Set Fields(strField) = fldBlank
    Next strField
End If
End Sub

Private Sub Class_Terminate()
Dispose
End Sub

Public Property Get XML() As MSXML2.IXMLDOMElement
Dim doc As MSXML2.DOMDocument60
Dim baseElement As MSXML2.IXMLDOMElement
Dim tempElement As MSXML2.IXMLDOMElement
Dim strField As Variant
Dim rField As cField
If Not isDisposing Then
    Set doc = New MSXML2.DOMDocument60
    Set baseElement = doc.createElement("cRecord")
    Set tempElement = doc.createElement("Fields")
    For Each strField In Fields
        Set rField = Fields(strField)
        tempElement.appendChild rField.XML
    Next strField
    baseElement.appendChild tempElement
    Set XML = baseElement
End If
End Property

Public Property Set XML(xmlIN As MSXML2.IXMLDOMElement)
Dim xmlFields As MSXML2.IXMLDOMNodeList
Dim tempElement As MSXML2.IXMLDOMElement
Dim tempField As cField
Set Fields = New Dictionary
If xmlIN.nodeName = "cRecord" Then
    Set xmlFields = xmlIN.getElementsByTagName("cField")
    For Each tempElement In xmlFields
        Set tempField = New cField
        Set tempField.XML = tempElement
        Fields.Add tempField.Name, tempField
    Next tempElement
End If
End Property
