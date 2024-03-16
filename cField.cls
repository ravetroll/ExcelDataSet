VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' Class Name:  cField
' Description:  Used to store a field of data in memory.
'               Ability to represent in XML
'               No data typing required
' Created by:   Brendon Raw
' History:      v1.00 04 October 2011
'               v1.01 24 October 2011 added Dispose
'               v1.02 08 February 2012 Changed Value to a Property
' References:   Microsoft XML v6
'               Microsoft Scripting Runtime
'
' Instructions: (NewField) Creates a new empty field
'               (Copy) returns a copy of this instance
'               (XML) returns/loads the object as an XML object
'               (Dispose) Disables propertys and functions as object is being de-allocated
'*************************************************************************
Public Name As String
Private varValue As Variant
Public Order As Long
Private isDisposing As Boolean

Public Property Get Value() As Variant
Value = varValue
End Property

Public Property Let Value(varIn)
If IsEmpty(varIn) Then
    varValue = Null
Else
    varValue = varIn
End If
End Property


Public Function Copy() As cField
Dim fCopy As cField
If Not isDisposing Then
    Set fCopy = New cField
    fCopy.Name = Name
    fCopy.Value = Value
    fCopy.Order = Order
    Set Copy = fCopy
End If
End Function

Public Function NewField() As cField
Dim fCopy As cField
If Not isDisposing Then
    Set fCopy = New cField
    fCopy.Name = Name
    fCopy.Value = Null
    fCopy.Order = Order
    Set NewField = fCopy
End If
End Function

Public Property Get XML() As MSXML2.IXMLDOMElement
Dim doc As MSXML2.DOMDocument60
Dim baseElement As MSXML2.IXMLDOMElement
Dim tempElement As MSXML2.IXMLDOMElement
If Not isDisposing Then
    Set doc = New MSXML2.DOMDocument60
    Set baseElement = doc.createElement("cField")
    Set tempElement = doc.createElement("Name")
    tempElement.nodeTypedValue = Name
    baseElement.appendChild tempElement
    Set tempElement = doc.createElement("Value")
    tempElement.nodeTypedValue = Value & ""
    baseElement.appendChild tempElement
    Set tempElement = doc.createElement("Order")
    tempElement.nodeTypedValue = Order
    baseElement.appendChild tempElement
    Set XML = baseElement
End If
End Property

Public Property Set XML(xmlIN As MSXML2.IXMLDOMElement)
Dim xmlFields As MSXML2.IXMLDOMNodeList
Dim tempElement As MSXML2.IXMLDOMElement
If Not isDisposing Then
    Name = ""
    Value = ""
    Order = 0
    If xmlIN.nodeName = "cField" Then
        Set xmlFields = xmlIN.childNodes
        For Each tempElement In xmlFields
            If tempElement.nodeName = "Name" Then
                Name = tempElement.nodeTypedValue
            ElseIf tempElement.nodeName = "Value" Then
                Value = tempElement.nodeTypedValue
            ElseIf tempElement.nodeName = "Order" Then
                Order = tempElement.nodeTypedValue
            Else
            End If
        Next tempElement
    End If
End If
End Property

Public Function Dispose()
isDisposing = True
End Function


