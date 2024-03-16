VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************
' Class Name:   cDataSet
' Description:  Used to store a collection of tables of data in memory.
'               Ability to Load and Save from XML
'
' Created by:   Brendon Raw
' History:      v1.00 05 January 2012 - Created

' References:   cDataTable v1.11
'               Microsoft XML v6
'               Microsoft Scripting Runtime
'
' Instructions: (AddDataTable, DeleteDataTable) Add or Remove a table from the DataSet
'               (XML) returns/loads the object as an XML object
'               (LoadXML, SaveAsXML) loads and saves object in XML file
'*************************************************************************
Const newTableName = "cDataTable"
Public DataTables As Dictionary
Public Application As String
Public Author As String
Public Version As String
Private isDisposing As Boolean

Private Sub Class_Initialize()
Set DataTables = New Dictionary
End Sub

Public Function AddDataTable(dat As cDataTable, blnOverWriteIfExists As Boolean) As Boolean
Dim lngCount As Long
lngCount = 1
If IsNull(dat) Then
    AddDataTable = False
ElseIf dat.TableName = "" Then
    Do Until Not DataTables.Exists(newTableName & CStr(lngCount))
        lngCount = lngCount + 1
    Loop
    dat.TableName = newTableName & CStr(lngCount)
    DataTables.Add dat.TableName, dat
    AddDataTable = True
ElseIf DataTables.Exists(dat.TableName) Then
    If blnOverWriteIfExists Then
        If DeleteDataTable(dat.TableName) Then
            DataTables.Add dat.TableName, dat
            AddDataTable = True
        Else
            AddDataTable = False
        End If
    Else
        AddDataTable = False
    End If
Else
    DataTables.Add dat.TableName, dat
    AddDataTable = True
End If
End Function

Public Function DeleteDataTable(strName As String) As Boolean
If DataTables.Exists(strName) Then
    DataTables.Remove strName
    DeleteDataTable = True
Else
    DeleteDataTable = False
End If
End Function

Public Property Get XML() As MSXML2.IXMLDOMElement
Dim strTab As Variant
Dim tbl As cDataTable
Dim doc As MSXML2.DOMDocument60
Dim baseElement As MSXML2.IXMLDOMElement
Dim tempElement As MSXML2.IXMLDOMElement
Dim applicationElement As MSXML2.IXMLDOMElement
Dim authorElement As MSXML2.IXMLDOMElement
Dim versionElement As MSXML2.IXMLDOMElement
If Not isDisposing Then
    Set doc = New MSXML2.DOMDocument60
    Set baseElement = doc.createElement("cDataSet")
    Set applicationElement = doc.createElement("Application")
    applicationElement.nodeTypedValue = Application
    Set authorElement = doc.createElement("Author")
    authorElement.nodeTypedValue = Author
    Set versionElement = doc.createElement("Version")
    versionElement.nodeTypedValue = Version
    Set tempElement = doc.createElement("Tables")
    For Each strTab In DataTables
        Set tbl = DataTables(strTab)
        tempElement.appendChild tbl.XML
    Next strTab
    baseElement.appendChild applicationElement
    baseElement.appendChild authorElement
    baseElement.appendChild versionElement
    baseElement.appendChild tempElement
    Set XML = baseElement
End If
End Property

Public Property Set XML(xmlIN As MSXML2.IXMLDOMElement)
Dim xmlTables As MSXML2.IXMLDOMNodeList
Dim xmlApplication As MSXML2.IXMLDOMNodeList
Dim xmlAuthor As MSXML2.IXMLDOMNodeList
Dim xmlVersion As MSXML2.IXMLDOMNodeList
Dim tempElement As MSXML2.IXMLDOMElement
Dim tempTable As cDataTable
Set DataTables = New Dictionary
If xmlIN.nodeName = "cDataSet" Then
    Set xmlApplication = xmlIN.getElementsByTagName("Application")
    Application = xmlApplication.Item(0).nodeTypedValue
    Set xmlAuthor = xmlIN.getElementsByTagName("Author")
    Author = xmlAuthor.Item(0).nodeTypedValue
    Set xmlVersion = xmlIN.getElementsByTagName("Version")
    Version = xmlVersion.Item(0).nodeTypedValue
    Set xmlTables = xmlIN.getElementsByTagName("cDataSet")
    For Each tempElement In xmlTables
        Set tempTable = New cDataset
        Set tempTable.XML = tempElement
        AddDataTable tempTable, True
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
Set DataTables = New Dictionary
Set tempNode = doc.getElementsByTagName("cDataSet")
Set tempElement = tempNode.Item(0)
Set Me.XML = tempElement
LoadXML = True
End Function


