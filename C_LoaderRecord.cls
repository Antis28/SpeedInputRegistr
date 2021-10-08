VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_LoaderRecord"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Должен вызываться только из класса C_ConstructorCover

Private documentElement As String

Private name As String
Private DateFirst As String
Private dateLast As String
Private sheetNumbers As String
Private sheetCount As String

Private documentNumber As String

' формирует список объектов
Public Function Construct(node As IXMLDOMNode) As Collection
    Dim nodeList As IXMLDOMNodeList
    Set nodeList = node.ChildNodes
   
    Dim recordList As New Collection
    Dim record As C_RecordInfo
    Dim i As Integer
    For i = 0 To nodeList.Length - 1
        Dim nodeRec As IXMLDOMNode
        Set nodeRec = nodeList.item(i)
        Set record = ConstructItem(nodeRec)
        recordList.Add record
    Next i

    Set Construct = recordList
End Function

' формирует объект из Xml текста
Private Function ConstructItem(node As IXMLDOMNode) As C_RecordInfo
        
    Dim record As New C_RecordInfo
    Dim attr As IXMLDOMNamedNodeMap
    Set attr = node.attributes
    
    record.docName = attr.getNamedItem(name).NodeValue
    record.sheetsNumber = attr.getNamedItem(sheetNumbers).NodeValue
    record.DateFirst = attr.getNamedItem(DateFirst).NodeValue
    record.dateLast = attr.getNamedItem(dateLast).NodeValue
    
    Dim xAttr As Variant
    Set xAttr = attr.getNamedItem(sheetCount)
    If Not (xAttr Is Nothing) Then
        record.SheetsCount = xAttr.NodeValue
    End If
    
    Set xAttr = attr.getNamedItem(documentNumber)
    If Not (xAttr Is Nothing) Then
        record.docNumber = xAttr.NodeValue
    End If
    
    
    
    Set ConstructItem = record
End Function

Private Sub Class_Initialize()
'Присвоение названий для атрибутов:
    name = "name"
    documentElement = "doc"
    DateFirst = "DateFirst"
    dateLast = "DateLast"
    sheetNumbers = "sheetNumbers"
    sheetCount = "sheetCount"
    documentNumber = "docNumber"
End Sub
