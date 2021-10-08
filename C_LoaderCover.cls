VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_LoaderCover"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Должен вызываться только из класса  C_WorkerWithBase

Private EnterpriseElement As String
Private okpo As String
Private name As String
Private years As String
Private sheetCount As String
Private index As String

Private lastChange As String

' формирует список объектов
Public Function Construct(path As String) As Collection
    Dim reader As New C_XmlReader
    Dim nodeList As IXMLDOMNodeList

    reader.Init path
    Set nodeList = reader.GetElements(EnterpriseElement)

    Dim coverList As New Collection
    Dim i As Integer
    Dim node As IXMLDOMNode
    Dim cover As C_CoverInfo
    For i = 0 To nodeList.Length - 1
        Set node = nodeList.item(i)
        Set cover = ConstructItem(node, reader)
        coverList.Add cover
    Next i
        
    Set Construct = coverList
End Function

Private Function ConstructItem(node As IXMLDOMNode, reader As C_XmlReader) As C_CoverInfo
    Dim cover As New C_CoverInfo
    Dim attr As IXMLDOMNamedNodeMap
    Set attr = node.attributes
    
    cover.NameEnterprise = GetAndCheckNodeValue(attr, name)
    
    cover.OkpoEnterprise = GetAndCheckNodeValue(attr, okpo)
    
    cover.sheetCount = GetAndCheckNodeValue(attr, sheetCount)
    
    cover.years = GetAndCheckNodeValue(attr, years)
    
    cover.index = GetAndCheckNodeValue(attr, index)
    
    
    cover.lastChange = GetAndCheckNodeValue(attr, lastChange)
    
    ' Старое название аттрибута
    If cover.lastChange = "" Then
        cover.lastChange = GetAndCheckNodeValue(attr, "createdDate")
    End If

    Dim recordConstr As New C_LoaderRecord
    Set cover.innerRegistry = recordConstr.Construct(node)

    Set ConstructItem = cover
End Function



Sub Class_Initialize()
'Присвоение названий атрибутов:
    EnterpriseElement = "item"
    okpo = "okpo"
    name = "name"
    years = "years"
    sheetCount = "sheetCount"
    index = "id"
    
    lastChange = "lastChange"
End Sub


Public Function GetAndCheckNodeValue(attr As IXMLDOMNamedNodeMap, NamedItem As String) As String
    Dim attrItem As Variant
    Set attrItem = attr.getNamedItem(NamedItem)
    
    If Not (attrItem Is Nothing) Then
        GetAndCheckNodeValue = attrItem.NodeValue
    End If
End Function
