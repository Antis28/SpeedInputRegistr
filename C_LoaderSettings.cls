VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_LoaderSettings"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Должен вызываться только из класса  C_Worker___

Private Element As String

' формирует список объектов
Public Function Construct(path As String) As Collection
    Dim reader As New C_XmlReader
    Dim nodeList As IXMLDOMNodeList

    reader.Init path
    Set nodeList = reader.GetElements(Element)

    
    Dim i As Integer
    Dim node As IXMLDOMNode
    Dim settingsList As New Collection

    For i = 0 To nodeList.Length - 1
        Set node = nodeList.item(i)
        settingsList.Add node.text
    Next i
        
    Set Construct = settingsList
End Function


Sub Class_Initialize()
'Присвоение названий атрибутов:
    Element = "item"
End Sub
