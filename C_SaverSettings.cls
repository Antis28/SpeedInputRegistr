VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SaverSettings"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Должен вызываться только из класса  C_WorkerWith____

Private elementName As String

Private saver As New C_XmlWriter

Public Sub Init(rootname As String, path As String)
    saver.Init rootname, path
End Sub

' формирует список объектов
Public Function Save(data As Collection)
    
    Dim item As Variant
    
    For Each item In data
        Dim Element As IXMLDOMElement
        'Создание элемента
        Set Element = saver.createElement(elementName)
        Element.text = item
        saver.Add Element
    Next
    
    saver.SaveToXmlFile
End Function


Sub Class_Initialize()
'Присвоение названий атрибутов:
    elementName = "item"
End Sub
