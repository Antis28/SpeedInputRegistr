VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SaverCover"
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
Private numberInBase As String

Private lastChange As String
Private index As String

Private documentElement As String
Private DateFirst As String
Private dateLast As String
Private sheetNumbers As String
Private documentNumber As String

Private saver As New C_XmlWriter

Public Sub Init(rootname As String, path As String)
    saver.Init rootname, path
End Sub

' формирует список объектов
Public Function Save(data As Collection)
    
    Dim item As C_CoverInfo
    Dim node As IXMLDOMNode
    
    For Each item In data
        'Вызов функции добавления
        Set node = createCoverElement(item)
        saver.Add node
    Next
    
    saver.SaveToXmlFile
End Function

'Формирует элемент по объекту, возвращает добавленный элемент XML
Private Function createCoverElement(enterprise As C_CoverInfo) As Variant
    Dim Element As IXMLDOMElement
    
    'Создание элемента
    Set Element = saver.createElement(EnterpriseElement)
    Element.setAttribute index, enterprise.index
    Element.setAttribute okpo, enterprise.OkpoEnterprise
    Element.setAttribute name, enterprise.NameEnterprise
    Element.setAttribute years, enterprise.years
    Element.setAttribute sheetCount, enterprise.sheetCount
    
    Element.setAttribute numberInBase, enterprise.numberInBase
    
    Element.setAttribute lastChange, enterprise.lastChange
    

    Dim ItemRecord As C_RecordInfo
    If (enterprise.innerRegistry.count > 0) Then
        For Each ItemRecord In enterprise.innerRegistry
            Element.appendChild createRegistrElement(ItemRecord)
        Next
    End If
    Set createCoverElement = Element
End Function

'Формирует элемент по объекту, возвращает добавленный элемент XML
Private Function createRegistrElement(record As C_RecordInfo) As Variant
    Dim Element As IXMLDOMElement
    
    'Создание элемента
    Set Element = saver.createElement(documentElement)
    Element.setAttribute name, record.docName
    Element.setAttribute sheetNumbers, record.sheetsNumber
    Element.setAttribute DateFirst, record.DateFirst
    Element.setAttribute dateLast, record.dateLast
    Element.setAttribute sheetCount, record.SheetsCount
    Element.setAttribute documentNumber, record.docNumber
    
      
    Set createRegistrElement = Element
End Function

Sub Class_Initialize()
'Присвоение названий атрибутов:
    EnterpriseElement = "item"
    okpo = "okpo"
    name = "name"
    years = "years"
    sheetCount = "sheetCount"
    numberInBase = "numberInBase"
    lastChange = "lastChange"
    index = "id"

    documentElement = "doc"
    DateFirst = "DateFirst"
    dateLast = "DateLast"
    sheetNumbers = "sheetNumbers"
    documentNumber = "docNumber"
End Sub
