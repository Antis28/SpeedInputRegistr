VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_XmlWriter"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Должен работать только в классе XML_Worker

Private xmlDocument As New DOMDocument
Private xmlRoot As IXMLDOMNode


' Путь для сохранения итогового XML файла
Private fullPath As String

'Private attributesNames As New XML_Attributes

' формирует XML объект
Public Sub Init(rootname As String, path As String)
    fullPath = path

    'Cоздание объекта XML
    'Set xmlDocument = CreateObject("MSXML2.DOMDocument")
    
    'Добавление описания XML, меняется при добавлении отсупов!!!
    'xmlDocument.appendChild xmlDocument.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    'xmlDocument.appendChild xmlDocument.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'")
       
    'Создание корневого элемента
    Set xmlRoot = xmlDocument.createElement(rootname)
    
    'Добавление корневого элемента в xml
    xmlDocument.appendChild xmlRoot
End Sub

Public Sub Add(node As IXMLDOMNode)
    xmlRoot.appendChild node
End Sub

Public Function createElement(name As String) As IXMLDOMNode
   Set createElement = xmlDocument.createElement(name)
End Function

Public Sub SaveToXmlFile()
    'Выполнение XSL-преобразования для добавления отступов в XML
    ''UTF-8'
    'NormalizeToReadable "UTF-8"
    NormalizeToReadable "windows-1251"
    
    'Сохранение файла (без выбора пути сохранения, удобно при отладке)
    'xmlDocument.Save fullPath
    RecordForwardInFile
    
 
    'Сохранение файла (с выбором пути сохранения)
    '    xml.Save Application.GetSaveAsFilename("", "Файл экспорта (*.xml),", , "Введите имя файла", "Сохранить")
    '    xml.Save GetSaveAsFilename
End Sub

Private Function GetSaveAsFilename() As String
    Const msoFileDialogSaveAs = 2 '-- without referencing to Microsoft Office xx.x library
    
    With Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
        .InitialFileName = ActiveDocument.path & "\abc.xml" '-- optional
        .FilterIndex = 15
        If .Show = True Then
            GetSaveAsFilename = .SelectedItems(1)
        'Else
        '    '-- return empty string
        '    MsgBox "SaveAs process was cancelled"
        End If
    End With
End Function

Private Function CreateNode(ByRef text As String) As Variant
    Dim newNode As Variant
    
    'Создание элемента
    Set newNode = xmlDocument.createElement(text)

    Set CreateNode = newNode
End Function

Private Function CreateAttribute(node As Variant, name As String, text As String) As Variant
    'Добавление к элементу атрибута
    node.setAttribute name, text

    Set CreateAttribute = node
End Function

'Процедура для придания XML читаемого вида (с отступами)
Private Sub NormalizeToReadable(encoding As String)
    Dim xsl As Variant
    'Cоздание объекта XSL
    Set xsl = CreateObject("MSXML2.DOMDocument")

    'Загрузка XSL из строки (не требует наличия отдельного XSL-файла)
    xsl.LoadXML ( _
        "<xsl:stylesheet version='1.0' xmlns:xsl='http://www.w3.org/1999/XSL/Transform'>" & vbCrLf & _
        "<xsl:output method='xml' version='1.0' encoding= '" & _
        encoding & _
        "' indent='yes'/>" & vbCrLf & _
        "<xsl:template match='@*|node()'>" & vbCrLf & _
        "<xsl:copy>" & vbCrLf & _
        "<xsl:apply-templates select='@*|node()' />" & vbCrLf & _
        "</xsl:copy>" & vbCrLf & _
        "</xsl:template>" & vbCrLf & _
        "</xsl:stylesheet>")

    'Выполнение преобразования
    xmlDocument.transformNodeToObject xsl, xmlDocument
End Sub



'Прямая запись текста в файл(альтернатива для xml.save)
Private Sub RecordForwardInFile()
    ' Прямая запись строки в файл
    Open fullPath For Output As #1
        Print #1, xmlDocument.XML
    Close #1
End Sub
