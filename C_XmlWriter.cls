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

' ������ �������� ������ � ������ XML_Worker

Private xmlDocument As New DOMDocument
Private xmlRoot As IXMLDOMNode


' ���� ��� ���������� ��������� XML �����
Private fullPath As String

'Private attributesNames As New XML_Attributes

' ��������� XML ������
Public Sub Init(rootname As String, path As String)
    fullPath = path

    'C������� ������� XML
    'Set xmlDocument = CreateObject("MSXML2.DOMDocument")
    
    '���������� �������� XML, �������� ��� ���������� �������!!!
    'xmlDocument.appendChild xmlDocument.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    'xmlDocument.appendChild xmlDocument.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'")
       
    '�������� ��������� ��������
    Set xmlRoot = xmlDocument.createElement(rootname)
    
    '���������� ��������� �������� � xml
    xmlDocument.appendChild xmlRoot
End Sub

Public Sub Add(node As IXMLDOMNode)
    xmlRoot.appendChild node
End Sub

Public Function createElement(name As String) As IXMLDOMNode
   Set createElement = xmlDocument.createElement(name)
End Function

Public Sub SaveToXmlFile()
    '���������� XSL-�������������� ��� ���������� �������� � XML
    ''UTF-8'
    'NormalizeToReadable "UTF-8"
    NormalizeToReadable "windows-1251"
    
    '���������� ����� (��� ������ ���� ����������, ������ ��� �������)
    'xmlDocument.Save fullPath
    RecordForwardInFile
    
 
    '���������� ����� (� ������� ���� ����������)
    '    xml.Save Application.GetSaveAsFilename("", "���� �������� (*.xml),", , "������� ��� �����", "���������")
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
    
    '�������� ��������
    Set newNode = xmlDocument.createElement(text)

    Set CreateNode = newNode
End Function

Private Function CreateAttribute(node As Variant, name As String, text As String) As Variant
    '���������� � �������� ��������
    node.setAttribute name, text

    Set CreateAttribute = node
End Function

'��������� ��� �������� XML ��������� ���� (� ���������)
Private Sub NormalizeToReadable(encoding As String)
    Dim xsl As Variant
    'C������� ������� XSL
    Set xsl = CreateObject("MSXML2.DOMDocument")

    '�������� XSL �� ������ (�� ������� ������� ���������� XSL-�����)
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

    '���������� ��������������
    xmlDocument.transformNodeToObject xsl, xmlDocument
End Sub



'������ ������ ������ � ����(������������ ��� xml.save)
Private Sub RecordForwardInFile()
    ' ������ ������ ������ � ����
    Open fullPath For Output As #1
        Print #1, xmlDocument.XML
    Close #1
End Sub
