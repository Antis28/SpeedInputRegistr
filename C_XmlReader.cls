VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_XmlReader"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
' �������� ������ �� ������ ������ �� XML

Private fullPath As String

' �������� ������� ���������
Private xmlRoot As IXMLDOMNode

' ���� XML ��������
Private xmlDocument As New DOMDocument


Public Sub Init(path As String)
    fullPath = path
End Sub

Public Function GetElements(name As String) As IXMLDOMNodeList
    ' ������ XML ���� � ������
    Dim xmlString As String
    
    xmlString = ReadFile()

    ' ��������� �� ������ � XML ������
    xmlDocument.LoadXML xmlString

    ' ��������� ������ ��������� ��������� ��������
    Set GetElements = xmlDocument.documentElement.getElementsByTagName(name)
End Function

' ��������� ������ ����� �� ���������
Public Function ReadAttributes(node As IXMLDOMNode, nameAttribute As String) As Collection
    Dim attributes As IXMLDOMNamedNodeMap
    Dim cureattribute As IXMLDOMNode
    Dim listValues As New Collection
    
    Set attributes = node.attributes
    
    For Each cureattribute In attributes
        listValues.Add cureattribute.text
    Next
    
    Set ReadAttributes = listValues
End Function


Private Function ReadFile() As String
    Dim currentLine As String
    Dim AllText As String
    
    
    Open fullPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, currentLine
        AllText = AllText & currentLine
    Loop
    Close #1
    ReadFile = AllText
End Function
