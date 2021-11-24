VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_CoverInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public NameEnterprise As String
Public OkpoEnterprise As String
Public years As String
Public sheetCount As String
Public index As Integer
Public lastChange As String
Public numberInBase As String
Public innerRegistry As New Collection



' Не используется?
Public Sub Construct(indx As Integer, name As String, _
                    okpo As String, _
                    numberInBaseVal As String, _
                    year As String, _
                    sheets As String, _
                    innerReg As Collection _
                    )
    NameEnterprise = name
    OkpoEnterprise = okpo
    years = year
    sheetCount = sheets
    index = indx
    Set innerRegistry = innerReg
    numberInBase
    'lastChange = DatePart("dd.mm.yyyy", Now)
End Sub
