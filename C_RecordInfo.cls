VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_RecordInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Public id As Long

Public DateFirst As Integer
Public dateLast As Integer
Public docName As String
Public docNumber As String
Public sheetsNumber As String
Public SheetsCount As Integer

Public Sub Construct(idVal As Integer, _
                        DateFirstVal As Integer, _
                        DateLastVal As Integer, _
                        DocNameVal As String, _
                        docNumberVal As String, _
                        SheetsNumberVal As String, _
                        SheetsCountVal As Integer)
    
   
    
    id = idVal
    
    DateFirst = DateFirstVal
    dateLast = DateLastVal
    docName = DocNameVal
    docNumber = docNumberVal
    sheetsNumber = SheetsNumberVal
    SheetsCount = SheetsCountVal
End Sub
