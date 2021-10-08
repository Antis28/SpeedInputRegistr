VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Base"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Обложка которая будет добавлена в базу
Private EnterpriseElement As C_CoverInfo

' Класс для работы с КПР базой
Private workerWithXml As New C_WorkerWithBase
' Класс для работы с настройками
Private listOfDocNameXml As New C_WorkerWithSettings

Private CurrentIndex As Integer


Public Sub SaveKprBase(cover As C_CoverInfo)
    workerWithXml.Add cover
    workerWithXml.Save
End Sub

Public Sub UpdateBase()
    workerWithXml.Save
End Sub

Public Sub UpdateRecordInKprBase(cover As C_CoverInfo)
    workerWithXml.Update cover
    workerWithXml.Save
End Sub

Public Sub LoadAllBases()
    Call LoadCoverBase
    Call LoadNamesBase
End Sub

Public Function IndexExists(index As Integer) As Boolean
    If index > 0 And index <= workerWithXml.count Then
        IndexExists = True
        Exit Function
    End If
        IndexExists = False
End Function

Public Function CountInKprBase() As Integer
    CountInKprBase = workerWithXml.count
End Function

Public Function GetKprItem(index As Integer) As C_CoverInfo
    Set GetKprItem = workerWithXml.GetItem(index)
End Function

Public Sub SaveSettings()
    listOfDocNameXml.Save
End Sub

Public Function GetSettingsItem(index As Integer) As String
   GetSettingsItem = listOfDocNameXml.GetItem(index)
End Function

Public Function CountInSettings() As Integer
    CountInSettings = listOfDocNameXml.count
End Function

Public Sub RemoveRegister(index As Integer)
    workerWithXml.Remove (index)
End Sub

Public Sub LoadNamesBase()
    Set listOfDocNameXml = New C_WorkerWithSettings
    listOfDocNameXml.Init "DocumentNames"
End Sub

Public Sub LoadCoverBase()
    Set workerWithXml = New C_WorkerWithBase
    workerWithXml.Init "Base"
End Sub

Public Sub AddNewNameDoc(name As String)
    listOfDocNameXml.Add (name)
End Sub
