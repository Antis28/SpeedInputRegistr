VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_WorkerWithBase"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Загруженый XML файла
Private KprBase As Collection

' Имя XML файла
Private fileName As String
' Путь для сохранения итогового XML файла
Private fullPath As String

Public Property Get count() As Variant
    count = KprBase.count
End Property

Public Function GetItem(index As Integer) As C_CoverInfo
   Set GetItem = KprBase(index)
End Function

Public Sub Add(cover As C_CoverInfo)
    If KprBase.count > 0 Then
        Dim previous As C_CoverInfo
        Set previous = KprBase(KprBase.count)
    
        'cover.index = previous.index + 1
        cover.index = KprBase.count + 1
    Else
        cover.index = 1
    End If
    
    KprBase.Add cover
End Sub

Public Sub Remove(index As Integer)
    If index = 0 Or index > KprBase.count Then
        Exit Sub
    End If
    Call KprBase.Remove(index)
    
    ' обновить индексацию
    Dim item As C_CoverInfo
    Dim id As Integer
    id = 1
    For Each item In KprBase
        item.index = id
        id = id + 1
    Next
End Sub

' Перезаписывает данные об обложке
Public Sub Update(cover As C_CoverInfo)
    Dim item As C_CoverInfo
    Set item = KprBase(cover.index)
    
    item.NameEnterprise = cover.NameEnterprise
    item.OkpoEnterprise = cover.OkpoEnterprise
    item.sheetCount = cover.sheetCount
    item.years = cover.years
    Set item.innerRegistry = cover.innerRegistry
    item.lastChange = cover.lastChange
End Sub

Public Sub Save()
    Dim saver As New C_SaverCover
    saver.Init fileName, fullPath
    saver.Save KprBase
End Sub

Public Sub Init(name As String)
    fileName = name
    InitFilePath
    
'    saver.InitXML name, fullPath
'    saverDocName.InitXML "docNames", fullPath
    LoadBase
End Sub

Private Function InitFilePath()
    Dim fileNameAndExtention As String
    fileNameAndExtention = fileName & ".xml"
    
    'Путь для сохранения итогового XML
    'fullPath = ActiveDocument.Path & "\" & FileNameAndExtention
    'fullPath = TemplateProject.ThisDocument.Path & "\" & FileNameAndExtention
    fullPath = TemplateProject.ThisDocument.path & "\" & fileNameAndExtention
End Function

Private Sub LoadBase()
    If FileExists Then
        Dim constructor As New C_LoaderCover
        Set KprBase = constructor.Construct(fullPath)
        
    Else
        CreateEmptyBase
    End If
End Sub

Private Function FileExists() As Boolean
    Dim result As Boolean
    If Not Dir(fullPath) = vbNullString Then
        result = True
    Else
        result = False
    End If
    
    FileExists = result
End Function

Private Sub CreateEmptyBase()
    Set KprBase = New Collection
    Save
End Sub
