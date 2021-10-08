VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_WorkerWithSettings"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Загруженый XML файла
Private Settings As Collection

' Имя XML файла
Private fileName As String
' Путь для сохранения итогового XML файла
Private fullPath As String

Public Property Get count() As Variant
    count = Settings.count
End Property

Public Function GetItem(index) As String
    GetItem = Settings(index)
End Function


Public Sub Add(docName As String)
    Settings.Add docName
End Sub

Public Sub Save()
    Dim saver As New C_SaverSettings
    saver.Init fileName, fullPath
    saver.Save Settings
End Sub

Public Sub Init(name As String)
    fileName = name
    InitFilePath
    LoadBase
End Sub

Public Function GetList() As Collection
   Set GetList = Settings
End Function

Private Function InitFilePath()
    Dim fileNameAndExtention As String
    fileNameAndExtention = fileName & ".xml"
    
    fullPath = TemplateProject.ThisDocument.path & "\" & fileNameAndExtention
End Function

Private Sub LoadBase()
    If FileExists Then
        Dim constructor As New C_LoaderSettings
        Set Settings = constructor.Construct(fullPath)
        
        'Сортироавка
        Dim ArrSort As New ArraySorter
        Dim arr As Variant
    
        arr = CollectionToArray(Settings)
        ArrSort.SortArrByAlpavit arr
        Set Settings = ArrayToCollection(arr)
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
    Set Settings = New Collection
    Save
End Sub
