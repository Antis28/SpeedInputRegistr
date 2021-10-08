Attribute VB_Name = "Save"
Option Explicit


Sub SaveAccompanying(nameDoc As String)
    On Error Resume Next
    Dim fileName As String
    Dim newFolderPath As String
    
     
    ' путь для сохранения
    ' для шаблона
    newFolderPath = TemplateProject.ThisDocument.path
    
    ' Для документа
    'newFolderPath = ActiveDocument.Path
    newFolderPath = newFolderPath
    ' создать папку для документа
    MkDir newFolderPath
    
    ' путь для файла для документа
    If Len(nameDoc) > 200 Then
       fileName = newFolderPath & "\" & left(nameDoc, 20)
    Else
       fileName = newFolderPath & "\" & nameDoc
    End If
    
    
    
    fileName = fileName & ".doc"
    
    Dim spaceCharacter As Variant
    
    'spaceCharacter = Chr(0)
    spaceCharacter = ""
    
    fileName = Replace(fileName, Chr(34), spaceCharacter)
    fileName = Replace(fileName, Chr(60), spaceCharacter)
    fileName = Replace(fileName, Chr(62), Chr(spaceCharacter))
    fileName = Replace(fileName, Chr(147), Chr(spaceCharacter))
    fileName = Replace(fileName, Chr(148), Chr(spaceCharacter))
    fileName = Replace(fileName, Chr(171), Chr(spaceCharacter))
    fileName = Replace(fileName, Chr(187), Chr(spaceCharacter))
       
    
    
    ActiveDocument.SaveAs fileName:=fileName
End Sub
