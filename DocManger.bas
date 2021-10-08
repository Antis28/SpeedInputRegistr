Attribute VB_Name = "DocManger"
Option Explicit

Public myDocs() As Document

' Создать документ по шаблону
Public Sub CreateNewDocment()
    Dim templatePath As String
    
    templatePath = "C:\Documents and Settings\PENSION\Рабочий стол\Обложки для архива КПР\#ОПИСЬ ВНУТРЕННЯЯ#_v_14_08_2021_1.dot"
    If Dir(templatePath) = "" Then
        templatePath = "C:\Users\Antis\Desktop\1111111111\#ОПИСЬ ВНУТРЕННЯЯ#_v_14_08_2021_1.dot"
    End If
    
    Documents.Add templatePath
End Sub

' Создать документ обложки по шаблону
Public Sub CreateNewCoverDoc()
    Dim templatePath As String
    Dim templateName As String
    Dim fullPath As String
    
    templateName = "#ОБЛОЖКА#.dot"
    
    templatePath = "C:\Documents and Settings\PENSION\Рабочий стол\Обложки для архива КПР\"
    If Dir(templatePath) = "" Then
        templatePath = "C:\Users\Antis\Desktop\1111111111\"
    End If
    
    fullPath = templatePath & templateName
    Documents.Add fullPath
End Sub

' Создать документ обложки по шаблону
Public Sub CreateNewRegisterDoc()
    Dim templatePath As String
    Dim templateName As String
    Dim fullPath As String
    
    templateName = "#ОПИСЬ ВНУТРЕННЯЯ#.dot"
    
    templatePath = "C:\Documents and Settings\PENSION\Рабочий стол\Обложки для архива КПР\"
    If Dir(templatePath) = "" Then
        templatePath = "C:\Users\Antis\Desktop\1111111111\"
    End If
    
    fullPath = templatePath & templateName
    Documents.Add fullPath
End Sub


Public Sub FillOneDocument(curRegister As C_RegisterInfo)
       
    Dim cover As C_CoverInfo
    Set cover = PrepareCover(curRegister)
    
    ' Сохранить в базу если стоит флаг или если индекс = 0
    If Form_Register.cb_SaveInBase.value Or cover.index = 0 Then
        myBase.SaveKprBase cover
    End If
    
    ' Индекс для названия файла
    Dim nameIndex As String
    nameIndex = TestNameZero(cover.index)
    
    Call FillOneCoverDoc(nameIndex, cover)
    Call FillOneRegisterDoc(nameIndex, cover)
    
End Sub

Public Sub FillDocument(registers As Collection)
            
    Dim item As C_RegisterInfo
    For Each item In registers
        Call Form_Register.FixPageNumbers(1, item)
        FillOneDocument item
    Next
    
End Sub

Public Function DivRegister() As Collection
    Dim registers As New Collection
    Dim numRegister As Integer
    Dim i As Integer
    
    Dim curRegister As New C_RegisterInfo
    
    Dim curRecord As C_RecordInfo
    
    Dim sheets(1 To 6) As Integer
    
    Dim SheetsCount As Integer
        
    
    registers.Add New C_RegisterInfo
   
    For i = 1 To register.count
        numRegister = 1
        Set curRecord = register.getRecord(i)
        Set curRegister = registers(1)
        
        SheetsCount = sheets(numRegister) + curRecord.SheetsCount
        
        ' если листов менее 250, то добавляем в опись под номером numRegister
        If (SheetsCount < 251) Then
            sheets(numRegister) = SheetsCount
        End If
        
        'если количество листов больше 250, ищем свободную опись или создаем новую опись
        Do While (SheetsCount > 250)
            numRegister = numRegister + 1
            
            ' нет свободных описей, создать новую
            If registers.count < numRegister Then
                registers.Add New C_RegisterInfo
            End If
            
            If numRegister > 6 Then
                MsgBox "количество томов больше 6-ти!!!"
                GoTo endMark
            End If
            ' вычислить сумму листов описи с номером numRegister и текущей записи
            SheetsCount = sheets(numRegister) + curRecord.SheetsCount
        Loop
        sheets(numRegister) = SheetsCount
        
        Set curRegister = registers(numRegister)
        curRegister.Add register.getRecord(i)
    Next i
endMark:
    Set DivRegister = registers
End Function

Private Function PrepareCover(curRegister As C_RegisterInfo) As C_CoverInfo
    Dim cover As C_CoverInfo
    Set cover = CreateCoverWithOutID(curRegister)
    
    Dim newIndex As Integer
    newIndex = Form_Register.l_indexInBase.Caption
    
    If newIndex > 0 Then
        cover.index = newIndex
    End If
    
    Set PrepareCover = cover
End Function

Private Sub FillOneCoverDoc(nameIndex As String, cover As C_CoverInfo)
    ' Создать новый файл по шаблону с обложкой
    CreateNewCoverDoc
    
    ' Заполнить документ с обложкой
    FillCoverDoc cover
    Dim fileNameWithCover As String
    fileNameWithCover = nameIndex & "_C_" & cover.sheetCount & "л_" & cover.NameEnterprise
    
    ' Сохранить  word документ
    SaveAccompanying "ОБЛОЖКИ\" & fileNameWithCover
    ActiveDocument.Close
End Sub

Private Sub FillOneRegisterDoc(nameIndex As String, cover As C_CoverInfo)
    ' Создать новый файл по шаблону с описью
    CreateNewRegisterDoc
    
    ' Заполнить документ внутреннюю опись
    FillRegister cover
    Dim fileNameWithRegister As String
    fileNameWithRegister = nameIndex & "_R_" & cover.sheetCount & "л_" & cover.NameEnterprise
    
    ' Сохранить в word документ
    SaveAccompanying "ОПИСИ\" & fileNameWithRegister
    ActiveDocument.Close
End Sub
