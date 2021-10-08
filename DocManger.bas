Attribute VB_Name = "DocManger"
Option Explicit

Public myDocs() As Document

' ������� �������� �� �������
Public Sub CreateNewDocment()
    Dim templatePath As String
    
    templatePath = "C:\Documents and Settings\PENSION\������� ����\������� ��� ������ ���\#����� ����������#_v_14_08_2021_1.dot"
    If Dir(templatePath) = "" Then
        templatePath = "C:\Users\Antis\Desktop\1111111111\#����� ����������#_v_14_08_2021_1.dot"
    End If
    
    Documents.Add templatePath
End Sub

' ������� �������� ������� �� �������
Public Sub CreateNewCoverDoc()
    Dim templatePath As String
    Dim templateName As String
    Dim fullPath As String
    
    templateName = "#�������#.dot"
    
    templatePath = "C:\Documents and Settings\PENSION\������� ����\������� ��� ������ ���\"
    If Dir(templatePath) = "" Then
        templatePath = "M:\My_projects\My_Progi\VBA\������������ �����\"
    End If
    
    fullPath = templatePath & templateName
    Documents.Add fullPath
End Sub

' ������� �������� ������� �� �������
Public Sub CreateNewRegisterDoc()
    Dim templatePath As String
    Dim templateName As String
    Dim fullPath As String
    
    templateName = "#����� ����������#.dot"
    
    templatePath = "C:\Documents and Settings\PENSION\������� ����\������� ��� ������ ���\"
    If Dir(templatePath) = "" Then
        templatePath = "M:\My_projects\My_Progi\VBA\������������ �����\"
    End If
    
    fullPath = templatePath & templateName
    Documents.Add fullPath
End Sub

Public Sub FillDocumentAndSave(registers As Collection)
            
    Dim item As C_RegisterInfo
    For Each item In registers
        Call Form_Register.FixPageNumbers(1, item)
        Call FillOneDocumentAndSave(item)
    Next
End Sub

Public Sub FillOneDocumentAndSave(curRegister As C_RegisterInfo)
       
    Dim cover As C_CoverInfo
    Set cover = PrepareCover(curRegister)
    
    ' ��������� � ���� ���� ����� ���� ��� ���� ������ = 0
    If Form_Register.cb_SaveInBase.value Or cover.index = 0 Then
        myBase.SaveKprBase cover
    End If
    
    ' ������ ��� �������� �����
    Dim nameIndex As String
    nameIndex = TestNameZero(cover.index)
    
    Call FillOneCoverDocAndSave(nameIndex, cover)
    Call FillOneRegisterDocAndSave(nameIndex, cover)
    
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
        
        ' ���� ������ ����� 250, �� ��������� � ����� ��� ������� numRegister
        If (SheetsCount < 251) Then
            sheets(numRegister) = SheetsCount
        End If
        
        '���� ���������� ������ ������ 250, ���� ��������� ����� ��� ������� ����� �����
        Do While (SheetsCount > 250)
            numRegister = numRegister + 1
            
            ' ��� ��������� ������, ������� �����
            If registers.count < numRegister Then
                registers.Add New C_RegisterInfo
            End If
            
            If numRegister > 6 Then
                MsgBox "���������� ����� ������ 6-��!!!"
                GoTo endMark
            End If
            ' ��������� ����� ������ ����� � ������� numRegister � ������� ������
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

Private Sub FillOneCoverDocAndSave(nameIndex As String, cover As C_CoverInfo)
    Dim newFolderName As String
    newFolderName = "�������"
    Call CreateFolder(newFolderName)
    
    Call FillOneCoverDoc(cover)
    
    ' ������������� ��� ��������� "�������"
    Dim fileNameWithCover As String
    fileNameWithCover = nameIndex & "_C_" & cover.sheetCount & "�_" & cover.NameEnterprise
        
    ' ���������  word ��������
    SaveAccompanying newFolderName & "\" & fileNameWithCover
    ActiveDocument.Close
End Sub



Private Sub FillOneRegisterDocAndSave(nameIndex As String, cover As C_CoverInfo)
    Dim newFolderName As String
    newFolderName = "�����"
    Call CreateFolder(newFolderName)
    
    Call FillOneRegisterDoc(cover)
    
    ' ������������� ��� ��������� "���������� �����"
    Dim fileNameWithRegister As String
    fileNameWithRegister = nameIndex & "_R_" & cover.sheetCount & "�_" & cover.NameEnterprise
    
    ' ��������� � word ��������
    SaveAccompanying newFolderName & "\" & fileNameWithRegister
    ActiveDocument.Close
End Sub



Private Sub CreateFolder(newFolderName As String)
    Dim newFolderPath As String
    
    newFolderPath = ActiveDocument.path & "\" & newFolderName
    
    If Dir(newFolderPath, vbDirectory) = "" Then
        MkDir newFolderPath
    End If
End Sub


'---------------------------------------------------------------------------------------------

Public Sub FillDocumentAndAggregate(registers As Collection)
            
    Dim item As C_RegisterInfo
    For Each item In registers
        Call Form_Register.FixPageNumbers(1, item)
        Call FillOneCoverDocAndAggregate(item)
    Next
    
    For Each item In registers
        Call FillOneRegisterDocAndAggregate(item)
    Next
End Sub


Public Sub FillOneCoverDocAndAggregate(curRegister As C_RegisterInfo)
    Dim cover As C_CoverInfo
    Set cover = PrepareCover(curRegister)
    Call FillOneCoverDoc(cover)
    
End Sub

Public Sub FillOneRegisterDocAndAggregate(curRegister As C_RegisterInfo)
    Dim cover As C_CoverInfo
    Set cover = PrepareCover(curRegister)
    Call FillOneCoverDoc(cover)
End Sub

Private Sub FillOneCoverDoc(cover As C_CoverInfo)
    ' ������� ����� ���� �� ������� � ��������
   Call CreateNewCoverDoc
    
    ' ��������� �������� � ��������
   Call FillCoverDoc(cover)
End Sub

Private Sub FillOneRegisterDoc(cover As C_CoverInfo)
    ' ������� ����� ���� �� ������� � ������
   Call CreateNewRegisterDoc
    
    ' ��������� �������� ���������� �����
   Call FillRegister(cover)
End Sub
