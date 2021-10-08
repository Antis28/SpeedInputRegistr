Attribute VB_Name = "Main"
Option Explicit

Public Sub AutoNew()
    'ShowForm
End Sub
Public Sub AutoOpen()
    ShowForm
End Sub


Public Sub ShowForm()
    Form_Register.Show
End Sub

Public Sub HideForm()
    Form_Register.Hide
End Sub

' Заполнить документ с обложкой
Public Sub FillCoverDoc(cover As C_CoverInfo)
 
    Dim nameEnterpriseMark As String
    Dim okpoEnterpriseMark As String
    Dim yearsMark As String
    Dim sheetCountMark As String
    
    nameEnterpriseMark = "NameEnterprise"
    okpoEnterpriseMark = "OkpoNumber"
    yearsMark = "Years"
    sheetCountMark = "SheetCount"
    
    
   '*********************************************
    Dim printer As C_PrinterMashine
    Set printer = New C_PrinterMashine
    
    printer.PrintByBookmark nameEnterpriseMark, cover.NameEnterprise
    printer.PrintByBookmark okpoEnterpriseMark, cover.OkpoEnterprise
    printer.PrintByBookmark yearsMark, cover.years
    printer.PrintByBookmark sheetCountMark, cover.sheetCount
     
End Sub

Public Function TestNameZero(nameIndex As Integer) As String
    Dim result As String
    result = nameIndex
    
    If Int(nameIndex) < 10 Then
        result = "0" & nameIndex
    End If
    
    TestNameZero = result
End Function
