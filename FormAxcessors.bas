Attribute VB_Name = "FormAxcessors"
Option Explicit

Public Function GetNameEnterprise()
    GetNameEnterprise = Form_Register.tb_NameEnterprise.value
End Function

Public Function GetOkpoEnterprise()
    GetOkpoEnterprise = Form_Register.tb_OkpoEnterprise.value
End Function

Public Function GetYears()
    GetYears = Form_Cover.tb_Years.value
End Function

Public Function GetSheetCount()
    GetSheetCount = Form_Cover.tb_sheetCount.value
End Function


Public Sub SetNameEnterprise()
    Form_Cover.tb_NameEnterprise.value = SelectElement().NameEnterprise
End Sub

Public Sub SetOkpoEnterprise()
    Form_Cover.tb_OkpoEnterprise.value = SelectElement().OkpoEnterprise
End Sub

Public Sub SetYears()
   Form_Cover.tb_Years.value = SelectElement().years
End Sub

Public Sub SetSheetCount()
    Form_Cover.tb_sheetCount.value = SelectElement().sheetCount
End Sub


Public Sub SetIndex()
    Form_Cover.tb_Index.value = SelectElement().index
End Sub

Private Function SelectElement() As C_CoverInfo
    Dim count As Integer
    count = KprBase.count
    If CurrentIndex <= 0 Then
        CurrentIndex = CurrentIndex + 1
    ElseIf CurrentIndex > KprBase.count Then
        CurrentIndex = CurrentIndex - 1
    End If
    Set SelectElement = KprBase.item(CurrentIndex)
End Function

Public Sub FillForm()
    SetIndex
    SetNameEnterprise
    SetOkpoEnterprise
    SetYears
    SetSheetCount
End Sub

' Добавляет на экран строку с информацией о предприятии
Public Sub AddInEnterpriseList(item As C_CoverInfo)
    Dim nameIndex As String
    Dim rowIndex As Integer
    
    nameIndex = TestNameZero(item.index)
    
    Form_Cover.lb_KprBase.AddItem nameIndex
    rowIndex = Form_Cover.lb_KprBase.ListCount - 1
    
    Form_Cover.lb_KprBase.List(rowIndex, 1) = item.OkpoEnterprise
    Form_Cover.lb_KprBase.List(rowIndex, 2) = item.NameEnterprise
    
    Form_Cover.lb_KprBase.BorderColor = vbBlack
    Form_Cover.lb_KprBase.TextColumn = 3
End Sub
