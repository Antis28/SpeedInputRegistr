Attribute VB_Name = "FormMemory"
Option Explicit

Public myBase As New C_Base

Public Sub SaveToBase()
    Dim listIndex As Integer
    Dim regNumber As Integer
    Dim cover As C_CoverInfo
    
    ' Индекс записи
    listIndex = Form_Register.lb_RecordsList.listIndex
    
    ' Индекс в базе
    regNumber = Int(Form_Register.l_indexInBase.Caption)
 
    ' Сохранить в базу

    If myBase.IndexExists(regNumber) Then
        ' Перезаписать существующую обложку
        Set cover = CreateCoverWithOutID(register)
        cover.index = regNumber
        myBase.UpdateRecordInKprBase cover
    Else
        ' Создать новую обложку
        Set cover = CreateCoverWithOutID(register)
        myBase.SaveKprBase cover
    End If
    
    myBase.LoadAllBases
    ' Если нет номера, загрузить крайнюю обложку
    If Form_Register.l_indexInBase.Caption = "0" Then
        Call Form_Register.LoadNextCover
    End If
    
    ' Сохраненный индекс записи в пределах списка
    If listIndex < Form_Register.lb_RecordsList.ListCount - 1 Then
        Form_Register.lb_RecordsList.listIndex = listIndex
    Else
    ' Иначе выставляем предпоследний
        Form_Register.lb_RecordsList.listIndex = Form_Register.lb_RecordsList.ListCount - 1
    End If
    
    Call Form_Register.ClearInputField
    Call Form_Register.cb_docNames.SetFocus
End Sub
