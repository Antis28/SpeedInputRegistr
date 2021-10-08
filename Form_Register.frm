VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Register 
   OleObjectBlob   =   "Form_Register.frx":0000
   Caption         =   "ОБЛОЖКА + ВНУТРЕННЯЯ ОПИСЬ"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7572
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   120
End
Attribute VB_Name = "Form_Register"
Attribute VB_Base = "0{1EB37D80-B564-4705-A43F-8041382E3FE4}{4473B81F-4065-4D70-B524-9C444C973791}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private id As Integer
Private regNumber As Integer 'номер описи в базе
Private formCaption As String
Private oldSpinValue As Integer

Private Sub cb_calc_Click()
    tb_SheetsNumber.text = CalculateSheetCount(tb_sheetCount.text, register.count)
End Sub

Private Sub bt_LoadNext_Click()
    regNumber = regNumber + 1
    Call LoadRegister(regNumber)
End Sub

Private Sub cb_FillAll_Click()
    Dim i As Integer
    For i = myBase.CountInKprBase To 1 Step -1
        Call LoadRegister(i)
        Call FillDocByTemplates
    Next i
End Sub

Private Sub LoadLast_Click()
    regNumber = regNumber - 1
    Call LoadRegister(regNumber)
End Sub

Private Sub cb_addOKPO_Click()
    tb_docNumber.text = tb_docNumber.text & "/" & tb_OkpoEnterprise.text
End Sub

Private Sub cb_AddTemplates_Click()
    myBase.AddNewNameDoc (cb_docNames.value)
    ShowRecordTemplates
End Sub

Private Sub cb_ClearRecords_Click()
   Call CreateNewRegister
End Sub

Private Sub cb_ClearRegistr_Click()
    Call CreateNewRegister
    Call ClearHeaderRegister
    Call ClearInputField
End Sub

Private Sub cb_RemoveRegisterFromBase_Click()
    Dim userAnswer As VbMsgBoxResult
    
    userAnswer = MsgBox("Вы уверены, что хотите удалить опись из базы? Действие необратимо!!!", vbYesNo)
    If userAnswer = vbNo Then
        Exit Sub
    End If
    
    myBase.RemoveRegister (l_indexInBase)
    myBase.UpdateBase
    
    Call CreateNewRegister
    Call ClearHeaderRegister
    Call ClearInputField
End Sub

Private Sub cb_SaveTobase_Click()
    Dim listIndex As Integer
    Dim regNumber As Integer
    Dim cover As C_CoverInfo
    
    listIndex = lb_RecordsList.listIndex
    
    regNumber = Int(l_indexInBase.Caption)
 
 ' Сохранить в базу

    If myBase.IndexExists(regNumber) Then
        ' Взять существующую
        Set cover = CreateCoverWithOutID(register)
        cover.index = regNumber
        myBase.UpdateRecordInKprBase cover
    Else
        ' Создать новую
        Set cover = CreateCoverWithOutID(register)
        myBase.SaveKprBase cover
    End If
    
    myBase.LoadAllBases
    
    If l_indexInBase.Caption = "0" Then
        bt_LoadNext_Click
    End If
    
    If listIndex < lb_RecordsList.ListCount - 1 Then
        lb_RecordsList.listIndex = listIndex
    Else
         lb_RecordsList.listIndex = lb_RecordsList.ListCount - 1
    End If
    
    Call ClearInputField
    cb_docNames.SetFocus
End Sub

Private Sub CommandButton6_Click()
    l_indexInBase.Caption = myBase.CountInKprBase + 1
End Sub



Private Sub cb_CleanAllFields_Click()
    Call ClearInputField
End Sub

Private Sub cb_docNames_Change()
    If cb_docNames.listIndex > -1 Then
        tb_DocName.text = cb_docNames.List(cb_docNames.listIndex)
    End If
    
    Call CheckExistDocumentName
    
    If cb_docNames.listIndex > -1 Then
        tb_DocName.text = cb_docNames.List(cb_docNames.listIndex)
    End If
    
End Sub

Private Sub cb_SaveTemplates_Click()
    myBase.SaveSettings
End Sub

Private Sub cb_UpdateTemplates_Click()
    myBase.LoadNamesBase
    ShowRecordTemplates
End Sub


Private Sub SpinButton1_Change()
    If SpinButton1.value > oldSpinValue Then
        MoveRecordToUp
    Else
        MoveRecordToDown
    End If
    oldSpinValue = SpinButton1.value
End Sub

Private Sub tb_DateFirst_Change()
    Dim sheetCount As String
    sheetCount = tb_DateFirst.value
    
    If Not IsNumeric(sheetCount) And Not (sheetCount = "") Then
        MsgBox "Количесво листов должно быть числом!"
        tb_DateFirst.value = ""
    End If
End Sub

Private Sub tb_DateLast_Change()
    Dim sheetCount As String
    sheetCount = tb_DateLast.value
    
    If Not IsNumeric(sheetCount) And Not (sheetCount = "") Then
        MsgBox "Количесво листов должно быть числом!"
        tb_DateLast.value = ""
    End If
End Sub

Private Sub tb_sheetCount_Change()
    Dim sheetCount As String
    sheetCount = tb_sheetCount.value
    
    If Not IsNumeric(sheetCount) And Not (sheetCount = "") Then
        MsgBox "Количесво листов должно быть числом!"
        tb_sheetCount.value = ""
    End If
        
End Sub

Private Sub UserForm_Activate()
    oldSpinValue = SpinButton1.value
    formCaption = Form_Register.Caption
    InitUniqNumbers
    Set register = New C_RegisterInfo
    
    'Настройка визуального списка
    Dim widthNumber As String, widthDateStart As String, _
        widthDateEnd As String, widthSheetNumber As String, _
        widthCount As String, widthDescription As String

    widthNumber = "20 pt"
    widthDateStart = "32 pt"
    widthDateEnd = "32 pt"
    widthSheetNumber = "55 pt"
    widthCount = "25 pt"
    widthDescription = "1000"
   
    lb_RecordsList.ColumnCount = 6
    lb_RecordsList.ColumnWidths = widthNumber & ";" & widthDateStart & ";" & widthDateEnd & ";" & widthSheetNumber & ";" & widthCount & ";" & widthDescription
    lb_RecordsList.BorderColor = vbBlack
    lb_RecordsList.TextColumn = 6
   
   Call CreateListBoxHeader(Me.lb_RecordsList, Me.listBox_Header, Array("№", "Нач.", "Кон.", "Лист.", "Кол.", "Название"))
   
   id = 0
   
   myBase.LoadAllBases
   ShowRecordTemplates
   
   regNumber = myBase.CountInKprBase
   
End Sub


Private Sub btn_DeleteRecord_Click()
    If register.count > 0 And lb_RecordsList.listIndex > -1 Then
        Dim id As Integer
        id = lb_RecordsList.listIndex
        
        register.Remove lb_RecordsList.listIndex + 1
        
        If cb_SortList.value Then SortRecordList
        
        Call FixPageNumbers(lb_RecordsList.listIndex + 1, register)
        UpdateScreen
        
        lb_RecordsList.listIndex = id - 1
    End If
End Sub



Private Sub cb_LineUp_Click()
    MoveRecordToUp
End Sub

Private Sub cb_LineDown_Click()
    MoveRecordToDown
End Sub


Private Sub cb_FillDocument_Click()
    
    Call FillDocByTemplates
    
    'HideForm
    If cb_CloseAfterFilling.value Then
        Application.Quit
    End If
End Sub


Private Sub cb_UpdateRecord_Click()
    If Not IsNumeric(tb_DateFirst.text) _
        Or Not IsNumeric(tb_DateLast.text) _
        Or lb_RecordsList.listIndex < 0 Then
       ' MsgBox ("Не заполнено поле: " & _
       '         "количество листов:" & tb_sheetCount.text & _
       '         "Начальная дата:" & tb_DateFirst.text & _
       '         "Конечная дата:" & tb_DateLast.text)
        Exit Sub
    End If
    
    ' Получить id текущего элемента перед сортировкой
    Dim id As Integer
    id = register.getRecord(lb_RecordsList.listIndex + 1).id
    ' обновить запись в реестре
    Call UpdateRecordInReg
    ' отсортировать по датам
    If cb_SortList.value Then SortRecordList
    ' пересчитать номера листов
    Call FixPageNumbers(lb_RecordsList.listIndex + 1, register)
    ' обновить список
    UpdateScreen
    
    SelectById id
    
    Call ClearInputField
    Call ClearTemlateField
End Sub

Private Sub CommandButton5_Click()
    TestSort
End Sub

Private Sub lb_RecordsList_Click()
    Dim record As C_RecordInfo
    Set record = register.getRecord(lb_RecordsList.listIndex + 1)
    tb_DateFirst.text = record.DateFirst
    tb_DateLast.text = record.dateLast
    tb_DocName.text = record.docName
    tb_sheetCount.text = record.SheetsCount
    tb_SheetsNumber.text = record.sheetsNumber
    tb_docNumber.text = record.docNumber
End Sub


Private Sub cb_SortRecordsList_Click()
    register.SortByLastDate
    ' Визуализировать список
    VisaliseRegister register
End Sub

Private Sub cb_AddDocInList_Click()

    'Call CheckExistDocumentName

     'Собрали информацию о документе(ах) одного типа
    Dim record As C_RecordInfo
    Set record = New C_RecordInfo
'------------------------------------------------------
    Dim dFirst As Integer
    Dim dLast As Integer
    Dim dtempDate As Integer
    
    If tb_DateFirst.text <> "" Then
        dFirst = tb_DateFirst.text
    ElseIf tb_DateLast.text <> "" Then
        dFirst = tb_DateLast.text
    End If
       
    If tb_DateLast.text <> "" Then
        dLast = tb_DateLast.text
     ElseIf tb_DateFirst.text <> "" Then
        dLast = tb_DateFirst.text
    End If
        
    If dLast = 0 And dFirst = 0 Then
        MsgBox "Дата пуста!"
        GoTo endsub
    End If
    
    'Поменять местами даты если они в неправильном порядке
    If dLast < dFirst Then
        dtempDate = dFirst
        dFirst = dLast
        dLast = dtempDate
    End If
    '-----------------------------------------------------
    If tb_DocName.text = "" Then
        MsgBox "Имя документа пустое!"
        GoTo endsub
    End If
    
    If tb_docNumber.text = "" Then
        tb_docNumber.text = "-"
    End If
    
    '------------------------------------------------------------------
    Dim sheetCount As Integer
    sheetCount = GetSheetCount()
         
    Dim newId As Integer
    newId = GetUniqNumber()
    
    
         
    record.Construct newId, _
                        dFirst, _
                        dLast, _
                        tb_DocName.text, _
                        tb_docNumber.text, _
                        "0", _
                        sheetCount

   ' Добавить в опись
   register.Add record
   
    
    If cb_SortList.value Then SortRecordList
    Call FixPageNumbers(1, register)
      
    UpdateScreen
    
    SelectById newId
    
    Call ClearInputField
    Call ClearTemlateField
    
endsub:
End Sub


'Визуализация списка на форме
Private Sub VisaliseRegister(register As C_RegisterInfo)
    Dim i As Integer
    Dim item As C_RecordInfo
    id = 0
    lb_RecordsList.Clear
    For i = 1 To register.count Step 1
        Set item = register.getRecord(i)
        
        VisaliseRecord item
    Next i
End Sub

Private Sub VisaliseRecord(item As C_RecordInfo)
    Dim nameIndex As String
    Dim rowIndex As Integer
    id = id + 1
    nameIndex = TestNameZero(id)
    
    lb_RecordsList.AddItem nameIndex
    rowIndex = lb_RecordsList.ListCount - 1
        
    
    
    If item.DateFirst <> 0 Then
        lb_RecordsList.List(rowIndex, 1) = item.DateFirst
    End If
    If item.dateLast <> 0 Then
        lb_RecordsList.List(rowIndex, 2) = item.dateLast
    End If
    
    With lb_RecordsList
        .List(rowIndex, 3) = item.sheetsNumber
        .List(rowIndex, 4) = item.SheetsCount
        .List(rowIndex, 5) = item.docName
    End With
End Sub



Public Sub UpdateScreen()
  
    VisaliseRegister register
    
    Dim sheetsLeft As Integer
    sheetsLeft = register.GetSheetCount
    l_sheetsLeft.Caption = sheetsLeft
    If sheetsLeft > 250 Then
        l_sheetsLeft.BackColor = vbRed
        l_sheetsLeft.ForeColor = vbYellow
    ElseIf sheetsLeft > 200 Then
        l_sheetsLeft.BackColor = vbYellow
        l_sheetsLeft.ForeColor = vbButtonText
    ElseIf sheetsLeft > 100 Then
       l_sheetsLeft.BackColor = vbGreen
       l_sheetsLeft.ForeColor = vbButtonText
    Else
        l_sheetsLeft.BackColor = "040001"
        l_sheetsLeft.ForeColor = vbButtonText
    End If
End Sub

Public Sub SortRecordList()
     ' Отсортировать и отрисовать список
    If register.count < 1 Then
        Exit Sub
    End If
    'register.SortByFirstDate
    'register.SortByLastDate
    register.SortByTwoParameters
End Sub




Private Sub cb_docNames_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookScroll Me.cb_docNames
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnHookScroll
End Sub


Private Sub lb_RecordsList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookScroll Me.lb_RecordsList
End Sub

Public Sub ShowRecordTemplates()
    cb_docNames.Clear
    Dim item As Variant
    Dim i As Integer
    For i = 1 To myBase.CountInSettings
        item = myBase.GetSettingsItem(i)
        cb_docNames.AddItem item
    Next
End Sub
'---------------------------------------------------



' Выравнивает ListBoxHeader с основным ListBox
Public Sub CreateListBoxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
        ' make column count match
        header.ColumnCount = body.ColumnCount
        header.ColumnWidths = body.ColumnWidths

        ' add header elements
        header.Clear
        header.AddItem
        Dim i As Integer
        For i = 0 To UBound(arrHeaders)
            header.List(0, i) = arrHeaders(i)
        Next i

        ' make it pretty
        body.ZOrder (1)
        header.ZOrder (0)
        header.SpecialEffect = fmSpecialEffectFlat
        header.BackColor = RGB(210, 210, 210)
        header.Height = 10

        ' align header to body (should be done last!)
        header.Width = body.Width
        header.left = body.left
        header.Top = body.Top - (header.Height - 1)
End Sub

' Пересчитывает все страницы от startIndex и до конца списка
Public Sub FixPageNumbers(startIndex As Integer, curRegister As C_RegisterInfo)
    Dim sheetsNumber As String
    Dim item As C_RecordInfo
    Dim i As Integer
    For i = startIndex To curRegister.count
        Set item = curRegister.getRecord(i)
        sheetsNumber = CalculateSheetCount(item.SheetsCount, i, curRegister)
        item.sheetsNumber = sheetsNumber
    Next
End Sub


' Расчитывет номера страниц(начало - конец) до lastIndex
Private Function CalculateSheetCount(count As Integer, lastIndex, curRegist As C_RegisterInfo) As String
    Dim item As C_RecordInfo
    Dim i As Integer
    
    Dim sheetSummaNminus1 As Integer
    Dim startPage As Integer
    Dim endPage As Integer
    Dim result As String
        
    For i = 1 To lastIndex - 1
        Set item = curRegist.getRecord(i)
        sheetSummaNminus1 = sheetSummaNminus1 + item.SheetsCount
    Next
    
    startPage = sheetSummaNminus1 + 1
    endPage = sheetSummaNminus1 + count
    
    If count > 1 Then
        result = startPage & " - " & endPage
    End If
    
    If count = 1 Then
        result = endPage
    End If
    
    CalculateSheetCount = result
End Function




Private Sub LoadRegisterByIndex(regNumber As Integer)
    Dim temp As C_CoverInfo
    Dim tempRecord As C_RecordInfo
    Dim newRegister As New Collection
    Dim item As C_RecordInfo
    
    Set temp = myBase.GetKprItem(regNumber)
    ' Создаем копию описи чтобы не затереть оригинал
    For Each item In temp.innerRegistry
        Set tempRecord = New C_RecordInfo
        tempRecord.Construct GetUniqNumber(), item.DateFirst, item.dateLast, item.docName, item.docNumber, item.sheetsNumber, item.SheetsCount
        newRegister.Add tempRecord
    Next
    
    register.setCollection newRegister
    
    tb_NameEnterprise.text = temp.NameEnterprise
    tb_OkpoEnterprise.text = temp.OkpoEnterprise
    
    l_indexInBase.Caption = temp.index
End Sub

Private Function CheckRegNumber(ByVal index As Integer) As Integer
    If index < 1 Then
       index = 1
    ElseIf index > myBase.CountInKprBase Then
       index = myBase.CountInKprBase
    End If
    CheckRegNumber = index
End Function

Public Sub ClearInputField()
    tb_DateFirst.text = ""
    tb_DateLast.text = ""
    tb_DocName.text = ""
    tb_SheetsNumber.text = ""
    tb_sheetCount.text = ""
    tb_docNumber.text = ""
    
    tb_lastSheet.text = ""
End Sub

Public Sub ClearTemlateField()
    cb_docNames.listIndex = -1
    'tb_DateFirst.SetFocus
    cb_docNames.SetFocus
End Sub

Public Sub CreateNewRegister()
    register.setCollection New Collection
    Call ClearInputField
    Call ClearTemlateField
    UpdateScreen
End Sub

Public Sub ClearHeaderRegister()
    tb_NameEnterprise.text = ""
    tb_OkpoEnterprise.text = ""
    l_indexInBase.Caption = "0"
End Sub

Private Sub UpdateRecordInReg()
    Dim record As C_RecordInfo
    Set record = register.getRecord(lb_RecordsList.listIndex + 1)
    
    record.DateFirst = tb_DateFirst.text
    record.dateLast = tb_DateLast.text
    
    record.docName = tb_DocName.text
    record.sheetsNumber = tb_SheetsNumber.text
     
    record.SheetsCount = tb_sheetCount.text
    
    record.docNumber = tb_docNumber.text
End Sub

Private Sub SelectById(id As Integer)
    
    Dim item As C_RecordInfo, counter As Integer
    ' Найти номер строки по id элемента
    For Each item In register.getCollection
        counter = counter + 1
        If id = item.id Then
            ' Выделить элемент согласно его номеру
            lb_RecordsList.listIndex = counter - 1 'lb_RecordsList.ListCount - 1
            Exit Sub
        End If
    Next
End Sub

Public Sub CheckExistDocumentName()
      'Проверка на существования документа с таким именем
    Dim item As C_RecordInfo
    For Each item In register.getCollection
        If Trim(item.docName) = Trim(tb_DocName.text) Then
            'MsgBox "Запись с таким именем документа уже существует!!"
            SelectById item.id
            Exit Sub
        End If
    Next
    lb_RecordsList.listIndex = -1
    Call ClearInputField
End Sub

Private Sub ClearAllFlags()
    cb_SortList.value = False
    cb_SaveInBase.value = False
    cb_CloseAfterFilling.value = False
End Sub

Private Function GetSheetCount() As Integer
    Dim sheetCount As Integer
    
    If tb_sheetCount.text = "" And tb_lastSheet.text = "" Then
        MsgBox "Не указано количество документов!"
        GoTo endsub
    End If
    
    If Not IsNumeric(tb_sheetCount.text) And Not IsNumeric(tb_lastSheet.text) Then
        MsgBox "Количество документов не число!"
        GoTo endsub
    End If
    
    If IsNumeric(tb_sheetCount.text) Then
        sheetCount = Int(tb_sheetCount.text)
    Else
        sheetCount = Int(tb_lastSheet.text) - register.GetSheetCount
    End If
endsub:
    GetSheetCount = sheetCount
End Function

Private Function CalculateNewSheetCount()

End Function

Private Sub ShowLastChanged(regNumber As Integer)
    Form_Register.Caption = ""
    Form_Register.Caption = formCaption & " Изменен: " & myBase.GetKprItem(regNumber).lastChange
End Sub

Private Sub LoadRegister(ByVal index As Integer)
    Dim regNumber As Integer
    regNumber = CheckRegNumber(index)
    Call LoadRegisterByIndex(regNumber)
    Call UpdateScreen
    Call ClearAllFlags
    
    Call ShowLastChanged(regNumber)
End Sub

Private Sub MoveRecordToUp()
    Dim index As Integer
    index = lb_RecordsList.listIndex
    
    If index < 1 Then
        Exit Sub
    End If
    
    Call register.swap(index - 1, index)
    
      ' Пересчет страниц
    Call FixPageNumbers(1, register)
    
    UpdateScreen
    
    ' Выделение строки
    lb_RecordsList.listIndex = index - 1
End Sub

Private Sub MoveRecordToDown()
    Dim index As Integer
    index = lb_RecordsList.listIndex

    If index < 0 Or index > lb_RecordsList.ListCount - 2 Then
          Exit Sub
    End If
    
    register.swap index + 1, index
    
    ' Пересчет страниц
    Call FixPageNumbers(1, register)
    
    UpdateScreen
    
     ' Выделение строки
    lb_RecordsList.listIndex = index + 1
End Sub

Private Sub FillDocByTemplates()
    Dim registers As Collection ' список описей

    ' Нет записей в описи
    If register.count = 0 Then
        Exit Sub
    End If
    
    ' если в деле больше 250 стр., то его необходимо поделить на тома
    Set registers = DivRegister()
    
    
    ' заполнить шаблоны для описи и обложки
    FillDocument registers
End Sub
