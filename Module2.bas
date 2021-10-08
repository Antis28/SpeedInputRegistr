Attribute VB_Name = "Module2"
Option Explicit


Public Sub NumberInWords___()
     Dim pages As String
     Dim pages2 As String
     Dim num As Currency
     pages = InputBox("Введите количество страниц", "Число страниц")
     num = Int(pages)
     FillTotalInWords num
End Sub

Public Sub TestSort()
    
    ' Добавить в опись
    register.Add RecConstructor(1991, 1992, "Test_1", 1)
    register.Add RecConstructor(1991, 1993, "Test_2", 1)
    register.Add RecConstructor(1990, 1991, "Test_0", 10)
    register.Add RecConstructor(1992, 1993, "Test_3", 3)
    register.Add RecConstructor(1994, 1995, "Test_4", 4)
    register.Add RecConstructor(1992, 1995, "Test_5", 5)
    register.Add RecConstructor(1993, 1994, "Test_6", 6)
    register.Add RecConstructor(1991, 1996, "Test_7", 7)
    register.Add RecConstructor(2000, 2005, "Test_8", 8)
    register.Add RecConstructor(2003, 2005, "Test_9", 2)
    register.Add RecConstructor(2005, 2005, "Test_10", 1)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    register.Add RecConstructor(2005, 2005, "Test_11", 2)
    
    
    
    Form_Register.SortRecordList
    Form_Register.UpdateScreen
End Sub

Public Sub TestSort3()
    
    ' Добавить в опись
    register.Add RecConstructor(1991, 1993, "Test_9", 9)
    register.Add RecConstructor(1990, 1991, "Test_1", 1)
    register.Add RecConstructor(1993, 1993, "Test_11", 11)
    register.Add RecConstructor(1990, 1992, "Test_4", 4)
    register.Add RecConstructor(1990, 1993, "Test_8", 8)
    register.Add RecConstructor(1992, 1992, "Test_7", 7)
    register.Add RecConstructor(1991, 1991, "Test_2", 2)
    register.Add RecConstructor(1992, 1993, "Test_10", 10)
    register.Add RecConstructor(1991, 1991, "Test_3", 3)
    register.Add RecConstructor(1993, 1993, "Test_12", 12)
    register.Add RecConstructor(1992, 1992, "Test_6", 6)
    register.Add RecConstructor(1991, 1992, "Test_5", 5)
      
    
    Form_Register.SortRecordList
    Form_Register.UpdateScreen
End Sub


Public Sub TestSort2()
    
    ' Добавить в опись
    register.Add RecConstructor(1990, 1991, "Test_1", 1)
    register.Add RecConstructor(1991, 1991, "Test_2", 2)
    register.Add RecConstructor(1991, 1991, "Test_3", 3)
    register.Add RecConstructor(1990, 1992, "Test_4", 4)
    register.Add RecConstructor(1991, 1992, "Test_5", 5)
    register.Add RecConstructor(1992, 1992, "Test_6", 6)
    register.Add RecConstructor(1992, 1992, "Test_7", 7)
    register.Add RecConstructor(1990, 1993, "Test_8", 8)
    register.Add RecConstructor(1991, 1993, "Test_9", 9)
    register.Add RecConstructor(1992, 1993, "Test_10", 10)
    register.Add RecConstructor(1993, 1993, "Test_11", 11)
    register.Add RecConstructor(1993, 1993, "Test_12", 12)
    
    Dim sorter As New ArraySorter
    register.SortByLastDate
    
    Dim newCol As Collection
    Set newCol = register.getCollection
    
    FilterEquils newCol
    
    register.setCollection newCol
    Form_Register.UpdateScreen
End Sub



Public Function RecConstructor(df As Integer, dl As Integer, name As String, count As Integer) As C_RecordInfo
    Dim record As C_RecordInfo
    Set record = New C_RecordInfo
    
    record.Construct GetUniqNumber(), df, dl, name, "2", "", count
    Set RecConstructor = record
End Function


Public Sub FilterEquils(ByRef coll As Collection)
    ' Разбиваем массив на отдельные массивы по 2-му значению
    Dim eq As New Collection, _
    Element As C_RecordInfo, _
    elementPivot As C_RecordInfo, _
    id As Integer
    
    Dim currColl As Collection
    Set currColl = New Collection
    id = 1
    Do While id <= coll.count
        Set currColl = New Collection
        Set Element = coll(id)
        Set elementPivot = coll(id)
        
        Do While Element.dateLast = elementPivot.dateLast
            currColl.Add Element
            id = id + 1
            If id > coll.count Then
                Exit Do
            End If
            Set Element = coll(id)
        Loop
        If currColl.count > 0 Then
            eq.Add currColl
        End If
    Loop
        
    
    ' Сортируем каждый отдельный массив по 1-му значению
    Dim sorter As New ArraySorter
    
    Dim newColl As New Collection, _
    newItemRec As Collection
    
    For id = 1 To eq.count
        Set newItemRec = eq(id)
        sorter.sortCollection newItemRec, "LeftMoreRightDateFirst", "LeftLessRightDateFirst"
        newColl.Add newItemRec
    Next id
    
    
    ' Соеденяем массивы обратно
    
    Dim allColl As New Collection, _
    innerColl As Collection, _
    endItem As C_RecordInfo
    
    For Each innerColl In newColl
        For Each endItem In innerColl
            allColl.Add endItem
        Next endItem
    Next innerColl
    
    Set coll = allColl
    
End Sub


Public Sub test55()
    Dim reader As New XmlReader
    Dim nodeList As IXMLDOMNodeList
    Dim fileNameAndExtention As String
    Dim attributesNames As New XML_Attributes
    
    fileNameAndExtention = "Base.xml"
    reader.InitXML TemplateProject.ThisDocument.path & "\" & fileNameAndExtention
    Set nodeList = reader.ReadRoot(attributesNames.EnterpriseElement)
    
    Dim i As Integer
    Dim node As IXMLDOMNode
    
    Dim coll As Collection
    
    For i = 0 To nodeList.Length - 1
        Set node = nodeList.item(i)
        
        Set coll = reader.ReadAttributes(node, attributesNames.name)
    Next i
   
End Sub


Public Sub test56()
    Dim fileName As String
    fileName = "Base"
    
    Dim worker As New C_WorkerWithBase
    worker.Init fileName
    
    Dim cov As New C_CoverInfo
    cov.Construct 1, "Test", "12345678", "11-11", "0", New Collection
    
    worker.Add cov
    
    worker.Save
End Sub


Public Sub test57()
    Dim reader As New C_LoaderSettings
    Dim writer As New C_SaverSettings
    
    Dim coll As New Collection
    coll.Add "Акты проверки правильности полноты начисления, своевременности уплаты страховых взносов Пенсионному фонду Украины"
    coll.Add "Информация о расчетах с Пенсионным фондом (Додаток №9)"
    
    coll.Add "Мемориальные ордера"
    coll.Add "Отчеты о начислении страховых взносов и расходования средств пенсионного фонда (форма 4-ПФ), додаток №6, 7)"
    
    coll.Add "Платежные поручения"
    coll.Add "Расчетные ведомости"
    coll.Add "Расчет определения сумм, причитающихся к уплате в Пенсионный фонд"
    coll.Add "Расчет обязательств по оплате сборов на обязательное государственное пенсионное страхование (додаток №22)"
    coll.Add "Расчет начисления пени (приложение №2)"
    
    coll.Add "Сведения о регистрации"
    coll.Add "Справки проверки правильности полноты начисления, своевременности уплаты страховых взносов Пенсионному фонду Украины"
    coll.Add "Справка о включении в единый государственный реестр предприятий и организаций Украины"
    
    Dim fileName As String
    Dim fullPath As String
    
    fileName = "Settings"
    fullPath = TemplateProject.ThisDocument.path & "\" & fileName & ".xml"
    
       
    writer.Init fileName, fullPath
    writer.Save coll
   
End Sub

Public Sub test58()
    Dim fileName As String
    fileName = "DocumentNames"
    
    Dim worker As New C_WorkerWithSettings
    worker.Init fileName
    
    Dim coll As New Collection
'    coll.Add "Акты проверки правильности полноты начисления, своевременности уплаты страховых взносов Пенсионному фонду Украины"
'    coll.Add "Информация о расчетах с Пенсионным фондом (Додаток №9)"
'
'    coll.Add "Мемориальные ордера"
'    coll.Add "Отчеты о начислении страховых взносов и расходования средств пенсионного фонда (форма 4-ПФ), додаток №6, 7)"
'
'    coll.Add "Платежные поручения"
'    coll.Add "Расчетные ведомости"
'    coll.Add "Расчет определения сумм, причитающихся к уплате в Пенсионный фонд"
'    coll.Add "Расчет обязательств по оплате сборов на обязательное государственное пенсионное страхование (додаток №22)"
'    coll.Add "Расчет начисления пени (приложение №2)"
'
'    coll.Add "Сведения о регистрации"
'    coll.Add "Справки проверки правильности полноты начисления, своевременности уплаты страховых взносов Пенсионному фонду Украины"
'    coll.Add "Справка о включении в единый государственный реестр предприятий и организаций Украины"
    
    worker.Add "Test"
    
    worker.Save
End Sub
