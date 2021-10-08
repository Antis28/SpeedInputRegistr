Attribute VB_Name = "RegisterFiller"
Option Explicit

Public register As C_RegisterInfo
Private worker  As C_TableWorker

Public Sub FillRegister(cover As C_CoverInfo)
    Dim record As C_RecordInfo
    Dim i As Integer
    
    Set worker = New C_TableWorker
    
    'Установить начальную ячейку
    'worker.MoveToBookmark "DocDate"
    worker.MoveToBookmark "DocNumber"
    
    
    'Получить первую запись
    i = 1
    Set record = cover.innerRegistry(i)
    ' Заполняем первую строку( чтобы убрать доп. проверку)
    Call FillRow(1, record)
    
    For i = 2 To cover.innerRegistry.count
        'Создать новую строку снизу
        Selection.InsertRowsBelow 1
        'Перейти во второй столбец
        worker.MoveRight 2
        'Получить текущую запись
        Set record = cover.innerRegistry(i)
        ' Заполнить строку
        Call FillRow(i, record)
    Next i
    
    ' Заполнить количество листов в деле
    If IsNumeric(cover.sheetCount) Then
        FillTotalInWords cover.sheetCount
    Else
        FillTotalInWords 0
    End If
    
    Dim sheetCount As Integer
    'Получить количество листов в документе
    'sheetCount = ActiveDocument.Content.Information(wdActiveEndAdjustedPageNumber)
    'sheetCount = ActiveDocument.BuiltInDocumentProperties("Number of Pages")
    sheetCount = ActiveDocument.Content.ComputeStatistics(wdStatisticPages)
    
    
    FillRegisterInWords sheetCount - 1
End Sub

' Заполнить всего страниц словами
Public Sub FillTotalInWords(totalSheets As Currency)
    Dim totalSheetsInWords_1Mark As String
    Dim totalSheetsInWords_2Mark As String
    
    totalSheetsInWords_1Mark = "TotalSheetsInWords_1"
    totalSheetsInWords_2Mark = "TotalSheetsInWords_2"
    
    Dim inWords As String
    
    inWords = NumberInWords(totalSheets)
    
    inWords = totalSheets & " (" & inWords & ")"
    
   '*********************************************
    Dim printer As C_PrinterMashine
    Set printer = New C_PrinterMashine
    
    printer.PrintByBookmark totalSheetsInWords_1Mark, inWords
    printer.PrintByBookmark totalSheetsInWords_2Mark, inWords
     
End Sub

' Заполнить страниц в описи словами
Public Sub FillRegisterInWords(totalSheets As Currency)
    Dim SheetsInWords_1Mark As String
    Dim SheetsInWords_2Mark As String
    
    SheetsInWords_1Mark = "RegisterSheetsInWords_1"
    SheetsInWords_2Mark = "RegisterSheetsInWords_2"
    
    Dim inWords As String
    
    inWords = NumberInWords(totalSheets)
    
    inWords = totalSheets & " (" & inWords & ")"
    
   '*********************************************
    Dim printer As C_PrinterMashine
    Set printer = New C_PrinterMashine
    
    printer.PrintByBookmark SheetsInWords_1Mark, inWords
    printer.PrintByBookmark SheetsInWords_2Mark, inWords
     
End Sub



Private Sub FillCellWithNumber(worker As C_TableWorker, record As C_RecordInfo)
    worker.PrintInCell record.docNumber
    worker.MoveRight 1
End Sub

Private Sub FillCellWithDate(worker As C_TableWorker, record As C_RecordInfo)
    'если дата 9999, то в документе нет даты и нужно поставить прочерк
    If record.DateFirst = 9999 Then
        worker.PrintInCell "-"
    'если дата одинаковая пишем только последнюю
    ElseIf record.DateFirst = record.dateLast And record.DateFirst <> 0 Then
        worker.PrintInCell record.dateLast
    Else
    ' Иначе пишем обе через тире
        If record.DateFirst <> 0 And record.dateLast <> 0 Then
            worker.PrintInCell record.DateFirst & "-" & record.dateLast
        ElseIf record.DateFirst <> 0 Then
            worker.PrintInCell record.DateFirst
        ElseIf record.dateLast <> 0 Then
            worker.PrintInCell record.dateLast
        End If
    End If
    worker.MoveRight 1
End Sub

Private Sub FillCellWithName(worker As C_TableWorker, record As C_RecordInfo)
    worker.PrintInCell record.docName
    worker.MoveRight 1
End Sub


Private Sub FillRow(i As Integer, record As C_RecordInfo)
      'Заполнить ячейку с номером документа
      Call FillCellWithNumber(worker, record)
      
    
      'Заполнить ячейку с датой
      Call FillCellWithDate(worker, record)
      
      ' Пишем название документа
      Call FillCellWithName(worker, record)
      
      ' Пишем страницы документа
      worker.PrintInCell record.sheetsNumber
End Sub
