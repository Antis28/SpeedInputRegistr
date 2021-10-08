VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_RegisterInfo"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public count As Integer

Private id As Integer
Private recordsList As Collection

Sub Class_Initialize()
    Set recordsList = New Collection
End Sub

Public Sub Add(record As C_RecordInfo)
    recordsList.Add record
    count = recordsList.count
End Sub

Public Sub Remove(id As Integer)
    recordsList.Remove id
    count = recordsList.count
End Sub

Public Function getCollection()
    Set getCollection = recordsList
End Function

Public Function setCollection(coll As Collection)
    Set recordsList = coll
    count = recordsList.count
End Function

Public Function getRecord(i As Integer) As C_RecordInfo
   Set getRecord = recordsList(i)
End Function

Public Function getPeriod() As String
    If recordsList.count = 0 Then
        getPeriod = 0
        Exit Function
    End If
    
    Dim firstYear As Integer
    Dim lastYear As Integer
    Dim item As C_RecordInfo
    
    firstYear = 9999
    lastYear = 0
    For Each item In recordsList
        If item.DateFirst = "9999" Then
            GoTo nextItem
        End If
    
        If item.DateFirst < firstYear Then
            firstYear = item.DateFirst
        End If
        If item.dateLast > lastYear Then
            lastYear = item.dateLast
        End If
nextItem:
    Next
    
    getPeriod = firstYear & " - " & lastYear
End Function

Public Function getPeriod_old() As String
    If recordsList.count = 0 Then
        getPeriod = 0
        Exit Function
    End If
    
    Dim firstYear As Integer
    Dim lastYear As Integer
    firstYear = recordsList(1).DateFirst
    lastYear = recordsList(recordsList.count).dateLast
    getPeriod = firstYear & " - " & lastYear
End Function

Public Sub SortByLastDate()
    ' sort "CompareDateLast", 0
    Dim sorter As New ArraySorter
    sorter.sortCollection recordsList, "LeftMoreRightDateLast", "LeftLessRightDateLast", "CompareEqualsDateLast"
End Sub


Public Sub SortByTwoParameters()
    SortByLastDate
    Dim divColl As Collection
    Set divColl = FilterByLastDate
    
     Set divColl = FilterByLastDate()
     Set divColl = SortByFirstDateFromColl(divColl)
     Set recordsList = MergeColl(divColl)
              
End Sub


' Разбиваем массив на отдельные массивы по 2-й дате
Private Function FilterByLastDate() As Collection
    Dim eq As New Collection, _
    Element As C_RecordInfo, _
    elementPivot As C_RecordInfo, _
    id As Integer
    
    Dim currColl As Collection
    Set currColl = New Collection
    id = 1
    Do While id <= recordsList.count
        Set currColl = New Collection
        Set Element = recordsList(id)
        Set elementPivot = recordsList(id)
        
        Do While Element.dateLast = elementPivot.dateLast
            currColl.Add Element
            id = id + 1
            If id > recordsList.count Then
                Exit Do
            End If
            Set Element = recordsList(id)
        Loop
        If currColl.count > 0 Then
            eq.Add currColl
        End If
    Loop
    
    Set FilterByLastDate = eq
End Function

' Сортирует коллекцию коллекций по 1-й дате после разделения  в FilterByLastDate
Private Function SortByFirstDateFromColl(ByRef coll As Collection) As Collection
    Dim sorter As New ArraySorter, _
    newColl As New Collection, _
    newItemRec As Collection
    
    For id = 1 To coll.count
        Set newItemRec = coll(id)
        sorter.sortCollection newItemRec, "LeftMoreRightDateFirst", "LeftLessRightDateFirst", "CompareEqualsDateFirst"
        newColl.Add newItemRec
    Next id
    
    Set SortByFirstDateFromColl = newColl
End Function

Private Function MergeColl(ByRef coll As Collection) As Collection
    Dim allColl As New Collection, _
    innerColl As Collection, _
    endItem As C_RecordInfo
    
    For Each innerColl In coll
        For Each endItem In innerColl
            allColl.Add endItem
        Next endItem
    Next innerColl
    Set MergeColl = allColl
End Function

' Есть ли уже в коллекции такой элемент
'Private Function CollExist(coll As Collection, cover As C_RecordInfo) As Boolean
'    Dim result As Boolean
'    Dim Item As C_RecordInfo
'    result = False
'    For Each Item In coll
'        If cover.id = Item.id Then
'           result = True
'           GoTo EndFor
'        End If
'    Next Item
'EndFor:
'    CollExist = result
'End Function

'Public Sub SortByFirstDate()
'    'sort "CompareDateFirst", 0
'    Dim sorter As New ArraySorter
'    sorter.sortCollection recordsList, "LeftMoreRightDateLast", "LeftLessRightDateLast"
'End Sub

'Private Sub sort(fieldBySort As String, startIndex As Long)
'    Dim a() As Variant
'    a = CollectionToArray(recordsList)
'    QuickSort a, startIndex, UBound(a), fieldBySort
'    Set recordsList = ArrayToCollection(a)
'End Sub

'Private Function QuickSortLastDate(vArray As Variant, inLow As Long, inHi As Long)
'
'  Dim pivot   As Variant
'  Dim tmpSwap As Variant
'  Dim tmpLow  As Long
'  Dim tmpHi   As Long
'
'  Dim recordHigh As C_RecordInfo
'  Dim recordLow As C_RecordInfo
'
'
'  tmpLow = inLow
'  tmpHi = inHi
'
'  Dim medium As Integer
'  medium = (inLow + inHi) \ 2
'
'  Set pivot = vArray(medium)
'
'  While (tmpLow <= tmpHi)
'   Set recordHigh = vArray(tmpHi)
'   Set recordLow = vArray(tmpLow)
'
'     While (recordLow.DateLast < pivot.DateLast And tmpLow < inHi)
'        tmpLow = tmpLow + 1
'        Set recordLow = vArray(tmpLow)
'     Wend
'
'     While (pivot.DateLast < recordHigh.DateLast And tmpHi > inLow)
'        tmpHi = tmpHi - 1
'        Set recordHigh = vArray(tmpHi)
'     Wend
'
'     If (tmpLow <= tmpHi) Then
'        Set tmpSwap = vArray(tmpLow)
'        Set vArray(tmpLow) = vArray(tmpHi)
'        Set vArray(tmpHi) = tmpSwap
'
'        tmpLow = tmpLow + 1
'        tmpHi = tmpHi - 1
'     End If
'
'  Wend
'
'  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
'  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
'
'End Function

' CompareMethod - callback который определяет поле для сравнения. См. CompareDateFirst, CompareDateLast, Application.Run(CompareMethod, recordLow, pivot)
'Private Function QuickSort(vArray As Variant, inLow As Long, inHi As Long, ByVal CompareMethod As String)
'
'  Dim pivot   As Variant
'  Dim tmpSwap As Variant
'  Dim tmpLow  As Long
'  Dim tmpHi   As Long
'
'  Dim recordHigh As C_RecordInfo
'  Dim recordLow As C_RecordInfo
'
'
'  tmpLow = inLow
'  tmpHi = inHi
'
'  Dim medium As Integer
'  medium = (inLow + inHi) \ 2
'
'  Set pivot = vArray(medium)
'
'  While (tmpLow <= tmpHi)
'   Set recordHigh = vArray(tmpHi)
'   Set recordLow = vArray(tmpLow)
'
'   Dim recordLessPivot As Boolean
'
'    recordLessPivot = Application.Run(CompareMethod, recordLow, pivot)
'    While (recordLessPivot And tmpLow < inHi)
'        tmpLow = tmpLow + 1
'        Set recordLow = vArray(tmpLow)
'        recordLessPivot = Application.Run(CompareMethod, recordLow, pivot)
'    Wend
'
'     While (Application.Run(CompareMethod, pivot, recordHigh) _
'            And tmpHi > inLow)
'        tmpHi = tmpHi - 1
'        Set recordHigh = vArray(tmpHi)
'     Wend
'
'     If (tmpLow <= tmpHi) Then
'        Set tmpSwap = vArray(tmpLow)
'        Set vArray(tmpLow) = vArray(tmpHi)
'        Set vArray(tmpHi) = tmpSwap
'        tmpLow = tmpLow + 1
'        tmpHi = tmpHi - 1
'     End If
'
'  Wend
'
'  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi, CompareMethod
'  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi, CompareMethod
'
'End Function

'Private Sub PrimitiveSort()
'    Dim newColl As Collection
'    Set newColl = New Collection
'
'    Dim i As Integer
'    i = 1
'    Do While recordsList.count > 0
'        Dim itemCurrent As C_RecordInfo
'        Dim itemMinimum As C_RecordInfo
'
'        Dim idMinimum As Integer
'
'        Set itemMinimum = recordsList(i)
'        idMinimum = i
'
'        Dim j As Integer
'        For j = 1 To recordsList.count
'             Set itemCurrent = recordsList(j)
'            If itemCurrent.DateLast < itemMinimum.DateLast Then
'               Set itemMinimum = itemCurrent
'               idMinimum = j
'            End If
'        Next j
'        recordsList.Remove idMinimum
'        newColl.Add itemMinimum
'    Loop
'    Set recordsList = newColl
'End Sub

Public Sub swap(indexA As Integer, indexB As Integer)
    Dim tmpSwap As Variant
    Dim tmpaArray As Variant
    tmpaArray = CollectionToArray(recordsList)

    Set tmpSwap = tmpaArray(indexA)
    Set tmpaArray(indexA) = tmpaArray(indexB)
    Set tmpaArray(indexB) = tmpSwap
    Set recordsList = ArrayToCollection(tmpaArray)
End Sub


Public Function GetSheetCount()
    Dim i As Integer
    Dim sheetSumma As Integer
    Dim item As C_RecordInfo
    
    For i = 1 To recordsList.count
        Set item = recordsList(i)
        sheetSumma = sheetSumma + item.SheetsCount
    Next
    
    GetSheetCount = sheetSumma
End Function
