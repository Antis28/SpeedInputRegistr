VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArraySorter"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private LeftMoreRight As String
Private LeftLessRight As String
Private CompareEquels As String

Public Sub sort(ByRef a() As Long)
    'QuickSort a, LBound(a), UBound(a)
    InsertionSort a, LBound(a), UBound(a)
End Sub

Public Sub sortCollection(ByRef coll As Collection, LMR As String, LLR As String, CE As String)
    LeftMoreRight = LMR
    LeftLessRight = LLR
    CompareEquels = CE


    Dim a() As C_RecordInfo
    a = CollectionToArray(coll)
    QuickSort a, LBound(a), UBound(a)
    InsertionSort a, LBound(a), UBound(a)
    Set coll = ArrayToCollection(a)
End Sub

Private Sub QuickSort(ByRef a As Variant, ByVal l As Long, ByVal r As Long)
    Dim M As Long, _
    i As Long, _
    j As Long, _
    v As C_RecordInfo
    
    Dim recordHigh As C_RecordInfo
    Dim recordLow As C_RecordInfo
    
    M = 4

    If ((r - l) > M) Then
        i = (r + l) / 2
       ' If (a(l) > a(i)) Then swap a, l, i '// Tri-Median Methode!'
        If (Application.Run(LeftMoreRight, a(l), a(i))) Then swap a, l, i
        
        'If (a(l) > a(r)) Then swap a, l, r
        If (Application.Run(LeftMoreRight, a(l), a(r))) Then swap a, l, r
         
        'If (a(i) > a(r)) Then swap a, i, r
        If (Application.Run(LeftMoreRight, a(i), a(r))) Then swap a, i, r
        
        
        j = r - 1
        
        swap a, i, j
        
        i = l
        Set v = a(j)
        Do
            'Do: i = i + 1: Loop While (a(i) < v)
            Do: i = i + 1: Loop While Application.Run(LeftLessRight, a(i), v)
            'Do: j = j - 1: Loop While (a(j) > v)
            Do: j = j - 1: Loop While Application.Run(LeftMoreRight, a(j), v)
            If (j < i) Then Exit Do
            swap a, i, j
        Loop
        
        swap a, i, r - 1
        
        QuickSort a, l, j
        QuickSort a, i + 1, r
    End If
End Sub

Private Sub swap(ByRef a As Variant, ByVal i As Long, ByVal j As Long)
    
    If Application.Run(CompareEquels, a(i), a(j)) Then
        Exit Sub
    End If

    Dim T As C_RecordInfo
    Set T = a(i)
    Set a(i) = a(j)
    Set a(j) = T
End Sub

Private Sub InsertionSort(ByRef a As Variant, ByVal lo0 As Long, ByVal hi0 As Long)
    Dim i As Long, j As Long, v As C_RecordInfo

    For i = lo0 + 1 To hi0
        Set v = a(i)
        j = i
        ' Сдвигаем элемент пока он не станет на свое место
        Do While j > lo0
            'If Not a(j - 1) > v Then Exit Do
            If Not Application.Run(LeftMoreRight, a(j - 1), v) Then Exit Do
            Set a(j) = a(j - 1)
            j = j - 1
        Loop
        Set a(j) = v
    Next i
End Sub

' Конвертация в массив
Private Function CollectionToArray(c As Collection) As C_RecordInfo()
    Dim a() As C_RecordInfo: ReDim a(0 To c.count - 1)
    Dim i As Integer
    For i = 1 To c.count
       Set a(i - 1) = c.item(i)
    Next
    CollectionToArray = a
End Function

' Конвертация в коллекцию
Private Function ArrayToCollection(a As Variant) As Collection
    Dim c As New Collection
    Dim i As Integer
    For i = 0 To UBound(a)
       c.Add a(i)
    Next
    Set ArrayToCollection = c
End Function

















Private Function QuickSortOLD(vArray As Variant, inLow As Long, inHi As Long, ByVal CompareMethod As String)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  
  Dim recordHigh As C_RecordInfo
  Dim recordLow As C_RecordInfo


  tmpLow = inLow
  tmpHi = inHi
  
  Dim medium As Integer
  medium = (inLow + inHi) \ 2

  Set pivot = vArray(medium)

  While (tmpLow <= tmpHi)
   Set recordHigh = vArray(tmpHi)
   Set recordLow = vArray(tmpLow)
   
   Dim recordLessPivot As Boolean
   
    recordLessPivot = Application.Run(CompareMethod, recordLow, pivot)
    While (recordLessPivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
        Set recordLow = vArray(tmpLow)
        recordLessPivot = Application.Run(CompareMethod, recordLow, pivot)
    Wend

     While (Application.Run(CompareMethod, pivot, recordHigh) _
            And tmpHi > inLow)
        tmpHi = tmpHi - 1
        Set recordHigh = vArray(tmpHi)
     Wend

     If (tmpLow <= tmpHi) Then
        Set tmpSwap = vArray(tmpLow)
        Set vArray(tmpLow) = vArray(tmpHi)
        Set vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi, CompareMethod
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi, CompareMethod

End Function

 ' Сортировка массива строк по алфавиту (в порядке возрастания)
Sub SortArrByAlpavit(astrNames As Variant)
    Dim i As Integer, j As Integer
    Dim strBuffer As String, fBuffer As Boolean
    For i = LBound(astrNames) To UBound(astrNames) - 1
        For j = i + 1 To UBound(astrNames)
            If astrNames(i) > astrNames(j) Then    ' Меняем i-й и j-й элементы массива местами
                strBuffer = astrNames(i): astrNames(i) = astrNames(j): astrNames(j) = strBuffer
            End If
        Next j
    Next i
End Sub
