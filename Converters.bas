Attribute VB_Name = "Converters"
Option Explicit

Public Function QuickSort(vArray As Variant, inLow As Long, inHi As Long)

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

     While (recordLow.dateLast < pivot.dateLast And tmpLow < inHi)
        tmpLow = tmpLow + 1
        Set recordLow = vArray(tmpLow)
     Wend

     While (pivot.dateLast < recordHigh.dateLast And tmpHi > inLow)
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

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Function

Public Sub tttt()
    Dim arr As New Collection
    
    Dim rec As C_RecordInfo
    Set rec = New C_RecordInfo
    
    rec.Construct 0, 1993, "test_1"
    arr.Add rec
    
    Set rec = New C_RecordInfo
    rec.Construct 0, 1992, "test_2"
    arr.Add rec
    
    Set rec = New C_RecordInfo
    rec.Construct 0, 1991, "test_3"
    arr.Add rec
    
    Set rec = New C_RecordInfo
    rec.Construct 0, 1990, "test_4"
    arr.Add rec
    
     Set rec = New C_RecordInfo
    rec.Construct 0, 1994, "test_5"
    arr.Add rec
    
    Set rec = New C_RecordInfo
    rec.Construct 0, 1995, "test_6"
    arr.Add rec
    Dim a() As Variant
    
    a = CollectionToArray(arr)
    QuickSort a, 0, UBound(a)
    Set arr = ArrayToCollection(a)
    
End Sub


Public Function CollectionToArray(c As Collection) As Variant()
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Integer
    For i = 1 To c.count
        If IsObject(c.item(i)) Then
            Set a(i - 1) = c.item(i)
        Else
            a(i - 1) = c.item(i)
        End If
    Next
    CollectionToArray = a
End Function

Public Function ArrayToCollection(a As Variant) As Collection
    Dim c As New Collection
    Dim i As Integer
    For i = 0 To UBound(a)
       c.Add a(i)
    Next
    Set ArrayToCollection = c
End Function
