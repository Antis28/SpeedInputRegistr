Attribute VB_Name = "CompareVariants"
Option Explicit

' Callbacks для сортировки масивов.
' Не работают при вызове из класса

Public Function CompareDateFirst(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    CompareDateFirst = (a.DateFirst < B.DateFirst)
End Function

Public Function CompareDateLast(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    CompareDateLast = (a.dateLast < B.dateLast)
End Function

Public Function CompareEqualsDateLast(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    CompareEqualsDateLast = (a.dateLast = B.dateLast)
End Function

Public Function CompareEqualsDateFirst(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    CompareEqualsDateFirst = (a.DateFirst = B.DateFirst)
End Function




Public Function LeftMoreRightDateLast(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    LeftMoreRightDateLast = (a.dateLast > B.dateLast)
End Function

Public Function LeftLessRightDateLast(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    LeftLessRightDateLast = (a.dateLast < B.dateLast)
End Function


Public Function LeftMoreRightDateFirst(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    LeftMoreRightDateFirst = (a.DateFirst > B.DateFirst)
End Function

Public Function LeftLessRightDateFirst(ByRef a As C_RecordInfo, ByRef B As C_RecordInfo) As Boolean
    LeftLessRightDateFirst = (a.DateFirst < B.DateFirst)
End Function
