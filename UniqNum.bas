Attribute VB_Name = "UniqNum"
Option Explicit

Private uniqNumbers As New Collection



Public Function GetUniqNumber() As Integer
    GetUniqNumber = uniqNumbers(1)
    uniqNumbers.Remove (1)
End Function

Public Sub InitUniqNumbers()
    Dim i As Integer
    
    For i = 1 To 30000
       uniqNumbers.Add (i)
    Next i
End Sub
