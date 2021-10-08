Attribute VB_Name = "Module1"

Public Sub iiii()
    Dim test1 As XML_Worker
    Set test1 = New XML_Worker
    test1.Init "test"
    test1.Load
    Dim c As C_CoverInfo
    Set c = New C_CoverInfo
    c.Construct 1, "r", "21212", "1221", "11"
   
End Sub

Public Sub dfdsf()
  Dim test As Variant
  test = ActiveDocument.Content.Information(wdActiveEndAdjustedPageNumber)
End Sub
