VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_TableWorker"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Public Sub MoveRight(cellCount As Integer)
     Selection.MoveRight Unit:=wdCell, count:=cellCount
End Sub

Public Sub MoveLeft(cellCount As Integer)
     Selection.MoveLeft Unit:=wdCell, count:=cellCount
End Sub

Public Sub MoveDown(cellCount As Integer)
     Selection.MoveDown Unit:=wdLine, count:=cellCount
End Sub

Public Sub PrintInCell(text As String)
    Selection.TypeText text:=text
End Sub

Public Sub MoveToBookmark(bookmarkName As String)
    '”становить начальную €чейку активной
    ActiveDocument.Bookmarks(bookmarkName).Select
End Sub
