Attribute VB_Name = "TestTable"
Option Explicit


Public Sub TestDataInsert()
   
    Selection.TypeText text:="1"
    
    Selection.MoveLeft Unit:=wdCell, count:=1
    Selection.MoveDown Unit:=wdLine, count:=1
    
    Selection.TypeText text:="1"
End Sub
