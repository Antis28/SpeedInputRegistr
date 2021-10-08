Attribute VB_Name = "WorkWithTable"
Dim AddNum As Integer

Public Sub AllFirstRowPensDelo()
' ������ �������� "���������� ����" � �����
    Dim i As Integer
    For i = 1 To 200 Step 1
        Selection.Move Unit:=wdRow, count:=9
        FirstRowPensDelo
    Next i
End Sub

Public Sub FirstRowPensDelo()
' ������ �������� "���������� ����" � 3-� �������
    Selection.Move Unit:=wdCell, count:=2
    Selection.MoveRight Unit:=wdWord, count:=3, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText text:="���������� ����" '& vbCrLf
End Sub

Public Sub AllTableHeader()
' ������ �������� ����� ������� � ����� ����� 8 �����
    Dim i As Integer
    For i = 1 To 200 Step 1
        Selection.Move Unit:=wdRow, count:=8
        InsertHeaderTable
        
    Next i
End Sub
Public Sub DelAllTableHeader()
' ������ ������� ����� ������� � �����
    Dim i As Integer
    For i = 1 To 200 Step 1
        Selection.Move Unit:=wdRow, count:=8
        Selection.Rows.Delete
    Next i
End Sub
Sub InsertHeaderTable()
Attribute InsertHeaderTable.VB_Description = "������ ������� 28.07.2017 PC_0101_10"
Attribute InsertHeaderTable.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������1"
'
' ������ �������� ����� �������
    Selection.InsertRowsAbove 1
'    Selection.Paste
    pasteNum
    Selection.Rows.Height = CentimetersToPoints(0.5)
    
    
    Selection.MoveDown Unit:=wdLine, count:=1
    FirstRowPensDelo
'
'    Selection.MoveUp Unit:=wdLine, Count:=2
'    Selection.Rows.Height = CentimetersToPoints(4.5)
    
End Sub



Sub pasteNum()
    Selection.MoveLeft Unit:=wdCharacter, count:=1
    Selection.TypeText text:="1"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="2"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="3"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="4"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="5"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="6"
    Selection.MoveRight Unit:=wdCharacter, count:=1
    Selection.TypeText text:="7"
    Selection.MoveLeft Unit:=wdCharacter, count:=7, Extend:=wdExtend
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Bold = wdToggle
    Selection.Font.Color = wdColorAutomatic

End Sub

Sub LineNum()
Attribute LineNum.VB_Description = "������ ������� 28.07.2017 PC_0101_10"
Attribute LineNum.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.������2"
'
' ������2 �������� �� �������
' 131 - ������
'
    Dim i As Integer, a As Integer
    AddNum = 4
    For i = 1 To 200 Step 1
        For a = AddNum To AddNum + 7 Step 1
            Selection.SelectCell
            Selection.Delete Unit:=wdCharacter, count:=1
            Selection.TypeText text:=a
            Selection.MoveDown Unit:=wdLine, count:=1
        Next a
        AddNum = a
'        Selection.SelectCell
'        Selection.Delete unit:=wdCharacter, Count:=1
'        Selection.TypeText Text:=1
'        Selection.MoveDown unit:=wdLine, Count:=1
    Next i
End Sub

Sub ������3()
Attribute ������3.VB_Description = "������ ������� 02.08.2017 PC"
Attribute ������3.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.������3"
'
' ������3 ������
' ������ ������� 02.08.2017 PC
'
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.MoveUp Unit:=wdLine, count:=1
    Selection.MoveRight Unit:=wdCharacter, count:=5
    Selection.MoveLeft Unit:=wdCharacter, count:=1, Extend:=wdExtend
    Selection.HomeKey Unit:=wdLine
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.MoveDown Unit:=wdLine, count:=1
    Selection.MoveUp Unit:=wdLine, count:=1
End Sub


Sub ������6()
Attribute ������6.VB_Description = "������ ������� 03.08.2017 PC"
Attribute ������6.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.������6"
'
' ������6 ������
' ������ ������� 03.08.2017 PC
'
    For i = 1 To 4 Step 1
        For a = 1 To 8 Step 1
            Selection.HomeKey Unit:=wdLine
            Selection.EndKey Unit:=wdLine
            Selection.MoveRight Unit:=wdWord, count:=1, Extend:=wdExtend
            Selection.TypeText text:="02-11"
            Selection.MoveDown Unit:=wdLine, count:=1
        Next a
        Selection.MoveDown Unit:=wdLine, count:=1
    Next i
End Sub

Sub DelRow()
Attribute DelRow.VB_Description = "������ ������� 08.08.2017 PC"
Attribute DelRow.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.DelRow"
'
' DelRow ������
' ������ ������� 08.08.2017 PC
'
    Selection.Rows.Delete
End Sub
'**************************************************************************
Sub P_�����(value As String)
'
' ������
' ������ ������� 02.08.2017 PC
'
    Selection.HomeKey Unit:=wdLine
    Selection.MoveRight Unit:=wdWord, count:=5, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, count:=1
    Selection.TypeText text:=value
End Sub

Sub P_������()
    P_����� "31.01.2015"
End Sub
Sub P_�������()
    P_����� "28.02.2014"
End Sub
Sub P_����()
    P_����� "31.03.2014"
End Sub
Sub P_������()
    P_����� "30.04.2015"
End Sub
Sub P_���()
    P_����� "31.05.2015"
End Sub
Sub P_����()
    P_����� "30.06.2014"
End Sub
Sub P_����()
    P_����� "31.07.2014"
End Sub
Sub P_������()
    P_����� "31.08.2015"
End Sub
Sub P_��������()
    P_����� "30.09.2015"
End Sub
Sub P_�������()
    P_����� "31.10.2015"
End Sub
Sub P_������()
    P_����� "30.11.2015"
End Sub
'*********************************************
Sub P_������_2()
    Selection.TypeText text:="������ " & ".01.2015"
End Sub
Sub P_�������_2()
    Selection.TypeText text:="������ " & ".02.2015"
End Sub
Sub P_����_2()
    Selection.TypeText text:="������ " & ".03.2015"
End Sub
Sub P_������_2()
    Selection.TypeText text:="������ " & ".04.2015"
End Sub
Sub P_���_2()
    Selection.TypeText text:="������ " & ".05.2015"
End Sub
Sub P_����_2()
    Selection.TypeText text:="������ " & ".06.2015"
End Sub
Sub P_����_2()
    Selection.TypeText text:="������ " & ".07.2015"
End Sub
Sub P_������_2()
    Selection.TypeText text:="������ " & ".08.2015"
End Sub

Sub MyInsertRow(count As Integer, isAbove As Boolean)
Attribute MyInsertRow.VB_Description = "������ ������� 25.08.2017 PC_309_06"
Attribute MyInsertRow.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.������1"
'
' ��������� � �����
' ������ ������� 25.08.2017 PC_309_06
'
    For X = 1 To count
        If isAbove Then
            Selection.InsertRowsAbove 1
            'Selection.MoveUp Unit:=wdLine, count:=1
        Else
            Selection.InsertRowsBelow 1
            'Selection.MoveUp Unit:=wdLine, count:=1
            'Selection.MoveDown Unit:=wdLine, count:=1
        End If
        
        'Selection.MoveRight Unit:=wdCharacter, count:=1
       ' Selection.TypeText Text:="02-11"
     '   Selection.MoveRight Unit:=wdCharacter, count:=1
     '   Selection.TypeText Text:="�� �� �����"
     '   Selection.MoveRight Unit:=wdCharacter, count:=3
     '   Selection.TypeText Text:="25 ���"
    Next X
End Sub

Public Sub InsertRowToForm()
    FormInsertRow.Show
End Sub
