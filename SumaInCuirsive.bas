Attribute VB_Name = "SumaInCuirsive"
'��������� ����� � ��������� �������
Public Function NumberInWords(����� As Currency) As String

'�� 999 999 999 999

On Error GoTo �����_Error

Dim str��������� As String, str�������� As String, str������ As String, str������� As String, str����� As String

Dim ��� As Integer

 

str����� = Format(Int(�����), "000000000000")

 

'���������'

��� = 1

str��������� = �����(Mid(str�����, ���, 1))

str��������� = str��������� & �������(Mid(str�����, ��� + 1, 2), "�")

str��������� = str��������� & ����������(str���������, Mid(str�����, ��� + 1, 2), "�������� ", "��������� ", "���������� ")

 

'��������'

��� = 4

str�������� = �����(Mid(str�����, ���, 1))

str�������� = str�������� & �������(Mid(str�����, ��� + 1, 2), "�")

str�������� = str�������� & ����������(str��������, Mid(str�����, ��� + 1, 2), "������� ", "�������� ", "��������� ")

 

'������'

��� = 7

str������ = �����(Mid(str�����, ���, 1))

str������ = str������ & �������(Mid(str�����, ��� + 1, 2), "�")

str������ = str������ & ����������(str������, Mid(str�����, ��� + 1, 2), "������ ", "������ ", "����� ")

 

'�������'

��� = 10

str������� = �����(Mid(str�����, ���, 1))

str������� = str������� & �������(Mid(str�����, ��� + 1, 2), "�")

If str��������� & str�������� & str������ & str������� = "" Then str������� = "���� "

'str������� = str������� & ����������(" ", Mid(str�����, ��� + 1, 2), "����� ", "����� ", "������ ")

 

 

'�����'

'str����� = str������� & " " & ����������(str�������, Right(str�������, 2), �"�������", "�������", "������")

 

NumberInWords = str��������� & str�������� & str������ & str�������

'NumberInWords = UCase(Left(NumberInWords, 1)) & Right(NumberInWords, Len(NumberInWords) - 1)

 
NumberInWords = Trim(NumberInWords)
Exit Function

 

�����_Error:

    MsgBox Err.Description
End Function

 

Function �����(n As String) As String

����� = ""

Select Case n

    Case 0: ����� = ""

    Case 1: ����� = "��� "

    Case 2: ����� = "������ "

    Case 3: ����� = "������ "

    Case 4: ����� = "��������� "

    Case 5: ����� = "������� "

    Case 6: ����� = "�������� "

    Case 7: ����� = "������� "

    Case 8: ����� = "��������� "

    Case 9: ����� = "��������� "

End Select

End Function

 

Function �������(n As String, Sex As String) As String

������� = ""

Select Case left(n, 1)

    Case "0": ������� = "": n = Right(n, 1)

    Case "1": ������� = ""

    Case "2": ������� = "�������� ": n = Right(n, 1)

    Case "3": ������� = "�������� ": n = Right(n, 1)

    Case "4": ������� = "����� ": n = Right(n, 1)

    Case "5": ������� = "��������� ": n = Right(n, 1)

    Case "6": ������� = "���������� ": n = Right(n, 1)

    Case "7": ������� = "��������� ": n = Right(n, 1)

    Case "8": ������� = "����������� ": n = Right(n, 1)

    Case "9": ������� = "��������� ": n = Right(n, 1)

End Select

 

Dim ��������� As String

��������� = ""

Select Case n

    Case "0": ��������� = ""

    Case "1"

        Select Case Sex

            Case "�": ��������� = "���� "

            Case "�": ��������� = "���� "

            Case "�": ��������� = "���� "

        End Select

    Case "2":

        Select Case Sex

            Case "�": ��������� = "��� "

            Case "�": ��������� = "��� "

            Case "�": ��������� = "��� "

        End Select

    Case "3": ��������� = "��� "

    Case "4": ��������� = "������ "

    Case "5": ��������� = "���� "

    Case "6": ��������� = "����� "

    Case "7": ��������� = "���� "

    Case "8": ��������� = "������ "

    Case "9": ��������� = "������ "

    Case "10": ��������� = "������ "

    Case "11": ��������� = "����������� "

    Case "12": ��������� = "���������� "

    Case "13": ��������� = "���������� "

    Case "14": ��������� = "������������ "

    Case "15": ��������� = "���������� "

    Case "16": ��������� = "����������� "

    Case "17": ��������� = "���������� "

    Case "18": ��������� = "������������ "

    Case "19": ��������� = "������������ "

End Select

 

������� = ������� & ���������

End Function

 

Function ����������(������ As String, n As String, ���1 As String, ���24 As String, ������� As String) As String

 

If ������ <> "" Then

    ���������� = ""

    Select Case left(n, 1)

        Case "0", "2", "3", "4", "5", "6", "7", "8", "9": n = Right(n, 1)

    End Select

 

    Select Case n

        Case "1": ���������� = ���1

        Case "2", "3", "4": ���������� = ���24

        Case Else: ���������� = �������

    End Select

End If

 

End Function
