Attribute VB_Name = "SumaInCuirsive"
'Переводит число в написание словами
Public Function NumberInWords(Число As Currency) As String

'до 999 999 999 999

On Error GoTo Число_Error

Dim strМиллиарды As String, strМиллионы As String, strТысячи As String, strЕдиницы As String, strСотые As String

Dim Поз As Integer

 

strЧисло = Format(Int(Число), "000000000000")

 

'Миллиарды'

Поз = 1

strМиллиарды = Сотни(Mid(strЧисло, Поз, 1))

strМиллиарды = strМиллиарды & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

strМиллиарды = strМиллиарды & ИмяРазряда(strМиллиарды, Mid(strЧисло, Поз + 1, 2), "миллиард ", "миллиарда ", "миллиардов ")

 

'Миллионы'

Поз = 4

strМиллионы = Сотни(Mid(strЧисло, Поз, 1))

strМиллионы = strМиллионы & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

strМиллионы = strМиллионы & ИмяРазряда(strМиллионы, Mid(strЧисло, Поз + 1, 2), "миллион ", "миллиона ", "миллионов ")

 

'Тысячи'

Поз = 7

strТысячи = Сотни(Mid(strЧисло, Поз, 1))

strТысячи = strТысячи & Десятки(Mid(strЧисло, Поз + 1, 2), "ж")

strТысячи = strТысячи & ИмяРазряда(strТысячи, Mid(strЧисло, Поз + 1, 2), "тысяча ", "тысячи ", "тысяч ")

 

'Единицы'

Поз = 10

strЕдиницы = Сотни(Mid(strЧисло, Поз, 1))

strЕдиницы = strЕдиницы & Десятки(Mid(strЧисло, Поз + 1, 2), "м")

If strМиллиарды & strМиллионы & strТысячи & strЕдиницы = "" Then strЕдиницы = "ноль "

'strЕдиницы = strЕдиницы & ИмяРазряда(" ", Mid(strЧисло, Поз + 1, 2), "рубль ", "рубля ", "рублей ")

 

 

'Сотые'

'strСотые = strКопейки & " " & ИмяРазряда(strКопейки, Right(strКопейки, 2), ‘"копейка", "копейки", "копеек")

 

NumberInWords = strМиллиарды & strМиллионы & strТысячи & strЕдиницы

'NumberInWords = UCase(Left(NumberInWords, 1)) & Right(NumberInWords, Len(NumberInWords) - 1)

 
NumberInWords = Trim(NumberInWords)
Exit Function

 

Число_Error:

    MsgBox Err.Description
End Function

 

Function Сотни(n As String) As String

Сотни = ""

Select Case n

    Case 0: Сотни = ""

    Case 1: Сотни = "сто "

    Case 2: Сотни = "двести "

    Case 3: Сотни = "триста "

    Case 4: Сотни = "четыреста "

    Case 5: Сотни = "пятьсот "

    Case 6: Сотни = "шестьсот "

    Case 7: Сотни = "семьсот "

    Case 8: Сотни = "восемьсот "

    Case 9: Сотни = "девятьсот "

End Select

End Function

 

Function Десятки(n As String, Sex As String) As String

Десятки = ""

Select Case left(n, 1)

    Case "0": Десятки = "": n = Right(n, 1)

    Case "1": Десятки = ""

    Case "2": Десятки = "двадцать ": n = Right(n, 1)

    Case "3": Десятки = "тридцать ": n = Right(n, 1)

    Case "4": Десятки = "сорок ": n = Right(n, 1)

    Case "5": Десятки = "пятьдесят ": n = Right(n, 1)

    Case "6": Десятки = "шестьдесят ": n = Right(n, 1)

    Case "7": Десятки = "семьдесят ": n = Right(n, 1)

    Case "8": Десятки = "восемьдесят ": n = Right(n, 1)

    Case "9": Десятки = "девяносто ": n = Right(n, 1)

End Select

 

Dim Двадцатка As String

Двадцатка = ""

Select Case n

    Case "0": Двадцатка = ""

    Case "1"

        Select Case Sex

            Case "м": Двадцатка = "один "

            Case "ж": Двадцатка = "одна "

            Case "с": Двадцатка = "одно "

        End Select

    Case "2":

        Select Case Sex

            Case "м": Двадцатка = "два "

            Case "ж": Двадцатка = "две "

            Case "с": Двадцатка = "два "

        End Select

    Case "3": Двадцатка = "три "

    Case "4": Двадцатка = "четыре "

    Case "5": Двадцатка = "пять "

    Case "6": Двадцатка = "шесть "

    Case "7": Двадцатка = "семь "

    Case "8": Двадцатка = "восемь "

    Case "9": Двадцатка = "девять "

    Case "10": Двадцатка = "десять "

    Case "11": Двадцатка = "одиннадцать "

    Case "12": Двадцатка = "двенадцать "

    Case "13": Двадцатка = "тринадцать "

    Case "14": Двадцатка = "четырнадцать "

    Case "15": Двадцатка = "пятнадцать "

    Case "16": Двадцатка = "шестнадцать "

    Case "17": Двадцатка = "семнадцать "

    Case "18": Двадцатка = "восемнадцать "

    Case "19": Двадцатка = "девятнадцать "

End Select

 

Десятки = Десятки & Двадцатка

End Function

 

Function ИмяРазряда(Строка As String, n As String, Имя1 As String, Имя24 As String, ИмяПроч As String) As String

 

If Строка <> "" Then

    ИмяРазряда = ""

    Select Case left(n, 1)

        Case "0", "2", "3", "4", "5", "6", "7", "8", "9": n = Right(n, 1)

    End Select

 

    Select Case n

        Case "1": ИмяРазряда = Имя1

        Case "2", "3", "4": ИмяРазряда = Имя24

        Case Else: ИмяРазряда = ИмяПроч

    End Select

End If

 

End Function
