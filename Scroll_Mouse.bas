Attribute VB_Name = "Scroll_Mouse"
'VeBAxceL
'https://www.youtube.com/channel/UCDd2VbwgxtsXtDGGWYY7vEQ

'Подпишитесь на каналы в мессенджерах, чтобы не пропустить свежие новости:
'ICQ - https://icq.im/VeBAxcel
'Telegram - https://t.me/VeBAxceL
'
'Выразите свою благодарность автору канала и поддержите дальнейшее развитие:
'Яндекс.Деньги
'https://money.yandex.ru/to/410012728519004

Option Explicit
Option Private Module

#If Win64 Then
    Private Type POINTAPI: XY As LongLong: End Type
#Else
    Private Type POINTAPI: X As Long: Y As Long: End Type
#End If
Private Type MOUSEHOOKSTRUCT: pt As POINTAPI: hwnd As Long: wHitTestCode As Long: dwExtraInfo As Long: End Type
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
    Private nMouseHook As LongPtr
#Else
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private nMouseHook As Long
#End If
Private oCtl As MSForms.Control

Sub HookScroll(Control As MSForms.Control)
UnHookScroll
Set oCtl = Control
#If VBA7 Then
    nMouseHook = SetWindowsHookEx(14, AddressOf MouseRotate_VBA7, 0, 0)
#Else
    nMouseHook = SetWindowsHookEx(14, AddressOf MouseRotate, 0, 0)
#End If
End Sub

Sub UnHookScroll()
Set oCtl = Nothing
UnhookWindowsHookEx nMouseHook
nMouseHook = 0
End Sub

#If VBA7 Then
    Private Function MouseRotate_VBA7(ByVal nCode As Long, ByVal wParam As LongPtr, ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
    Dim n&
    On Error GoTo Errr7
    If wParam = &H20A Then
        If lParam.hwnd > 0 Then n = -1 Else n = 1
        n = n + oCtl.TopIndex
        If n >= 0 Then oCtl.TopIndex = n: Exit Function
    End If
    MouseRotate_VBA7 = CallNextHookEx(nMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
Errr7:
    Err.Clear
    UnHookScroll
    End Function
#Else
    Private Function MouseRotate(ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As MOUSEHOOKSTRUCT) As Long
    Dim n&
    On Error GoTo Errr
    If wParam = &H20A Then
        If lParam.hwnd > 0 Then n = -1 Else n = 1
        n = n + oCtl.TopIndex
        If n >= 0 Then oCtl.TopIndex = n: Exit Function
    End If
    MouseRotate = CallNextHookEx(nMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
Errr:
    Err.Clear
    UnHookScroll
    End Function
#End If
