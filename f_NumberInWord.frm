VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_NumberInWord 
   OleObjectBlob   =   "f_NumberInWord.frx":0000
   Caption         =   "UserForm1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   7
End
Attribute VB_Name = "f_NumberInWord"
Attribute VB_Base = "0{E663B658-4787-4AEF-BCC2-656363D5C407}{8975596E-33F3-44A8-8A83-C1F84A0B1857}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
 Dim inputText As Currency
   inputText = Int(tb_Input.text)
   tb_Output.text = inputText & " (" & NumberInWords(inputText) & ") "
End Sub
