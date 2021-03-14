VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Parametry 
   Caption         =   "Parametry"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2385
   OleObjectBlob   =   "Parametry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Parametry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
  Slabik.AddItem "1"
  Slabik.AddItem "2"
  Slabik.AddItem "3"
  Slabik = "1"
  Rod.AddItem "Muž"
  Rod.AddItem "Žena"
  Rod.AddItem "Mìsto"
  Rod = "Muž"
End Sub

Private Sub Jmen_Exit(ByVal Storno As MSForms.ReturnBoolean)
  If Not IsNumeric(Jmen) Then
    MsgBox "Poèet jmen musí být èíslo!!!", vbCritical, "Parametry"
    Storno = True
  ElseIf Jmen <= 0 Then
    MsgBox "Poèet jmen musí být kladný!!!", vbCritical, "Parametry"
    Storno = True
  Else
    Jmen = Int(Jmen)
  End If
End Sub

Private Sub CommandButton1_Click()
  Me.Hide
End Sub
