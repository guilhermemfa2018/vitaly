VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Dim mo_Events As Collection
 
Private Sub Form_Load()
    Dim i&
    Dim Btn As CommandButton
    Set mo_Events = New Collection
    For i = 1 To 3
        Set Btn = Me.Controls.Add("VB.CommandButton", "Cmd_" & i)
        Btn.Move 0, 360 * (i - 1), 3600, 360
        Btn.Visible = True
        Btn.Caption = "Teste" & i
        mo_Events.Add New cEvents
        mo_Events(i).Add_CommandButton Btn, i
    Next
End Sub
 
Public Sub ButtonClick(p_idx As Long)
    MsgBox "Button is clicked # " & p_idx

End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set mo_Events = Nothing
End Sub
