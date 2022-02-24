VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Declare object variable as CommandButton and handle the events.'
Private WithEvents cmdObject As CommandButton
Attribute cmdObject.VB_VarHelpID = -1

Private Sub Form_Load()
   'Add button control and keep a reference in the WithEvents variable'
   Set cmdObject = Form3.Controls.Add("VB.CommandButton", "cmdOne")
   cmdObject.Visible = True
   cmdObject.Caption = "Dynamic CommandButton"
End Sub


'Handle the events of the dynamically-added control'
Private Sub cmdObject_Click()
    Print "This is a dynamically added control"
End Sub
