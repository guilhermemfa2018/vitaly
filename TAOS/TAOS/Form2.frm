VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13860
   LinkTopic       =   "Form2"
   ScaleHeight     =   9075
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Form1.Show 1
End Sub

Private Sub Form_Load()
    AlwaysOnTop Me, True ' Mantem o formulário sempre em primeiro plano
End Sub
