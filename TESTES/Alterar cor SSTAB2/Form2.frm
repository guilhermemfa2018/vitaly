VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7200
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Show Form1"
      Height          =   555
      Left            =   915
      TabIndex        =   0
      Top             =   690
      Width           =   1905
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show 1
End Sub
