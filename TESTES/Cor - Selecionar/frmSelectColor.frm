VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cor Selecionada "
      Height          =   855
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecione uma cor"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgCores 
      Left            =   960
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    selectColor
End Sub

Private Sub selectColor()
    With dlgCores
        .ShowColor
        Label1.BackColor = dlgCores.Color
        Label1.Caption = dlgCores.Color
    End With
End Sub

