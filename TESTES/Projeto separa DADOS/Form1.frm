VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Separa Dados"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   2610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2415
      Begin VB.Label Label5 
         Caption         =   "Sequência - 5"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Sequencia - 4"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Sequencia - 3"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Sequencia - 2"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Sequencia - 1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dividir em sequência"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim RECEBE As String
    Dim CONTADOR As Integer
    MsgBox "Você digitou:" & Len(Text1) & " caracteres"
    CONTADOR = 0
    For x = 1 To Len(Text1)
        If Mid(Text1, x, 1) = ";" Then
            If CONTADOR = 0 Then Label1 = RECEBE
            If CONTADOR = 1 Then Label2 = RECEBE
            If CONTADOR = 2 Then Label3 = RECEBE
            If CONTADOR = 3 Then Label4 = RECEBE
            If CONTADOR = 4 Then Label5 = RECEBE
            CONTADOR = CONTADOR + 1
            RECEBE = ""
        Else
            MsgBox Mid(Text1, x, 1)
            RECEBE = RECEBE & Mid(Text1, x, 1)
        End If
    Next
    If CONTADOR = 0 Then Label1 = RECEBE
    If CONTADOR = 1 Then Label2 = RECEBE
    If CONTADOR = 2 Then Label3 = RECEBE
    If CONTADOR = 3 Then Label4 = RECEBE
    If CONTADOR = 4 Then Label5 = RECEBE
End Sub
