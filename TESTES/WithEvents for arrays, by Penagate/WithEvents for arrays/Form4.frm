VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13065
   LinkTopic       =   "Form4"
   ScaleHeight     =   7110
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1575
      Left            =   6240
      TabIndex        =   6
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teste de criação dinamica de componentes "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   12615
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "2º: ADD MAIS 6 COMMAND2 BUTTONS (UM COMMAND2 BUTTON EXISTE NA TELA COM INDICE = 0)"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   7935
      End
      Begin VB.Label Label1 
         Caption         =   "1º: CRIA 6 COMMANDBUTOM DINAMICAMENTE"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   7200
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Exibe 1 a 1 o name de cada componente que está no Form4"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5760
      Top             =   3240
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmd1(5) As VB.CommandButton

Private Sub Form_Load()
    Dim X As Integer
    '1º: CRIA 6 COMMANDBUTOM DINAMICAMENTE
    'FIRST, CREATE 6 COMMAND BUTTONS 'FROM SCRATCH' (NO CTLCOMMAND BUTTONS ON THE FORM)
    For X = 0 To 5
        Set cmd1(X) = Controls.Add("VB.CommandButton", "cmd1" & CStr(X), Frame1)
        cmd1(X).Visible = True
        cmd1(X).Caption = "cmd1" & CStr(X)
        cmd1(X).Top = 600
        cmd1(X).Left = 240
    Next X
    For X = 0 To 4
        cmd1(X + 1).Left = cmd1(X).Left + cmd1(X).Width + 50
    Next X
    
    '2º: ADD MAIS 6 COMMAND2 BUTTONS (UM COMMAND2 BUTTON EXISTE NA TELA COM INDICE = 0)
    'NEXT, ADD 3 MORE COMMAND2 BUTTONS (A COMMAND2 BUTTON EXISTS ON THE SCREEN WITH INDEX = 0)
    For X = 1 To 6
       Load Command2(X)
       Command2(X).Visible = True
    Next X
    For X = 0 To 5
       Command2(X + 1).Left = Command2(X).Left + Command2(X).Width + 40
       Command2(X + 1).Top = Command2(X).Top
    Next X
End Sub

Private Sub Command2_Click(index As Integer) 'WHEN MANUALLY CLICKING ANY COMMADN2 BUTTON, WORKS FINE
    MsgBox index
End Sub

'EXECUTA OS EVENTOS DO BOTÃO CMD1(X) - CRIADOS DINAMICAMENTE
'QUANDO CLICAR MANUALMENTE EM QUALQUER BOTÃO CMD1, NENHUM EVENTO É ACIONADO, MAS QUANDO CHAMADO, FUNCIONA BEM
Private Sub Command3_Click() 'WHEN MANUALLY CLICKING ANY CMD1 BUTTONS, NO EVENT IS TRIGGERED, BUT WHEN CALLED, WORKS FINE
    Dim X As Integer
    For X = 0 To 5
        Call cmd1_click(X)
    Next
End Sub

'EXECUTA O CLICK DOS BOTÕES CMD1(X) - CRIADOS DINAMICAMENTE
'SUB CHAMADA DO CLICK DO BOTÃO COMMAND3
Private Sub cmd1_click(index As Integer) 'SUB CALLED FROM COMMAND3 CLICK EVENT ABOVE
    Select Case index
        Case 0
            MsgBox cmd1(index).Name
        Case 1
            MsgBox cmd1(index).Name
        Case 2
            MsgBox cmd1(index).Name
        Case 3
            MsgBox cmd1(index).Name
    End Select
End Sub

Private Sub Command1_Click()
    Dim cntrl As Control
     For Each cntrl In Me.Controls
          MsgBox cntrl.Name
     Next
End Sub
