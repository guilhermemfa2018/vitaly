VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Protótipo M&P"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14580
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selecione uma das opções abaixo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Option5 
         Caption         =   "Design Final"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Metodos e Processos"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Cadastrar Centro de Custo"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2655
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Testar fórmula"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cadastrar fórmula"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Configuração de conexão DB RM Sistemas "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5280
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Text            =   "server\sql2008"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Text            =   "PROTOTIPO"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   1200
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   720
      Picture         =   "Form1.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Option1.Value = True Then
        frmFormulaCC.Show 1
    ElseIf Option2.Value = True Then
        frmMetodo.Show 1
    ElseIf Option3.Value = True Then
        frmCCusto.Show 1
    ElseIf Option4.Value = True Then
        frmMP.Show 1
    ElseIf Option5.Value = True Then
        frmMPCompleto.Show 1
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
'    Conectar
End Sub

