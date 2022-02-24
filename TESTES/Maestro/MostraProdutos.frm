VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form MostraProdutos 
   Caption         =   "PRODUTOS"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   Icon            =   "MostraProdutos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   15045
   Begin VB.CommandButton cmdEntrada 
      Caption         =   "En&trada em Estoque (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   15
      ToolTipText     =   "Entrada de Produtos em Estoque"
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "S&aída do Estoque(F6)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      TabIndex        =   14
      ToolTipText     =   "Saíd de Produtos do Estoque"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Height          =   6015
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1815
      Begin VB.CommandButton cmdCadastrar 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Caption         =   "&Novo (F2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdAlterar 
         BackColor       =   &H8000000B&
         Caption         =   "&Editar (F7)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdExcluir 
         BackColor       =   &H8000000B&
         Caption         =   "E&xcluir (F8)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H8000000B&
         Caption         =   "&Imprimir (F10)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H8000000B&
         Caption         =   "&Sair (Esc)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PESQUISA PRODUTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   1920
      TabIndex        =   8
      Top             =   0
      Width           =   12975
      Begin VB.TextBox Nome 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "Fone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "MostraProdutos.frx":030A
         Left            =   240
         List            =   "MostraProdutos.frx":0320
         TabIndex        =   0
         Text            =   "Descrição"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "Fone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "MostraProdutos.frx":0372
         Left            =   2880
         List            =   "MostraProdutos.frx":0388
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "MostraProdutos.frx":03A1
         TabIndex        =   12
         Top             =   240
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "MostraProdutos.frx":0433
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "MostraProdutos.frx":04AD
      Height          =   4875
      Left            =   1920
      Negotiate       =   -1  'True
      OleObjectBlob   =   "MostraProdutos.frx":04C1
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1080
      Width           =   12975
   End
End
Attribute VB_Name = "MostraProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCadastrar_Click()
    CadProd.Show
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid2_GotFocus()
Call BordasControle(Me, DBGrid2, False)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then 'F2
    cmdCadastrar_Click
ElseIf KeyCode = 118 Then 'F7
    cmdAlterar_Click
ElseIf KeyCode = 119 Then 'F8
    'cmdExcluir_Click
ElseIf KeyCode = 121 Then 'F10
    'cmdImprimir_Click
ElseIf KeyCode = 116 Then 'F5
    'cmdEntrada_Click
ElseIf KeyCode = 117 Then 'F6
    'cmdSaida_Click
End If
End Sub

Private Sub Form_Resize()
OrganizaControles
End Sub

Private Sub cmdAlterar_Click()
AlteraProd.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
AplicarSkin Me, Principal.Skin1
NewColorDBGrid Me
On Error GoTo ErrHandler

Call BordasControle(Me, DBGrid2, False)

OrganizaForm
OrganizaControles

Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Function OrganizaForm()
Me.Move 0, 0, Principal.ScaleWidth - 200, Principal.ScaleHeight - 350
End Function

Private Function OrganizaControles()
On Error Resume Next
Frame1.Move 1920, 0, Me.ScaleWidth - 2000
Frame3.Move 0, 0, Me.ScaleWidth - Me.ScaleWidth + 1800, Me.ScaleHeight
DBGrid2.Move 1920, 1080, Me.ScaleWidth - 2000, Me.ScaleHeight - 1080
End Function
