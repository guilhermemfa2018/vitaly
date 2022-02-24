VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form MostraClientes 
   Caption         =   "CLIENTES"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   Icon            =   "MostraClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   14490
   Begin VB.Frame Frame1 
      Caption         =   "PESQUISA CLIENTES"
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
      TabIndex        =   11
      Top             =   0
      Width           =   10575
      Begin VB.OptionButton Inativos 
         Caption         =   "Inativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Ativos 
         Caption         =   "Ativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
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
         ItemData        =   "MostraClientes.frx":0E42
         Left            =   3840
         List            =   "MostraClientes.frx":0E73
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
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
         Left            =   6240
         TabIndex        =   0
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox Tipo 
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
         ItemData        =   "MostraClientes.frx":0EE8
         Left            =   1560
         List            =   "MostraClientes.frx":0EF5
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label1 
         Height          =   255
         Left            =   6240
         OleObjectBlob   =   "MostraClientes.frx":0F20
         TabIndex        =   12
         Top             =   240
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "MostraClientes.frx":0FB2
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "MostraClientes.frx":102C
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5415
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1815
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2640
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2040
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1440
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "MostraClientes.frx":10A0
      Height          =   4935
      Left            =   1920
      OleObjectBlob   =   "MostraClientes.frx":10B4
      TabIndex        =   1
      Top             =   1080
      Width           =   12435
   End
End
Attribute VB_Name = "MostraClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlterar_Click()
    AlteraClientes.Show
End Sub

Private Sub cmdCadastrar_Click()
    CadClientes.Show
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_GotFocus()
Call BordasControle(Me, DBGrid1, False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then 'F2
    cmdCadastrar_Click
ElseIf KeyCode = 118 Then 'F7
    cmdAlterar_Click
ElseIf KeyCode = 119 Then 'F8
    'cmdExcluir_Click
ElseIf KeyCode = 121 Then 'F10
    'cmdImprimir_Click
End If
End Sub

Private Sub Form_Resize()
OrganizaControles
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AlteraClientes.Show
End If
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
AplicarSkin Me, Principal.Skin1
NewColorDBGrid Me

On Error GoTo ErrHandler

Call BordasControle(Me, DBGrid1, False)

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
DBGrid1.Move 1920, 1080, Me.ScaleWidth - 2000, Me.ScaleHeight - 1080
End Function
