VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Relatório"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Coligada "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1080
      TabIndex        =   27
      Top             =   6120
      Width           =   3375
      Begin VB.CheckBox chkRel 
         Caption         =   "Luna"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chkRel 
         Caption         =   "Vitaly"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkRel 
         Caption         =   "Viga"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Movimento "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton OptMov 
         Caption         =   "Solicitação de Compra"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton OptMov 
         Caption         =   "Ordem de Compra"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Período "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93716481
         CurrentDate     =   40969
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93716481
         CurrentDate     =   40969
      End
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "Form1.frx":3469A
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "Gerar relatório"
      ToolTipText     =   "Gerar relatório"
      Top             =   6720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   4335
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "PRODUTO:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "FORNECEDOR:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "FCE:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de relatório "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.OptionButton OptRel 
         Caption         =   "PRODUTO"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton OptRel 
         Caption         =   "Produto por FORNECEDOR"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton OptRel 
         Caption         =   "Produto por FCE"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Configuração de conexão DB RM Sistemas "
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Text            =   "SRV1002\CORPORERM"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Text            =   "CORPORERM"
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "sa"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "vigamax"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Nome do SERVIDOR:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label17 
         Caption         =   "Nome do BANCO:"
         Height          =   255
         Left            =   3120
         TabIndex        =   19
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label18 
         Caption         =   "USUÁRIO:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label19 
         Caption         =   "SENHA:"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    vDataFilter1 = DTPicker1
    vDataFilter2 = DTPicker2
    
    If chkRel(0).Value = 1 And chkRel(1).Value = 0 And chkRel(2).Value = 0 Then
        vCodColigada = "1"
        vNomeColigada = "VIGA"
    ElseIf chkRel(0).Value = 0 And chkRel(1).Value = 1 And chkRel(2).Value = 0 Then
        vCodColigada = "5"
        vNomeColigada = "VITALY"
    ElseIf chkRel(0).Value = 0 And chkRel(1).Value = 0 And chkRel(2).Value = 1 Then
        vCodColigada = "6"
        vNomeColigada = "LUNA"
    ElseIf chkRel(0).Value = 1 And chkRel(1).Value = 1 And chkRel(2).Value = 0 Then
        vCodColigada = "'1','5'"
        vNomeColigada = "VIGA/VITALY"
    ElseIf chkRel(0).Value = 1 And chkRel(1).Value = 0 And chkRel(2).Value = 1 Then
        vCodColigada = "'1','6'"
        vNomeColigada = "VIGA/LUNA"
    ElseIf chkRel(0).Value = 0 And chkRel(1).Value = 1 And chkRel(2).Value = 1 Then
        vCodColigada = "'5','6'"
        vNomeColigada = "VITALY/LUNA"
    ElseIf chkRel(0).Value = 1 And chkRel(1).Value = 1 And chkRel(2).Value = 1 Then
        vCodColigada = "'1','5','6'"
        vNomeColigada = "VIGA/VITALY/LUNA"
    Else
        MsgBox "Deve ser selecionado ao menos uma coligada"
        Exit Sub
    End If
    
    
    If OptMov(0).Value = True Then
        vMovimento = "1.1.03"
    End If
    If OptMov(1).Value = True Then
        vMovimento = "1.1.02"
    End If
    If OptRel(0).Value = True Then
        vFCE = Text1.Text
        vProduto = Text3.Text
        Form4.Show 1
    End If
    If OptRel(1).Value = True Then
        vFornec = Text2.Text
        Form2.Show 1
    End If
    If OptRel(2).Value = True Then
        vProduto = Text3.Text
        Form3.Show 1
    End If
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date - 7
    DTPicker2.Value = Date
    Conectar
End Sub

Private Sub OptMov_Click(Index As Integer)
    If OptMov(0).Value = True Then
        OptRel(1).Enabled = True
        Label2.Enabled = True
        Text2.Enabled = True
    End If
    If OptMov(1).Value = True Then
        OptRel(1).Enabled = False
        Label2.Enabled = False
        Text2.Enabled = False
    End If
End Sub

Private Sub OptRel_Click(Index As Integer)
    Label1.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    If OptRel(0).Value = True Then
        Label2.Enabled = False
        Text2.Enabled = False
    End If
    If OptRel(1).Value = True Then
        Label1.Enabled = False
        Text1.Enabled = False
        Label3.Enabled = False
        Text3.Enabled = False
    End If
    If OptRel(2).Value = True Then
        Label1.Enabled = False
        Label2.Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
    End If
End Sub
