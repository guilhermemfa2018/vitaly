VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmCCusto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Centro de Custo"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmCCusto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmCCusto.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   4920
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         Picture         =   "frmCCusto.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Excluir Grupo"
         ToolTipText     =   "Excluir Grupo"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         Picture         =   "frmCCusto.frx":265E
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Editar Nome do Grupo"
         ToolTipText     =   "Editar Nome do Grupo"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmCCusto.frx":3328
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Novo Grupo"
         ToolTipText     =   "Novo Grupo"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Tag             =   "Nome do Centro de Custo"
         ToolTipText     =   "Nome do Centro de Custo"
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox txtCadastro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do Centro de Custo"
         ToolTipText     =   "Código do Centro de Custo"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmCCusto.frx":3FF2
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Incluir Grupo"
         ToolTipText     =   "Incluir Grupo"
         Top             =   840
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5318
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmCCusto.frx":4CBC
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   4920
      Width           =   615
   End
End
Attribute VB_Name = "frmCCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipoTrei As New ADODB.Recordset
Private sqlTipoTrei As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
        LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
        'txtformula(2) = Format(GeraCodigoLV(ListView1), "000")
        'IncluirTipo
        'LimpaControles
    Case 1
        LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    Case 2
        AlteraLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    Case 3
        ExcluirItemLV ListView1
    Case 4
        'Grava dados ListView1
        limpaQualquerDado
        ordenaLVArray ListView1, "0", "1", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbCCusto", "", "S", txtCadastro(0)
        MsgBox "Dados Salvos com sucesso!", vbInformation, "PrototipoX"
        Unload Me
    Case 5
        Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    listview_cabecalho
    'Abaixo Compoe Listview =========================
    chamaSQL "Select a.idprd,a.nome from tbCCusto as a Order by a.idprd"
    Compoe_Listview ListView1, Sqlp, ""
    '================================================
    'Compoe_Listview
    LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    Exit Sub
ErrHandler:
    MsgBox "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", vbCritical, "Atenção"
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 1.5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    End If
End Sub
