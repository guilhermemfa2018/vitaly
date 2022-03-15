VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Grupos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmGrupo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmGrupo.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmGrupo.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do Grupo"
         ToolTipText     =   "Nome do Grupo"
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox txtCadastro 
         Enabled         =   0   'False
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
         Tag             =   "Código do Grupo"
         ToolTipText     =   "Código do Grupo"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmGrupo.frx":265E
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmGrupo.frx":26C6
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Centro de Custo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   10
         Top             =   840
         Width           =   2535
         Begin VB.TextBox txtCadastro 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Text            =   "Centro de Custo"
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         Picture         =   "frmGrupo.frx":2732
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
         Picture         =   "frmGrupo.frx":33FC
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
         Picture         =   "frmGrupo.frx":40C6
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Novo Grupo"
         ToolTipText     =   "Novo Grupo"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmGrupo.frx":4D90
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
   End
End
Attribute VB_Name = "frmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipoTrei As New ADODB.Recordset
Private sqlTipoTrei As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(2), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
        LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
        txtCadastro(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 1
        LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
        txtCadastro(0) = Format(GeraCodigoLV(ListView1), "00")
    Case 2
        AlteraLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(2), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    Case 3
        ExcluirItemLV ListView1
    Case 4
        'Grava dados ListView1
        limpaQualquerDado
        ordenaLVArray ListView1, "2", "0", "1", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbGrupoClass", "idprd", "I", txtCadastro(2)
        Msgbox "Dados Salvos com sucesso!", vbInformation, "PrototipoX"
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
'On Error GoTo ErrHandler
    txtCadastro(2).Text = frmFormulaCC.txtformula(0).Text
    listview_cabecalho
    'Abaixo Compoe Listview =========================
    chamaSQL "Select a.idgrupo,a.nmgrupo,a.idprd from tbGrupoClass as a where idprd = '" & frmFormulaCC.txtformula(0) & "' Order by a.idgrupo"
    Compoe_Listview ListView1, Sqlp, "00"
    '================================================
    LimpaControles txtCadastro(0), txtCadastro(1), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    txtCadastro(0) = Format(GeraCodigoLV(ListView1), "00")
    carregarIconBotao
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadastro(0), 0, 46 'Inserir
    carregaImagemBotao cmdCadastro(1), 1, 31 'Novo
    carregaImagemBotao cmdCadastro(2), 2, 32 'Editar
    carregaImagemBotao cmdCadastro(3), 3, 33 'Excluir
    carregaImagemBotao cmdCadastro(4), 4, 45 'Salvar
    carregaImagemBotao cmdCadastro(5), 5, 34 'Sair
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "C.Custo", ListView1.Width / 10000
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraLV ListView1, txtCadastro(0), txtCadastro(1), txtCadastro(2), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0), txtCadastro(0)
    End If
End Sub

