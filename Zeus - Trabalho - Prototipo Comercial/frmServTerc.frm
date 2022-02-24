VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmServTerc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Serviços Terceirizados"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmServTerc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmServTerc.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
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
      TabIndex        =   10
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   2
         Left            =   6240
         TabIndex        =   2
         Tag             =   "Unidade de medida"
         ToolTipText     =   "Unidade de medida"
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   6240
         OleObjectBlob   =   "frmServTerc.frx":1994
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         Picture         =   "frmServTerc.frx":19FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "Excluir Serviço"
         ToolTipText     =   "Excluir Serviço"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         Picture         =   "frmServTerc.frx":26C4
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "Editar Nome do Serviço"
         ToolTipText     =   "Editar Nome do Serviço"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmServTerc.frx":338E
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "Novo Serviço"
         ToolTipText     =   "Novo Serviço"
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmServTerc.frx":4058
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "Incluir Serviço"
         ToolTipText     =   "Incluir Serviço"
         Top             =   840
         Width           =   615
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
         Tag             =   "Código do Serviço"
         ToolTipText     =   "Código do Serviço"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadastro 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do Serviço"
         ToolTipText     =   "Nome do Serviço"
         Top             =   480
         Width           =   4815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmServTerc.frx":4D22
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmServTerc.frx":4D8A
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
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
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmServTerc.frx":4DF6
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Salvar Serviço"
      ToolTipText     =   "Salvar Serviço"
      Top             =   4920
      Width           =   615
   End
End
Attribute VB_Name = "frmServTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTipoTrei As New ADODB.Recordset
Private sqlTipoTrei As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If ValidaCampos(ListView2, txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)) = False Then Exit Sub
        
        IncluirLV ListView2, txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
        LimpaControles txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
        txtcadastro(0) = Format(GeraCodigoLV(ListView2), "00")
    Case 1
        LimpaControles txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
        txtcadastro(0) = Format(GeraCodigoLV(ListView2), "00")
    Case 2
        AlteraLV ListView2, txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
    Case 3
        ExcluirItemLV ListView2
    Case 4
        'Grava dados ListView2
        limpaQualquerDado
        ordenaLVArray ListView2, "0", "1", "2", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbServTerc", "idservTerc", "I", txtcadastro(0)
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
    listview_cabecalho
    'Abaixo Compoe Listview =========================
    chamaSQL "Select a.idservterc,a.nmserv,a.unidade from tbServTerc as a Order by a.idservterc"
    Compoe_Listview ListView2, Sqlp, "00"
    '================================================
    LimpaControles txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
    txtcadastro(0) = Format(GeraCodigoLV(ListView2), "00")
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Código", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 3
    ListView2.ColumnHeaders.Add , , "Un.", ListView2.Width / 12
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ListView2_DblClick()
    If vEdi <> "N" Then
        AlteraLV ListView2, txtcadastro(0), txtcadastro(1), txtcadastro(2), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0), txtcadastro(0)
    End If
End Sub

