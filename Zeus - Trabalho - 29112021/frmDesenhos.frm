VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmDesenhos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desenhos"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesenhos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   39
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   2640
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.CommandButton cmdDesenho 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   720
      Picture         =   "frmDesenhos.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdDesenho 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmDesenhos.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "Salvar"
      ToolTipText     =   "Salvar"
      Top             =   5640
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Desenho"
      TabPicture(0)   =   "frmDesenhos.frx":265E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Histórico de Revisões"
      TabPicture(1)   =   "frmDesenhos.frx":267A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Revisões "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   7575
         Begin VB.CommandButton cmdDesenho 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   1920
            Picture         =   "frmDesenhos.frx":2696
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "Excluir"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdDesenho 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   1320
            Picture         =   "frmDesenhos.frx":3360
            Style           =   1  'Graphical
            TabIndex        =   37
            Tag             =   "Editar"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdDesenho 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   720
            Picture         =   "frmDesenhos.frx":402A
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "Novo"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdDesenho 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   120
            Picture         =   "frmDesenhos.frx":4CF4
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "Inserir"
            Top             =   960
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   31
            Top             =   1680
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3625
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.TextBox txtDesenho 
            BackColor       =   &H80000018&
            Height          =   1095
            Index           =   8
            Left            =   2640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   480
            Width           =   4815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "frmDesenhos.frx":59BE
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   345
            Left            =   960
            TabIndex        =   28
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
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
            Format          =   166985729
            CurrentDate     =   41463
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmDesenhos.frx":5A28
            TabIndex        =   27
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtDesenho 
            Height          =   345
            Index           =   7
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmDesenhos.frx":5A8A
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblStatusRev 
            BackColor       =   &H8000000C&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dados "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   7575
         Begin VB.TextBox txtDesenho 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Tag             =   "Nº da FCE"
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdtDesenho 
            Caption         =   "..."
            Height          =   345
            Index           =   0
            Left            =   1440
            TabIndex        =   21
            Top             =   480
            Width           =   375
         End
         Begin VB.CommandButton cmdtDesenho 
            Caption         =   "..."
            Height          =   345
            Index           =   1
            Left            =   7080
            TabIndex        =   20
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtDesenho 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Tag             =   "Desenho"
            Top             =   1200
            Width           =   4695
         End
         Begin VB.TextBox txtDesenho 
            Height          =   1935
            Index           =   6
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   13
            Tag             =   "Descrição do desenho"
            Top             =   1920
            Width           =   7335
         End
         Begin VB.TextBox txtDesenho 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   3
            Left            =   2040
            TabIndex        =   11
            Tag             =   "Nº do projeto"
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtDesenho 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   5
            Left            =   4920
            TabIndex        =   18
            Tag             =   "Revisão do desenho"
            Top             =   1200
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "frmDesenhos.frx":5AE8
            TabIndex        =   10
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "frmDesenhos.frx":5B50
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmDesenhos.frx":5BB8
            TabIndex        =   14
            Top             =   1680
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   5760
            TabIndex        =   15
            Tag             =   "Data de cadastro do desenho"
            Top             =   1200
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
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
            Format          =   168427521
            CurrentDate     =   41407
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   5760
            OleObjectBlob   =   "frmDesenhos.frx":5C24
            TabIndex        =   16
            Top             =   960
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmDesenhos.frx":5C98
            TabIndex        =   19
            Top             =   960
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmDesenhos.frx":5D00
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo "
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
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboDesenho 
         Height          =   345
         ItemData        =   "frmDesenhos.frx":5D60
         Left            =   120
         List            =   "frmDesenhos.frx":5D6D
         TabIndex        =   2
         Tag             =   "Tipo de desenho"
         Text            =   "Fabricação"
         ToolTipText     =   "Tipo de desenho"
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Código do Projeto "
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
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtDesenho 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Código da FCE/Projetos"
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificador "
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
      TabIndex        =   5
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtDesenho 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador do desenho"
         ToolTipText     =   "Identificador do desenho"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmDesenhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsDesenhos As New ADODB.Recordset
Private sqlDesenhos As String
Private rsFCE As New ADODB.Recordset
Private sqlFCE As String
Private rsProjeto As New ADODB.Recordset
Private SqlProjeto As String
Private rsRevisao As New ADODB.Recordset
Private SqlRevisao As String
Private rsLocal As New ADODB.Recordset

Private Status As String

Private Sub cmdDesenho_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravarDados
            LimpaControles
            txtDesenho(4).SetFocus
'            gravaLog "Código esc.: " & txtDesenhos(0), "Nome esc: " & txtDesenhos(1), "Peso: " & txtDesenhos(2)
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela de cadastro?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Unload Me
            Set frmCD = Nothing
        End If
    Case 2
        IncluirRevisao
        LimpaControlesRevisao
    Case 3
        LimpaControlesRevisao
    Case 4
        AlteraRevisao
    Case 5
        If Msgbox("Deseja EXCLUIR essa revisão do Desenho?", vbQuestion + vbYesNo, "ZEUS") = vbYes Then
            ExcluirItemLV ListView1
            LimpaControlesRevisao
        End If
    End Select
End Sub

Private Sub cmdtDesenho_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridFCE
        CarregaFCE
    Case 1
        If txtDesenho(2) <> "" Then
            ChamaGridProjeto
            CarregaProjeto
        Else
            mobjMsg.Abrir "Informe o nº da FCE", Ok, critico, "Atenção"
            txtDesenho(2).SetFocus
        End If
    End Select
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    inicializa_tabs SSTab1, Picture1
    Status = Pesquisa
    listview_cabecalho
    DTPicker1 = Date
    SSTab1.Tab = 0
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
    End If
    carregarIconBotao
    MudaTool
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdDesenho(2), 2, 46 'Inserir
    carregaImagemBotao cmdDesenho(3), 3, 31 'Novo
    carregaImagemBotao cmdDesenho(4), 4, 32 'Editar
    carregaImagemBotao cmdDesenho(5), 5, 33 'Excluir
    carregaImagemBotao cmdDesenho(0), 0, 45 'Salvar
    carregaImagemBotao cmdDesenho(1), 1, 34 'Sair
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Revisão", ListView1.Width / 9
    ListView1.ColumnHeaders.Add , , "Data", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Detalhes", ListView1.Width / 1.5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub LimpaControlesRevisao()
    Dim x As Integer
    For x = 7 To 8
        txtDesenho(x) = ""
    Next
    DTPicker2 = Date
End Sub

Private Sub IncluirRevisao()
    Dim ItemLst As ListItem
    Dim x As Integer, y As Integer
    'If ValidaCampo = False Then Exit Sub
    y = ListView1.ListItems.Count
    If y > 0 Then
        For x = 1 To y
            If ListView1.ListItems.Item(x) = Me.txtDesenho(7) Then
                ListView1.ListItems.Item(x).Selected = True
                Me.txtDesenho(7) = ListView1.ListItems.Item(x)
                ListView1.SelectedItem.ListSubItems.Item(1) = DTPicker2
                ListView1.SelectedItem.ListSubItems.Item(2) = txtDesenho(8)
                y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtDesenho(7))
        y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtDesenho(7))
        y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = DTPicker2
    ItemLst.SubItems(2) = txtDesenho(8)
    txtDesenho(7).Text = ""
    DTPicker1 = Date
    txtDesenho(8).Text = ""
    txtDesenho(7).SetFocus
    lblStatusRev = "REVISADO"
End Sub

Private Sub AlteraRevisao()
    Dim y As Integer, x As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        If ListView1.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    Me.txtDesenho(7).Text = ListView1.ListItems.Item(x)
    Me.txtDesenho(8).Text = ListView1.SelectedItem.ListSubItems.Item(2)
    DTPicker2 = ListView1.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub GravarDados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    Dim rsDesenhos As New ADODB.Recordset
    Dim sqlDesenhos As String
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim y As Integer
10  cnBanco.BeginTrans
   
    sqlDesenhos = "select * from tbDesenhos as a where a.iddesenho = '" & txtDesenho(0) & "'"
    rsDesenhos.Open sqlDesenhos, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsDesenhos.EOF Then rsDesenhos.AddNew
    rsDesenhos.Fields(0) = Val(txtDesenho(0)) 'Identificador do desenho
    rsDesenhos.Fields(1) = DTPicker1.Value ' Data de cadastro do desenho
    rsDesenhos.Fields(2) = Val(txtDesenho(1).Text) 'Código de identificação FCE/Projeto
    rsDesenhos.Fields(3) = txtDesenho(4).Text 'Desenho
    rsDesenhos.Fields(4) = txtDesenho(5).Text 'Nº da revisão do desenho
    rsDesenhos.Fields(5) = txtDesenho(6).Text 'Descrição do desenho
    rsDesenhos.Fields(6) = cboDesenho.Text 'Tipo do desenho
    
    If Check1.Value = 0 Then
        rsDesenhos.Fields(7) = "N" 'Ativo
    Else
        rsDesenhos.Fields(7) = "S" 'Ativo
    End If
    rsDesenhos.Fields(8) = vCodcoligada 'Código da coligada
    
    rsDesenhos.Update
    cnBanco.CommitTrans
    rsDesenhos.Close
    Set rsDesenhos = Nothing
    
    '>>>> GRAVA REVISAO DE DESENHO
    sqlDeletar = "Delete from tbdesenhosrev where tbdesenhosrev.codcoligada = '" & vCodcoligada & "' and tbdesenhosrev.iddesenho = '" & Val(txtDesenho(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbdesenhosrev where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For x = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(x).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtDesenho(0).Text)
        rsSalvar.Fields(1) = ListView1.ListItems.Item(x)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(4) = vCodcoligada 'Codigo da Coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    AtualizaListview
    mobjMsg.Abrir "Os dados foram salvos com sucesso", Ok, informacao, "ZEUS"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub LimpaControles()
    Dim x As Integer
    For x = 4 To txtDesenho.Count - 1
        txtDesenho(x) = ""
    Next
    cboDesenho = "Fabricação"
    txtDesenho(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub CompoeControles()
    Dim x As Integer
    txtDesenho(0).Text = Format(rsDesenhos.Fields(0), "000000") 'IDDesenho
    txtDesenho(1).Text = Format(rsDesenhos.Fields(1), "000000") 'Código do Projeto
    cboDesenho.Text = rsDesenhos.Fields(2) 'Tipo
    txtDesenho(2).Text = rsDesenhos.Fields(3) 'FCE
    txtDesenho(3).Text = rsDesenhos.Fields(4) 'Projeto
    txtDesenho(4).Text = rsDesenhos.Fields(5) 'Desenho
    txtDesenho(5).Text = rsDesenhos.Fields(6) 'Revisão
    DTPicker1.Value = rsDesenhos.Fields(7) 'Data cadastro
    txtDesenho(6).Text = rsDesenhos.Fields(8) 'Descrição
    If rsDesenhos.Fields(9) = "S" Then 'Ativo
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    For x = 0 To 5
        If txtDesenho(x).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtDesenho(x).Tag, Ok, critico, "Atenção"
            Me.txtDesenho(x).SetFocus
            Exit Function
        End If
    Next
    If cboDesenho.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.cboDesenho.Tag, Ok, critico, "Atenção"
        Me.cboDesenho.SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirDesenhos
    SqlGera = "Select top 1 * from tbDesenhos order by iddesenho Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDesenhos.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtDesenho(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharDesenhos
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function

Private Sub AbrirDesenhos()
On Error GoTo Err
    sqlDesenhos = "Select * from tbDesenhos Order by iddesenho"
    rsDesenhos.Open sqlDesenhos, cnBanco, adOpenKeyset, adLockOptimistic
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub AbrirRevisao()
On Error GoTo Err
    SqlRevisao = "Select * from tbdesenhosrev where codcoligada = '" & vCodcoligada & "' and iddesenho = '" & Val(txtDesenho(0)) & "'Order by iddesenho,revisao"
    rsRevisao.Open SqlRevisao, cnBanco, adOpenKeyset, adLockOptimistic
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub FecharRevisao()
    rsRevisao.Close
    Set rsRevisao = Nothing
End Sub

Private Sub FecharDesenhos()
    rsDesenhos.Close
    Set rsDesenhos = Nothing
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    sqlDesenhos = "Select a.iddesenho,a.codprojeto,a.tipo,b.fce,b.projeto,a.desenho,a.revisao,a.datacadastro,a.descricao,a.ativo from tbDesenhos as a left join tbProjetos as b on a.codprojeto = b.codprojeto Where a.iddesenho= '" & Val(varGlobal) & "' and ativo = 'S' order by a.iddesenho"
    rsDesenhos.Open sqlDesenhos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDesenhos.RecordCount > 0 Then
        CompoeControles
        AbrirRevisao
        Compoe_Listview
        FecharRevisao
    Else
        mobjMsg.Abrir "Identificador não encontrado", Ok, critico, "Atenção"
    End If
    rsDesenhos.Close
    Set rsDesenhos = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub Compoe_Listview()
    'PREENCHE O LISTVIEW DE REVISAO
    Dim ItemLst As ListItem
    Dim x As Integer
    x = 0
    While Not rsRevisao.EOF
        Set ItemLst = ListView1.ListItems.Add(, , rsRevisao.Fields(1))
        ItemLst.SubItems(1) = "" & rsRevisao.Fields(2)
        ItemLst.SubItems(2) = "" & rsRevisao.Fields(3)
        rsRevisao.MoveNext
        x = x + 1
    Wend
End Sub

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    y = vListViewPrincipal.ListItems.Count
    For x = 1 To y
        If vListViewPrincipal.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = vListViewPrincipal.ListItems.Add(, , Format(txtDesenho(0), "000000")) 'Identificador
        ItemLst.SubItems(1) = txtDesenho(4).Text 'Desenho
        ItemLst.SubItems(2) = txtDesenho(5).Text 'Revisão
        ItemLst.SubItems(3) = txtDesenho(2).Text 'FCE
        ItemLst.SubItems(4) = txtDesenho(3).Text 'Projeto
        ItemLst.SubItems(5) = DTPicker1.Value 'Data cadastro
        ItemLst.SubItems(6) = cboDesenho.Text 'Tipo de desenho
        If Check1.Value = 0 Then 'Ativo
            ItemLst.SubItems(7) = ""
            ItemLst.ListSubItems.Item(7).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(7) = ""
            ItemLst.ListSubItems.Item(7).ReportIcon = "OK"
        End If
    Else
        vListViewPrincipal.SelectedItem.ListSubItems.Item(1) = txtDesenho(4).Text 'Desenho
        vListViewPrincipal.SelectedItem.ListSubItems.Item(2) = txtDesenho(5).Text 'Revisão
        vListViewPrincipal.SelectedItem.ListSubItems.Item(3) = txtDesenho(2).Text 'FCE
        vListViewPrincipal.SelectedItem.ListSubItems.Item(4) = txtDesenho(3).Text 'Projeto
        vListViewPrincipal.SelectedItem.ListSubItems.Item(5) = DTPicker1.Value 'Data cadastro
        vListViewPrincipal.SelectedItem.ListSubItems.Item(6) = cboDesenho.Text 'Tipo de desenho
        If Check1.Value = 0 Then 'Ativo
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7).ReportIcon = "EXC"
        Else
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7) = ""
            vListViewPrincipal.SelectedItem.ListSubItems.Item(7).ReportIcon = "OK"
        End If
        
        'If cboDesenhos(2).Text <> "" Then vListViewPrincipal.SelectedItem.ListSubItems.Item(16) = cboDesenhos(2).Text 'Detalhista
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
        Exit Sub
    End If
End Sub

Private Sub txtDesenho_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo Error
    Select Case Index
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaFCE
        End If
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If txtDesenho(2) <> "" Then
                CarregaProjeto
            Else
                mobjMsg.Abrir "FCE não informada", Ok, critico, "Atenção"
                txtDesenho(3) = ""
            End If
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaFCE()
On Error GoTo Err
    Dim x As Integer
    sqlFCE = "Select a.* from tbprojetos as a inner join tbFCE as b on a.fce = b.fce where a.fce = '" & txtDesenho(2) & "' and b.status = 0 order by a.fce"
    'sqlFCE = "Select * from tbprojetos where fce = '" & txtDesenho(2) & "' order by fce"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFCE.EOF Then
        txtDesenho(2).Text = txtDesenho(2)
        mobjMsg.Abrir "FCE não cadastrada", Ok, critico, "Atenção"
    Else
        txtDesenho(2).Text = rsFCE.Fields(1)
    End If
    rsFCE.Close
    Set rsFCE = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Resume Next
    End If
End Sub

Private Sub ChamaGridFCE()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
'    Sqlp = "Select fce,MAX(oc) from tbprojetos group by FCE order by fce"
    Sqlp = "Select a.fce,MAX(a.oc) from tbprojetos as a inner join tbFCE as b on a.fce=b.fce where b.status = 0 group by a.FCE,b.status order by a.fce"
    procnom = "fce"
    campo = 0
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de FCE"
    Pesquisa = frmDesenhos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "fce=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtDesenho(2).Text = rsLocal.Fields(0)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub CarregaProjeto()
On Error GoTo Err
    Dim x As Integer
    SqlProjeto = "Select * from tbprojetos where fce = '" & txtDesenho(2) & "' order by fce"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsProjeto.EOF Then rsProjeto.MoveFirst
    rsProjeto.Find "projeto=" & "'" & Me.txtDesenho(3) & "'"
    If rsProjeto.EOF Then
        txtDesenho(3).Text = txtDesenho(3)
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Projeto não cadastrado", Ok, critico, "Atenção"
        End If
    Else
        txtDesenho(3).Text = rsProjeto.Fields(2)
        txtDesenho(1).Text = Format(rsProjeto.Fields(0), "000000")
    End If
    rsProjeto.Close
    Set rsProjeto = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

Private Sub ChamaGridProjeto()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbprojetos where fce = '" & txtDesenho(2) & "' order by fce,Projeto"
    procnom = "projeto"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Projetos"
    Pesquisa = frmDesenhos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "projeto=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtDesenho(3).Text = rsLocal.Fields(2)
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Sub

