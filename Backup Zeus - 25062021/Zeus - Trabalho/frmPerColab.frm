VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPerColab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permissões"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPerColab.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Colaborador"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtPermissao 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtPermissao 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   5655
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmPerColab.frx":0A02
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmPerColab.frx":0A66
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Permitir fechamento de OS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "                                                            (Informe os CC)"
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
      Height          =   6135
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   7455
      Begin VB.CheckBox Check3 
         Caption         =   "Apropria OS"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Inserir código para encerrar sistema"
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   720
         Picture         =   "frmPerColab.frx":0AC8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmPerColab.frx":1792
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtPermissao 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "ID Centro de Custo"
         ToolTipText     =   "ID Centro de Custo"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPermissao 
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
         Height          =   330
         Index           =   3
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   4935
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4215
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7435
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   2400
         OleObjectBlob   =   "frmPerColab.frx":245C
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmPerColab.frx":24C8
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin ZEUS.chameleonButton cmdPermissao 
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   7440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPerColab.frx":2530
      PICN            =   "frmPerColab.frx":254C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdPermissao 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   7440
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPerColab.frx":3226
      PICN            =   "frmPerColab.frx":3242
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmPerColab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Frame2.Enabled = True
        SkinLabel6.Enabled = True
        SkinLabel7.Enabled = True
        txtPermissao(2).Enabled = True
        cmdCadastro(0).Enabled = True
        cmdCadastro(1).Enabled = True
        cmdCadastro(3).Enabled = True
        ListView1.Enabled = True
    Else
        Frame2.Enabled = False
        SkinLabel6.Enabled = False
        SkinLabel7.Enabled = False
        txtPermissao(2).Enabled = False
        cmdCadastro(0).Enabled = False
        cmdCadastro(1).Enabled = False
        cmdCadastro(3).Enabled = False
        ListView1.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        insereIDFechamento
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        insereApropOS
    End If
End Sub

Private Sub insereIDFechamento()
    Dim rsIDFechamento As New ADODB.Recordset
    Dim sqlIDFechamento As String
    sqlIDFechamento = "Select a.codigo,a.nmparada from tbparadas as a where a.tipo = 'Fechamento'"
    rsIDFechamento.Open sqlIDFechamento, cnBanco, adOpenKeyset, adLockOptimistic
    If rsIDFechamento.RecordCount = 0 Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = rsIDFechamento.Fields(0) Then
                ListView1.ListItems.Item(X).Selected = True
                rsIDFechamento.Fields(0) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = rsIDFechamento.Fields(1)
                ListView1.SelectedItem.ListSubItems.Item(2) = txtPermissao(0).Text
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , rsIDFechamento.Fields(0))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , rsIDFechamento.Fields(0))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = rsIDFechamento.Fields(1)
    ItemLst.SubItems(2) = txtPermissao(0).Text
End Sub

Private Sub insereApropOS()
    Set ItemLst = ListView1.ListItems.Add(, , "1")
    Y = ListView1.ListItems.Count
    ItemLst.SubItems(1) = "APROPRIAR OS"
    ItemLst.SubItems(2) = txtPermissao(0).Text
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGrid "CORPORERM.dbo.GCCUSTO", "codreduzido", txtPermissao(2), frmPerColab, "codreduzido", "nome"
        CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtPermissao(2), txtPermissao(3), 7, 2, txtPermissao(2), "S", txtPermissao(3), "1"
    Case 1
        IncluirLV ListView1, txtPermissao(2), txtPermissao(3), txtPermissao(0), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2)
        LimpaControles txtPermissao(2), txtPermissao(3), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2), txtPermissao(2)
    Case 3
        ExcluirItemLV ListView1
    End Select
End Sub

Private Sub cmdPermissao_Click(Index As Integer)
    Select Case Index
    Case 0
        limpaQualquerDado
        'Grava dados do formulário
        'O 1º parametro é o valor que sera gravado no campo
        'O 2º parametro é o tipo de dado que o campo armazena
        vQualquerDado(1, 1) = txtPermissao(0).Text
        vQualquerDado(1, 2) = "S"
        GravaDados "tbAutFechaOS", "chapa", "S", txtPermissao(0), 1, "", "", txtPermissao(0)
    
        limpaQualquerDado
        ordenaLVArray ListView1, "2", "0", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbAutCCusto", "chapa", "S", txtPermissao(0)
        
        mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
    
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    txtPermissao(0) = varGlobal
    
    If Mid$(varGlobal, 1, 5) = "CONTR" Then
        chamaSQL "select a.CHAPA,a.NOME from tbTerceirizados as a where a.chapa = '" & varGlobal & "' and a.ativo = 'S' ORDER BY a.chapa"
    Else
        chamaSQL "select a.CHAPA,b.NOME from " & vBancoTotvs & ".dbo.PFUNC as a inner join " & vBancoTotvs & ".dbo.PPESSOA as b on a.CODSITUACAO in('A','F','P','Z') and a.CODPESSOA = b.CODIGO and cast(a.CHAPA as int)> 10  where a.chapa = '" & varGlobal & "' ORDER BY a.chapa"
    End If
    
    CarregaTxt "", "codreduzido", "S", "", "", txtPermissao(0), txtPermissao(1), 0, 1, txtPermissao(0), "S", txtPermissao(1), "2"
    Check1_Click
    listview_cabecalho
    'Abaixo Compoe Listview =========================
    chamaSQL "select b.idcc,c.NOME,a.chapa from tbAutFechaOS as a inner join tbAutCCusto as b on a.chapa = b.chapa inner join " & vBancoTotvs & ".dbo.GCCUSTO as c on b.idcc COLLATE SQL_Latin1_General_CP1_CI_AS = c.CODREDUZIDO and c.nome <> 'VIGA CALDEIRARIA LTDA' where a.chapa = '" & varGlobal & "' order by b.idcc"
    Compoe_Listview ListView1, Sqlp, ""
    
    chamaSQL "select b.idcc,c.nmparada,a.chapa from tbAutFechaOS as a inner join tbAutCCusto as b on a.chapa = b.chapa inner join tbParadas as c on b.idcc = c.codigo where a.chapa = '" & varGlobal & "' order by b.idcc"
    Compoe_Listview ListView1, Sqlp, ""
    
    '================================================
    If ListView1.ListItems.Count > 0 Then Check1.Value = 1 Else Check1.Value = 0
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID. C.Custo", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nome C. Custo", ListView1.Width / 1.6
    ListView1.ColumnHeaders.Add , , "Chapa", ListView1.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub txtPermissao_GotFocus(Index As Integer)
    mudaCorText txtPermissao(Index)
End Sub

Private Sub txtPermissao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If txtPermissao(2).Text = "" Then
                mobjMsg.Abrir "Selecione primeiro um CC - Centro de Custo", Ok, critico, "ZEUS"
                Exit Sub
            End If
            CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtPermissao(2), txtPermissao(3), 7, 2, txtPermissao(2), "S", txtPermissao(3), "1"
            'montaEstrutTreeview
            'LimpaVariaveis
            'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
            'compoeDadosLVs
        End If
    End Select
End Sub

Private Sub txtPermissao_LostFocus(Index As Integer)
    voltaCorText txtPermissao(Index)
End Sub
