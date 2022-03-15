VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmProcessos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Processos"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmProcessos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcadastro 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   720
      Picture         =   "frmProcessos.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "Editar constante"
      ToolTipText     =   "Editar constante"
      Top             =   7800
      Width           =   615
   End
   Begin VB.CommandButton cmdcadastro 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   120
      Picture         =   "frmProcessos.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "Insere nova constante"
      ToolTipText     =   "Insere nova constante"
      Top             =   7800
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Processos"
      TabPicture(0)   =   "frmProcessos.frx":265E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Fases"
      TabPicture(1)   =   "frmProcessos.frx":267A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Dados do Processo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   1320
            Picture         =   "frmProcessos.frx":2696
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "Exclui constante selecionada"
            ToolTipText     =   "Exclui constante selecionada"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   720
            Picture         =   "frmProcessos.frx":3360
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "Editar constante"
            ToolTipText     =   "Editar constante"
            Top             =   840
            Width           =   615
         End
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   120
            Picture         =   "frmProcessos.frx":402A
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "Insere nova constante"
            ToolTipText     =   "Insere nova constante"
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtCadastro 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   1
            Top             =   480
            Width           =   8415
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmProcessos.frx":4CF4
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmProcessos.frx":4D66
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   5055
            Left            =   120
            TabIndex        =   2
            Top             =   1560
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   8916
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "Dados da Fase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74880
         TabIndex        =   13
         Top             =   1440
         Width           =   9375
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   7
            Left            =   1320
            Picture         =   "frmProcessos.frx":4DD2
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "Exclui constante selecionada"
            ToolTipText     =   "Exclui constante selecionada"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   6
            Left            =   720
            Picture         =   "frmProcessos.frx":5A9C
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Editar constante"
            ToolTipText     =   "Editar constante"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   5
            Left            =   120
            Picture         =   "frmProcessos.frx":6766
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "Insere nova constante"
            ToolTipText     =   "Insere nova constante"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtCadastro 
            Height          =   285
            Index           =   4
            Left            =   5640
            TabIndex        =   7
            Tag             =   "Título da FASE que irá ser exibido nos Relatórios de Inspeção de Fabricação"
            ToolTipText     =   "Título da FASE que irá ser exibido nos Relatórios de Inspeção de Fabricação"
            Top             =   480
            Width           =   3615
         End
         Begin VB.TextBox txtCadastro 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   6
            Top             =   480
            Width           =   4335
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "frmProcessos.frx":7430
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmProcessos.frx":749C
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmProcessos.frx":750E
            TabIndex        =   20
            Top             =   240
            Width           =   735
         End
         Begin VB.Frame Frame6 
            Caption         =   "% disponível "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7800
            TabIndex        =   17
            Top             =   840
            Width           =   1455
            Begin ACTIVESKINLibCtl.SkinLabel Label7 
               Height          =   375
               Left            =   240
               OleObjectBlob   =   "frmProcessos.frx":757A
               TabIndex        =   25
               Top             =   240
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   375
               Left            =   840
               OleObjectBlob   =   "frmProcessos.frx":75D2
               TabIndex        =   26
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Peso de fabricação "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5520
            TabIndex        =   16
            Top             =   840
            Width           =   2175
            Begin VB.TextBox txtCadastro 
               Height          =   285
               Index           =   3
               Left            =   480
               TabIndex        =   9
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Relatório "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4200
            TabIndex        =   15
            Top             =   840
            Width           =   1215
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "frmProcessos.frx":762A
               Left            =   120
               List            =   "frmProcessos.frx":7634
               TabIndex        =   8
               Text            =   "Não"
               Top             =   240
               Width           =   975
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3975
            Left            =   120
            TabIndex        =   10
            Top             =   1680
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   7011
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Processo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   9375
         Begin VB.CommandButton cmdcadastro 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   8880
            TabIndex        =   35
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtCadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   4
            Top             =   480
            Width           =   7575
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmProcessos.frx":7642
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmProcessos.frx":76B4
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmProcessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProcesso As New ADODB.Recordset
Private rsFase As New ADODB.Recordset
Private rsLocal As New ADODB.Recordset
Private SqlProcesso As String
Private Status As String
Private SqlFase As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirItemProcesso
    Case 1
        AlterarItem
    Case 2
        ExcluirItem
    Case 5
        IncluiTreeview
    Case 6
        mskCadastro(2).PromptInclude = False
        mskCadastro(2) = ""
        mskCadastro(2).PromptInclude = True
        txtCadastro(2) = ""
        AlteraTreeview
    Case 7
        DeletaTreeview
        CompoeTreeview
    Case 10
        Mskcadastro_GotFocus (1)
        ChamaGridProcesso
        CarregaProcesso
    Case 11
        Unload Me
    Case 12
        Bot_salvar
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    AbrirListaVer
    frmProcessos.Left = 2710
    frmProcessos.Top = 0
    SSTab1.Tab = 0
    listview_cabecalho1
    Compoe_Listview1
    mskCadastro(0).PromptInclude = False
    mskCadastro(0).Text = Format(GeraCodigo, "000")
    mskCadastro(0).PromptInclude = True
    FecharListaVer
    Status = "novo"
    CompoeTreeview
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
    carregaImagemBotao cmdCadastro(1), 1, 32 'Editar
    carregaImagemBotao cmdCadastro(2), 2, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(5), 5, 46 'Inserir
    carregaImagemBotao cmdCadastro(6), 6, 32 'Editar
    carregaImagemBotao cmdCadastro(7), 7, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(12), 12, 45 'Salvar
    carregaImagemBotao cmdCadastro(11), 11, 34 'Sair

End Sub

Private Sub AbrirListaVer()
On Error GoTo Err
    SqlProcesso = "Select * from tbProcessos Order by codprocesso"
    rsProcesso.Open SqlProcesso, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlFase = "Select * from tbfases Order by codprocesso,codfase"
    rsFase.Open SqlFase, cnBanco, adOpenKeyset, adLockOptimistic
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
End Sub

Private Sub FecharListaVer()
    rsProcesso.Close
    Set rsProcesso = Nothing
    
    rsFase.Close
    Set rsFase = Nothing
End Sub

Private Sub listview_cabecalho1()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delas e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.2
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub Compoe_Listview1()
    ' Declaração de variaveis
    Dim X As Integer
    If rsProcesso.RecordCount > 0 Then Principal.ProgressBar1.Max = rsProcesso.RecordCount
    X = 0
    While Not rsProcesso.EOF
        Principal.ProgressBar1.Value = X
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsProcesso(0), "000"))
        ItemLst.SubItems(1) = "" & rsProcesso.Fields(1)
        rsProcesso.MoveNext
        X = X + 1
    Wend
    Principal.ProgressBar1.Value = 0
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub IncluirItemProcesso()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.mskCadastro(0) Then
                AbrirListaVer
                Me.mskCadastro(0) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadastro(0)
                mskCadastro(0).PromptInclude = False
                mskCadastro(0).Text = Format(GeraCodigo, "000")
                mskCadastro(0).PromptInclude = True
                txtCadastro(0) = ""
                Y = ListView1.ListItems.Count
                FecharListaVer
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , mskCadastro(0))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , mskCadastro(0))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadastro(0)
    mskCadastro(0) = Format(Val(ListView1.ListItems.Item(Y)) + 1, "000")
    txtCadastro(0) = ""
    txtCadastro(0).SetFocus
End Sub

Private Sub AlterarItem()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.mskCadastro(0).Text = ListView1.ListItems.Item(X)
    Me.txtCadastro(0).Text = ListView1.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub ExcluirItem()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    Dim llng_Contador As Long
    
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If ListView1.ListItems.Item(X) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3) Then
            mobjMsg.Abrir "Existem  itens cadastrados para esse Processo. O Processo não pode ser excluido", Ok, critico, "Atenção"
            Exit Sub
        End If
    Next
    
    ListView1.ListItems.Remove (X)
End Sub

Private Sub Bot_salvar()
On Error GoTo Err
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
10  cnBanco.BeginTrans
    SqlSalvar = "Delete from tbProcessos"
    rsSalvar.Open SqlSalvar, cnBanco

    SqlSalvar = "Select * from tbProcessos"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = ListView1.ListItems.Item(X)
        rsSalvar.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1)
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    cnBanco.CommitTrans
    
    rsSalvar.Close
    Set rsSalvar = Nothing
    mobjMsg.Abrir "Os dados foram salvos com sucesso", Ok, informacao, "ZEUS"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from tbProcessos order by codprocesso Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProcesso.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    mskCadastro(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
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

Private Sub ListView1_DblClick()
    AlterarItem
End Sub

Private Sub CompoeTreeview()
On Error GoTo Err
    Dim rsTree As New ADODB.Recordset
    Dim SqlTree
    Dim no As Node
    Dim X As Integer, Y As Integer, Contador As Integer
    Dim vProc As String
    SqlTree = "Select tbProcessos.codprocesso, tbProcessos.descricao, tbFases.codfase, tbFases.descricao,tbFases.relger,tbfases.pesofab,tbfases.titulofase from tbProcessos,tbFases where tbFases.codprocesso=tbProcessos.codprocesso Order by tbfases.codprocesso,tbfases.codfase"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockOptimistic
    
    
    Contador = 1
    TreeView1.Nodes.Clear
    For X = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & X, Format(rsTree.Fields(0), "000") & "-" & rsTree.Fields(1))
        no.Tag = "PAI"
        no.Sorted = True
        Y = rsTree.Fields(0)
        While Y = rsTree.Fields(0)
            Set no = TreeView1.Nodes.Add("no" & X, tvwChild, Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3), Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3))
            no.Tag = "FILHOS"
            no.Sorted = True
            
            If rsTree.Fields(4) = "N" Then vProc = "Não"
            If rsTree.Fields(4) = "S" Then vProc = "Sim"
            
            TreeView1.Nodes.Add Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3), tvwChild, , ">> Relatório: " & vProc
            TreeView1.Nodes.Add Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3), tvwChild, , ">> Peso %...: " & rsTree.Fields(5)
            TreeView1.Nodes.Add Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3), tvwChild, , ">> Título...: " & rsTree.Fields(6)
            Contador = Contador + 1
            rsTree.MoveNext
            If rsTree.EOF Then Exit Sub
        Wend
    Next
    rsTree.Close
    Set rsTree = Nothing
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

Private Sub IncluiTreeview()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    SqlFase = "Select * from tbFases where tbFases.codprocesso =" & " '" & Val(Me.mskCadastro(1)) & "'" & _
    "and tbFases.codfase =" & " '" & Val(mskCadastro(2)) & "'"
    rsFase.Open SqlFase, cnBanco, adOpenKeyset, adLockOptimistic
    
    Dim vRelGer As String
    If Combo1 = "Sim" Then vRelGer = "S"
    If Combo1 = "Não" Then vRelGer = "N"
    
    If rsFase.RecordCount = 0 Then
        rsFase.AddNew
        rsFase.Fields(0) = Val(mskCadastro(1))
        rsFase.Fields(1) = Val(mskCadastro(2))
        mskCadastro(1).SetFocus
    End If
    rsFase.Fields(2) = txtCadastro(2).Text
    rsFase.Fields(3) = vRelGer
    rsFase.Fields(4) = txtCadastro(3)
    rsFase.Fields(5) = txtCadastro(4)
    rsFase.Update
    Set rsFase = Nothing
    CompoeTreeview
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

Private Sub AlteraTreeview()
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") <> 0 Then
                Status = "altera"
                mskCadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)
                mskCadastro(2) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)
                txtCadastro(2) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 5, 45)
                mskCadastro_KeyDown 1, 13, 1
            Else
                Status = "novo"
                mskCadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)
                Combo1.Text = "Não"
                mskCadastro_KeyDown 1, 13, 1
            End If
        End If
    Next
End Sub

Private Sub DeletaTreeview()
On Error GoTo Err
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Msgbox("Confirma Exclusão", vbQuestion + vbYesNo, "ZEUS") = vbYes Then
                SqlFase = "Delete from tbFase where tbFase.codprocesso =" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)) & "'" & _
                "and tbFases.codfase=" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)) & "'"
                rsFase.Open SqlFase, cnBanco, adOpenKeyset, adLockOptimistic
            End If
        End If
    Next
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

Private Function ValidaCampo()
    ValidaCampo = False
    If SSTab1.Tab = 0 Then
        If Me.txtCadastro(0) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Descrição!", Ok, critico, "Atenção"
            Me.txtCadastro(0).SetFocus
            Exit Function
        End If
    End If
    If SSTab1.Tab = 1 Then
        mskCadastro(1).PromptInclude = False
        mskCadastro(2).PromptInclude = False
        If Me.mskCadastro(1) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Código do Processo", Ok, critico, "Atenção"
            Me.mskCadastro(1).SetFocus
            mskCadastro(1).PromptInclude = True
            mskCadastro(2).PromptInclude = True
            Exit Function
        ElseIf Me.mskCadastro(2) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Código do Fase", Ok, critico, "Atenção"
            Me.mskCadastro(1).SetFocus
            mskCadastro(1).PromptInclude = True
            mskCadastro(2).PromptInclude = True
            Exit Function
        ElseIf Me.txtCadastro(2) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Descrição da Fase", Ok, critico, "Atenção"
            Me.txtCadastro(2).SetFocus
            Exit Function
        End If
    End If
    ValidaCampo = True
End Function

Private Sub Mskcadastro_GotFocus(Index As Integer)
    Dim X As Integer
    For X = 0 To mskCadastro.Count - 1
        mskCadastro(X).SelStart = 0
        mskCadastro(X).SelLength = Len(mskCadastro(X).Text)
    Next
    mskCadastro(2).PromptInclude = False
    mskCadastro(2) = ""
    mskCadastro(2).PromptInclude = True
    txtCadastro(2) = ""
End Sub

Private Sub mskCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            CarregaProcesso
        End If
    End Select
End Sub

Private Sub CarregaProcesso()
On Error GoTo Err
    Dim ContaPerc As Integer
    If Status = "altera" Then
        SqlProcesso = "Select tbProcessos.*, tbFases.codfase, tbfases.descricao,tbfases.relger,tbfases.pesofab,tbfases.titulofase from tbProcessos left join tbfases on tbfases.codprocesso = tbProcessos.codprocesso where tbProcessos.codprocesso = '" & Val(Me.mskCadastro(1)) & "'"
        rsProcesso.Open SqlProcesso, cnBanco, adOpenKeyset, adLockOptimistic
        While Not rsProcesso.EOF
            If Not IsNull(rsProcesso.Fields(5)) Then ContaPerc = ContaPerc + rsProcesso.Fields(5)
            rsProcesso.MoveNext
        Wend
        Label7 = 100 - ContaPerc
        rsProcesso.Close
        
        SqlProcesso = "Select tbProcessos.*, tbFases.codfase, tbfases.descricao,tbfases.relger,tbfases.pesofab,tbfases.titulofase from tbProcessos left join tbfases on tbfases.codprocesso = tbProcessos.codprocesso where tbProcessos.codprocesso = '" & Val(Me.mskCadastro(1)) & "'" & _
        "and tbFases.codfase=" & " '" & Val(mskCadastro(2)) & "'"
        rsProcesso.Open SqlProcesso, cnBanco, adOpenKeyset, adLockOptimistic
    Else
        SqlProcesso = "Select tbProcessos.*, tbFases.codfase, tbfases.descricao,tbfases.relger,tbfases.pesofab,tbfases.titulofase from tbProcessos left join tbfases on tbfases.codprocesso = tbProcessos.codprocesso where tbProcessos.codprocesso = '" & Val(Me.mskCadastro(1)) & "'"
        rsProcesso.Open SqlProcesso, cnBanco, adOpenKeyset, adLockOptimistic
        While Not rsProcesso.EOF
            If Not IsNull(rsProcesso.Fields(5)) Then ContaPerc = ContaPerc + rsProcesso.Fields(5)
            rsProcesso.MoveNext
        Wend
        Label7 = 100 - ContaPerc
        rsProcesso.MoveFirst
    End If
    mskCadastro(1).PromptInclude = False
    mskCadastro(2).PromptInclude = False
        
    If rsProcesso.RecordCount <> 0 Then
        mskCadastro(1).Text = Format(rsProcesso.Fields(0), "000") & ""
        txtCadastro(1).Text = rsProcesso.Fields(1)
        rsProcesso.MoveLast
        If rsProcesso.Fields(2) <> "Null" Then
            If mskCadastro(2).Text = "" Then Me.mskCadastro(2).Text = Format(rsProcesso.Fields(2) + 1, "000")
            If Status = "altera" Then
                txtCadastro(2) = rsProcesso.Fields(3)
                If rsProcesso.Fields(4) = "N" Then Combo1.Text = "Não"
                If rsProcesso.Fields(4) = "S" Then Combo1.Text = "Sim"
                If Not IsNull(rsProcesso.Fields(5)) Then txtCadastro(3) = rsProcesso.Fields(5) Else txtCadastro(3) = ""
                If Not IsNull(rsProcesso.Fields(6)) Then txtCadastro(4) = rsProcesso.Fields(6) Else txtCadastro(4) = ""
                'Combo2.Text = rsProcesso.Fields(5)
            Else
                Combo1.Text = "Não"
                txtCadastro(3) = ""
                'Combo2.Text = "Posição"
                If mskCadastro(2).Text = "" Then Me.mskCadastro(2).Text = Format(rsProcesso.Fields(2) + 1, "000")
            End If
        Else
            If mskCadastro(2).Text = "" Then Me.mskCadastro(2).Text = Format(1, "000")
        End If
        txtCadastro(2).SetFocus
    End If
    mskCadastro(1).PromptInclude = True
    mskCadastro(2).PromptInclude = True
    rsProcesso.Close
    Set rsProcesso = Nothing
    Status = "novo"
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

Private Sub ChamaGridProcesso()
On Error GoTo Err
    Dim Iposicao As Variant
    Dim F As New frmpesqger
    Sqlp = "Select * from tbProcessos order by descricao"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Processos"
    Pesquisa = frmProcessos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "descricao=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            mskCadastro(1).PromptInclude = False
            mskCadastro(1) = Val(rsLocal.Fields(0))
            mskCadastro(1).Text = Format(mskCadastro(1), "000")
            mskCadastro(1).PromptInclude = True
            txtCadastro(1).Text = rsLocal.Fields(1)
        Else
            Msgbox "Processo não cadastrado", vbInformation, "ZEUS"
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

Private Sub TreeView1_DblClick()
    mskCadastro(2).PromptInclude = False
    mskCadastro(2) = ""
    mskCadastro(2).PromptInclude = True
    txtCadastro(2) = ""
    txtCadastro(4) = ""
    AlteraTreeview
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
    Dim X As Integer
    For X = 1 To txtCadastro.Count - 1
        txtCadastro(X).SelStart = 0
        txtCadastro(X).SelLength = Len(txtCadastro(X).Text)
    Next
End Sub

