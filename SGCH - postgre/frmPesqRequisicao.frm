VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesqRequisicao 
   BorderStyle     =   0  'None
   Caption         =   "Exemplo de Consulta usando o ListView"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   14895
   Icon            =   "frmPesqRequisicao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Informa��es "
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmPesqRequisicao.frx":0CCA
            Left            =   2040
            List            =   "frmPesqRequisicao.frx":0CD4
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.PictureBox picBg 
         Height          =   495
         Left            =   13680
         ScaleHeight     =   435
         ScaleMode       =   0  'User
         ScaleWidth      =   936.333
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.ImageList ImgList 
         Left            =   240
         Top             =   8160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqRequisicao.frx":0CEE
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPesqRequisicao.frx":1700
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   7
         Left            =   4440
         TabIndex        =   5
         Tag             =   "Sair"
         ToolTipText     =   "Sair"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":2112
         PICN            =   "frmPesqRequisicao.frx":212E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7695
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13573
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList"
         SmallIcons      =   "ImgList"
         ColHdrIcons     =   "ImgList"
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   6
         Left            =   3840
         TabIndex        =   7
         Tag             =   "Cancelar registro"
         ToolTipText     =   "Cancelar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":2E08
         PICN            =   "frmPesqRequisicao.frx":2E24
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   5
         Left            =   3240
         TabIndex        =   8
         Tag             =   "Editar registro"
         ToolTipText     =   "Editar registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":3AFE
         PICN            =   "frmPesqRequisicao.frx":3B1A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   4
         Left            =   2640
         TabIndex        =   9
         Tag             =   "Novo registro"
         ToolTipText     =   "Novo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":47F4
         PICN            =   "frmPesqRequisicao.frx":4810
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   3
         Left            =   2040
         TabIndex        =   10
         Tag             =   "�ltimo registro"
         ToolTipText     =   "�ltimo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":54EA
         PICN            =   "frmPesqRequisicao.frx":5506
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   2
         Left            =   1440
         TabIndex        =   11
         Tag             =   "Pr�ximo registro"
         ToolTipText     =   "Pr�ximo registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":61E0
         PICN            =   "frmPesqRequisicao.frx":61FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   1
         Left            =   840
         TabIndex        =   12
         Tag             =   "Registro anterior"
         ToolTipText     =   "Registro anterior"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":6ED6
         PICN            =   "frmPesqRequisicao.frx":6EF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Primeiro registro"
         Top             =   360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         MICON           =   "frmPesqRequisicao.frx":7BCC
         PICN            =   "frmPesqRequisicao.frx":7BE8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   495
         Index           =   8
         Left            =   10080
         TabIndex        =   14
         Tag             =   "Filtro"
         ToolTipText     =   "Filtro"
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
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
         MICON           =   "frmPesqRequisicao.frx":88C2
         PICN            =   "frmPesqRequisicao.frx":88DE
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
End
Attribute VB_Name = "frmPesqRequisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsPesquisar As New ADODB.Recordset
Private SqlPesquisar As String

Private Sub cmdconsulta_Click(Index As Integer)
'On Error GoTo Err
    Dim Y As Integer, X As Integer
    Select Case Index
    Case 0
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            ListView1.ListItems(1).Selected = True
            ListView1.ListItems(1).EnsureVisible
            ListView1.SetFocus
        End If
    Case 1
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X > 1 Then
            ListView1.ListItems(X - 1).Selected = True
            ListView1.ListItems(X - 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 2
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        If X < Y Then
            ListView1.ListItems(X + 1).Selected = True
            ListView1.ListItems(X + 1).EnsureVisible
        End If
        ListView1.SetFocus
    Case 3
        Y = ListView1.ListItems.Count
        If Y > 0 Then
            ListView1.ListItems(Y).Selected = True
            ListView1.ListItems(Y).EnsureVisible
            ListView1.SetFocus
        End If
    Case 4
        DesabBotoesN1 frmPesqRequisicao
        Pesquisa = "novo"
        frmRequisicao.Show 1
        HabBotoesN1 frmPesqRequisicao
    Case 5
        DesabBotoesN1 frmPesqRequisicao
        Pesquisa = "editar"
        AlteraListview
        If varGlobal <> "" Then frmRequisicao.Show 1
        HabBotoesN1 frmPesqRequisicao
    Case 6
        On Error GoTo Err
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            If ListView1.ListItems.Item(X).Selected = True Then
                Exit For
            End If
        Next
        varGlobal = ListView1.ListItems.Item(X)
        Pesquisa = "excluir"
        ExcluirListview
        ListView1.ListItems.Clear
        Form_Load
    Case 7
        'VerifMenu
        'HabiliIcons
        Unload Me
    Case 8
        Formulario = "Requisi��o"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then Compoe_Listview
    End Select
    Exit Sub
Err:
    MsgBox "Nenhum item selecionado", vbInformation, "SGCH"
    Exit Sub
End Sub

Private Sub cmdconsulta_MouseOver(Index As Integer)
    Legenda = cmdconsulta(Index).ToolTipText
    'MDIPrincipal.StatusBar1.Panels(3).Text = Legenda
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdConsulta_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'MDIPrincipal.StatusBar1.Panels(3).Text = Legenda
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'MDIPrincipal.StatusBar1.Panels(3).Text = Legenda
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    'frmMenu2
    Frame1.Caption = "Treinamentos"
    frmMenu2.StatusBar1.Panels(3) = Legenda
    frmPesqRequisicao.Top = frmMenu2.ACPRibbon1.Height + 290
    frmPesqRequisicao.Left = frmMenu2.Left + 130
    frmPesqRequisicao.Width = frmMenu2.Width - 300
    
    frmPesqRequisicao.Frame1.Width = frmPesqRequisicao.Width - (frmPesqRequisicao.Width * 1.5 / 100)
    frmPesqRequisicao.ListView1.Width = frmPesqRequisicao.Frame1.Width - (frmPesqRequisicao.Frame1.Width * 1.5 / 100)
    
    frmPesqRequisicao.Height = frmMenu2.Height - 2700
    frmPesqRequisicao.Frame1.Height = frmPesqRequisicao.Height - 250
    frmPesqRequisicao.ListView1.Height = frmPesqRequisicao.Frame1.Height - (frmPesqRequisicao.Frame1.Height * 15 / 90)
    
    ''MDIPrincipal
    'Frame1.Caption = "Treinamentos"
    'MDIPrincipal.StatusBar1.Panels(3) = Legenda
    'frmPesqRequisicao.Top = MDIPrincipal.pctGer.Height + (MDIPrincipal.pctGer.Height * 50 / 50)
    'frmPesqRequisicao.Left = MDIPrincipal.Left + 110
    'frmPesqRequisicao.Width = MDIPrincipal.Width - (MDIPrincipal.Width * 1.5 / 100)
    'frmPesqRequisicao.Frame1.Width = frmPesqRequisicao.Width - (frmPesqRequisicao.Width * 1.5 / 100)
    'frmPesqRequisicao.ListView1.Width = frmPesqRequisicao.Frame1.Width - (frmPesqRequisicao.Frame1.Width * 1.5 / 100)
    
    'frmPesqRequisicao.Height = MDIPrincipal.Height - (MDIPrincipal.Height * 18 / 100)
    'frmPesqRequisicao.Frame1.Height = frmPesqRequisicao.Height - (frmPesqRequisicao.Height * 1.2 / 100)
    'frmPesqRequisicao.ListView1.Height = frmPesqRequisicao.Frame1.Height - (frmPesqRequisicao.Frame1.Height * 15 / 90)
    
    AbrirTabelas
    listview_cabecalho 'Chama a Sub que monta o cabe�alho das colunas do Listview
    Combo1.Text = "Treinamento" 'Inicializa o combo com a palavra "Codigo"
    Compoe_Listview 'Chama a Sub q lista os dados no Listview
    IniciaBarra
    FecharTabelas
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esbo�o do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "C�digo", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Data Requisi��o", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Origem", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Requisitante", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Ativo", ListView1.Width / 15
    ListView1.View = lvwReport 'Modo de Exibi��o do seu Listview
End Sub

Private Sub Compoe_Listview()
    ' Declara��o de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim sql As String ' Variavel q recebe a query de conex�o com a tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer
    If FiltroGeral = "Todos" Then sql = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes"
    If FiltroGeral = "Ativos" Then sql = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes where ativo='S'"
    If FiltroGeral = "N�o ativos" Then sql = "select codrequisicao,datarequisicao,origem,nomerequisitante,ativo from  tbrequisicoes where ativo='N'"
    
    rsListview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear 'Limpa o listview
    If rsListview.RecordCount <> 0 Then frmMenu2.ProgressBar1.Max = rsListview.RecordCount
    X = 0
    While Not rsListview.EOF
        'MDIPrincipal.ProgressBar1.Value = X
        frmMenu2.ProgressBar1.Value = X
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsListview(0), "000000"))
        ItemLst.SubItems(1) = "" & rsListview.Fields(1)
        ItemLst.SubItems(2) = "" & rsListview.Fields(2)
        ItemLst.SubItems(3) = "" & rsListview.Fields(3)
        If rsListview.Fields(4) = "S" Then
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "OK"
        Else
            ItemLst.SubItems(4) = "" 'Ativo
            ItemLst.ListSubItems.Item(4).ReportIcon = "EXC"
        End If
        ItemLst.ListSubItems(4).Bold = True
        rsListview.MoveNext
        X = X + 1
    Wend
    frmMenu2.ProgressBar1.Value = 0
    Legenda = ""
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    'Ao preencher todo Listview, ele � ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
    'Fecha a conexao com a tabela Orders e limpa a mem�ria
    rsListview.Close
    Set rsListview = Nothing
End Sub

'As duas Subs abaixo faz com que ordene o listview pela coluna que vc clicar
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
    With ListView1
    If .SortKey <> Column.Index - 1 Then
        .SortKey = Column.Index - 1
        .SortOrder = lvwAscending
    Else
        If .SortOrder = lvwAscending Then
            .SortOrder = lvwDescending
        Else
            .SortOrder = lvwAscending
        End If
    End If
    .Sorted = -1
    End With
End Sub

Private Sub ListView1_DblClick()
    Pesquisa = "editar"
    AlteraListview
    frmRequisicao.Show 1
    'ListView1.ListItems.Clear
    'Form_Load
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        Pesquisar ' Sub que realiza a Pesquisa no Listview mediante ao que foi digitado no TexBox Text1 e ao q foi selecionado no ComboBox Combo1
    End If
End Sub

Private Sub Pesquisar()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count 'Conta as linhas preenchidas do Listview
    If Y > 0 Then 'Entra nessa condi��o se o Listview n�o estiver vazio
        'Nesse caso o "X" vai trabalhar como contador e
        'tamb�m ser� utilizado para percorrer as linhas do listview
        'come�ando de 1 at� o numero de linha preenchidas no Listview
        
        '----------------------------
        picBg.Width = ListView1.Width
        picBg.Height = ListView1.ListItems(1).Height * (ListView1.ListItems.Count)
        picBg.ScaleHeight = ListView1.ListItems.Count
        picBg.ScaleWidth = 1
        picBg.DrawWidth = 1
        picBg.Cls
        '----------------------------
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True 'Seleciona a linha de acordo com o valor de "X"
            'Os procedimentos abaixo ser�o realizados de acordo com o q for selecionado no ComboBox Combo1
            If Combo1.Text = "C�digo" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Codigo" em todo Listview
                If UCase(ListView1.ListItems.Item(X)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
         
                    'picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    'ListView1.Picture = picBg.Image
                    
                    ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Requisitante" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Nome" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(3)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    'picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    'ListView1.Picture = picBg.Image
                    ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "" Then
                'Se n�o for selecionado nada no ComboBox Combo1
                MsgBox "Nenhum filtro de pesquisa selecionado"
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub AbrirTabelas()
    SqlPesquisar = "Select * from tbrequisicoes order by codrequisicao"
    rsPesquisar.Open SqlPesquisar, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharTabelas()
    rsPesquisar.Close
    Set rsPesquisar = Nothing
End Sub

Private Sub ExcluirListview()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim SqlExcRequisicao As String
    Dim rsExcRequisicao As New ADODB.Recordset
    cnBanco.BeginTrans
    If MsgBox("Confirma exclus�o do Requisi��o selecionado?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        SqlExcRequisicao = "Delete from tbRequisicoes where codrequisicao= " & Val(varGlobal)
        rsExcRequisicao.Open SqlExcRequisicao, cnBanco
        MsgBox "Registro excluido com sucesso", vbInformation, "Ok!"
    End If
    cnBanco.CommitTrans
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alter��es nos registros ser�o desfeitas!", vbInformation, "Aten��o"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub AlteraListview()
    On Error GoTo Err
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    varGlobal = ListView1.ListItems.Item(X)
    Exit Sub
Err:
    MsgBox "Nenhuma Requisi��o cadastrado ou selecionado", vbInformation, "SGCH"
    Exit Sub
End Sub

Private Sub IniciaBarra()
    '-------------------------
    'Incializa o estilo do PictureBox
    '------------------------
    picBg.BackColor = ListView1.BackColor
    picBg.ScaleMode = vbTwips
    picBg.BorderStyle = vbBSNone
    picBg.AutoRedraw = True
    picBg.Visible = False
End Sub



