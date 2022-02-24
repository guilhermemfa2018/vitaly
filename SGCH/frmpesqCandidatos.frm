VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPesqCandidatos 
   BorderStyle     =   0  'None
   Caption         =   "Exemplo de Consulta usando o ListView"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   1365
   ClientWidth     =   14895
   Icon            =   "frmpesqCandidatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin SGCH.chameleonButton cmdconsulta 
         Height          =   495
         Index           =   9
         Left            =   10680
         TabIndex        =   15
         Tag             =   "Admitir candidato"
         ToolTipText     =   "Admitir candidato"
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
         MICON           =   "frmpesqCandidatos.frx":0CCA
         PICN            =   "frmpesqCandidatos.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   8040
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
               Picture         =   "frmpesqCandidatos.frx":19C0
               Key             =   "OK"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmpesqCandidatos.frx":23D2
               Key             =   "EXC"
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   3975
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmpesqCandidatos.frx":2DE4
            Left            =   2040
            List            =   "frmpesqCandidatos.frx":2DF1
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1815
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
         MICON           =   "frmpesqCandidatos.frx":2E09
         PICN            =   "frmpesqCandidatos.frx":2E25
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
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
         MICON           =   "frmpesqCandidatos.frx":3AFF
         PICN            =   "frmpesqCandidatos.frx":3B1B
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
         MICON           =   "frmpesqCandidatos.frx":47F5
         PICN            =   "frmpesqCandidatos.frx":4811
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
         MICON           =   "frmpesqCandidatos.frx":54EB
         PICN            =   "frmpesqCandidatos.frx":5507
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
         Tag             =   "Último registro"
         ToolTipText     =   "Último registro"
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
         MICON           =   "frmpesqCandidatos.frx":61E1
         PICN            =   "frmpesqCandidatos.frx":61FD
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
         Tag             =   "Próximo registro"
         ToolTipText     =   "Próximo registro"
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
         MICON           =   "frmpesqCandidatos.frx":6ED7
         PICN            =   "frmpesqCandidatos.frx":6EF3
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
         MICON           =   "frmpesqCandidatos.frx":7BCD
         PICN            =   "frmpesqCandidatos.frx":7BE9
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
         MICON           =   "frmpesqCandidatos.frx":88C3
         PICN            =   "frmpesqCandidatos.frx":88DF
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
         MICON           =   "frmpesqCandidatos.frx":95B9
         PICN            =   "frmpesqCandidatos.frx":95D5
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
Attribute VB_Name = "frmPesqCandidatos"
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
        DesabBotoesN1 frmPesqCandidatos
        Pesquisa = "novo"
        frmCandidatos.Show 1
        HabBotoesN1 frmPesqCandidatos
    Case 5
        DesabBotoesN1 frmPesqCandidatos
        Pesquisa = "editar"
        AlteraListview
        If varGlobal <> "" Then frmCandidatos.Show 1
        HabBotoesN1 frmPesqCandidatos
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
        Formulario = "Candidatos"
        MontaFiltro
        frmFiltro.Show 1
        If TiPo = True Then Compoe_Listview
    Case 9
        DesabBotoesN1 frmPesqCandidatos
        'Pesquisa = "editar"
        AlteraListview
        If varGlobal <> "" Then frmAdmitirCandidato.Show 1
        HabBotoesN1 frmPesqCandidatos
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
    Frame1.Caption = "Candidatos"
    frmMenu2.StatusBar1.Panels(3) = Legenda
    frmPesqCandidatos.Top = frmMenu2.ACPRibbon1.Height + 290
    frmPesqCandidatos.Left = frmMenu2.Left + 130
    frmPesqCandidatos.Width = frmMenu2.Width - 300
    
    frmPesqCandidatos.Frame1.Width = frmPesqCandidatos.Width - (frmPesqCandidatos.Width * 1.5 / 100)
    frmPesqCandidatos.ListView1.Width = frmPesqCandidatos.Frame1.Width - (frmPesqCandidatos.Frame1.Width * 1.5 / 100)
    
    frmPesqCandidatos.Height = frmMenu2.Height - 2700
    frmPesqCandidatos.Frame1.Height = frmPesqCandidatos.Height - 250
    frmPesqCandidatos.ListView1.Height = frmPesqCandidatos.Frame1.Height - (frmPesqCandidatos.Frame1.Height * 15 / 90)
    
    ''MDIPrincipal
    'Frame1.Caption = "Candidatos"
    'MDIPrincipal.StatusBar1.Panels(3) = Legenda
    'frmPesqCandidatos.Top = MDIPrincipal.pctGer.Height + (MDIPrincipal.pctGer.Height * 50 / 50)
    'frmPesqCandidatos.Left = MDIPrincipal.Left + 110
    'frmPesqCandidatos.Width = MDIPrincipal.Width - (MDIPrincipal.Width * 1.5 / 100)
    'frmPesqCandidatos.Frame1.Width = frmPesqCandidatos.Width - (frmPesqCandidatos.Width * 1.5 / 100)
    'frmPesqCandidatos.ListView1.Width = frmPesqCandidatos.Frame1.Width - (frmPesqCandidatos.Frame1.Width * 1.5 / 100)
    
    'frmPesqCandidatos.Height = MDIPrincipal.Height - (MDIPrincipal.Height * 18 / 100)
    'frmPesqCandidatos.Frame1.Height = frmPesqCandidatos.Height - (frmPesqCandidatos.Height * 1.2 / 100)
    'frmPesqCandidatos.ListView1.Height = frmPesqCandidatos.Frame1.Height - (frmPesqCandidatos.Frame1.Height * 15 / 90)
    
    AbrirTabelas
    listview_cabecalho 'Chama a Sub que monta o cabeçalho das colunas do Listview
    Combo1.Text = "Nome" 'Inicializa o combo com a palavra "Nome"
    Compoe_Listview 'Chama a Sub q lista os dados no Listview
    IniciaBarra
    FecharTabelas
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "CPF", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "CTPS nº", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Série", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Média", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Ativo", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "C. Avaliadas", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Cargos", ListView1.Width / 5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub Compoe_Listview()
    ' Declaração de variaveis
    Dim rsListview As New ADODB.Recordset ' Variavel que vai receber os dados da tabela
    Dim sql As String ' Variavel q recebe a query de conexão com a tabela
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim X As Integer
    
    If FiltroGeral = "Todos" Then sql = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.ativo='S'"
    If FiltroGeral = "Ativos" Then sql = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo='S' and b.ativo='S'"
    If FiltroGeral = "Não ativos" Then sql = "Select a.cpf,a.nomecolaborador,a.ctpsnumero,a.ctpsserie,a.mediageral,a.ativo,a.compav,d.nomecargo from tbcolaboradores as a inner join tbcolaboradoresHist as b on a.cpf = b.cpf and a.tipo = b.tipo and a.tipo = 'candidato' inner join tbMatriz as c on c.codmatriz=b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where a.ativo is null  and b.ativo='S' or a.ativo='N' and b.ativo='S'"
    
    
'    If FiltroGeral = "Todos" Then sql = "select * from tbColaboradores where tipo = 'candidato'"
'    If FiltroGeral = "Ativos" Then sql = "select * from tbColaboradores where ativo = 'S' and tipo = 'candidato'"
'    If FiltroGeral = "Não ativos" Then sql = "select * from tbColaboradores where ativo is null or ativo = 'N' and tipo = 'candidato'"
    
    rsListview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear 'Limpa o listview
    'If rsListview.RecordCount <> 0 Then MDIPrincipal.ProgressBar1.Max = rsListview.RecordCount
    If rsListview.RecordCount <> 0 Then frmMenu2.ProgressBar1.Max = rsListview.RecordCount
    X = 0
    While Not rsListview.EOF
        'MDIPrincipal.ProgressBar1.Value = X
        frmMenu2.ProgressBar1.Value = X
        Set ItemLst = ListView1.ListItems.Add(, , rsListview(0))
        ItemLst.SubItems(1) = "" & rsListview.Fields(1)
        ItemLst.SubItems(2) = "" & rsListview.Fields(2)
        ItemLst.SubItems(3) = "" & rsListview.Fields(3)
        ItemLst.SubItems(4) = "" & Format(rsListview.Fields(4), "#,##0.00;(#,##0.00)") & "%"
        If rsListview.Fields(5) = "S" Then
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "OK"
        Else
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "EXC"
        End If
        ItemLst.SubItems(6) = "" & rsListview.Fields(6)
        ItemLst.SubItems(7) = "" & rsListview.Fields(7)
        
        ItemLst.ListSubItems(4).Bold = True
        If rsListview.Fields(4) >= 70 Then
            ItemLst.ListSubItems(4).ForeColor = &H8000&
        Else
            ItemLst.ListSubItems(4).ForeColor = &HC0&
        End If
        
        rsListview.MoveNext
        X = X + 1
    Wend
    'MDIPrincipal.ProgressBar1.Value = 0
    frmMenu2.ProgressBar1.Value = 0
    Legenda = ""
    'MDIPrincipal.StatusBar1.Panels(3).Text = Legenda
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    
    'Fecha a conexao com a tabela Orders e limpa a memória
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
    frmCandidatos.Show 1
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
    If Y > 0 Then 'Entra nessa condição se o Listview não estiver vazio
        'Nesse caso o "X" vai trabalhar como contador e
        'também será utilizado para percorrer as linhas do listview
        'começando de 1 até o numero de linha preenchidas no Listview
        
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
            'Os procedimentos abaixo serão realizados de acordo com o q for selecionado no ComboBox Combo1
            If Combo1.Text = "CPF" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Codigo" em todo Listview
                If UCase(ListView1.ListItems.Item(X)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
         
                    'picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    'ListView1.Picture = picBg.Image
                    
                    ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "Nome" Then
                'Compara o que foi digitado no TextBox Text1 com a coluna "Estabelecimento" em todo Listview
                If UCase(ListView1.SelectedItem.ListSubItems.Item(2)) Like UCase(Me.Text1.Text & "*") Then
                    ListView1.ListItems(X).Selected = True
                    ListView1.ListItems(X).EnsureVisible
                    
                    'picBg.Line (0, X - 1)-(1, X), &HC0FFC0, BF
                    'ListView1.Picture = picBg.Image
                    ListView1.SetFocus
                    Exit Sub
                End If
            ElseIf Combo1.Text = "" Then
                'Se não for selecionado nada no ComboBox Combo1
                MsgBox "Nenhum filtro de pesquisa selecionado"
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub AbrirTabelas()
    SqlPesquisar = "Select * from tbColaboradores where tipo = 'candidato' order by cpf"
    rsPesquisar.Open SqlPesquisar, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharTabelas()
    rsPesquisar.Close
    Set rsPesquisar = Nothing
End Sub

Private Sub ExcluirListview()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim SqlExcColaborador As String
    Dim rsExcColaborador As New ADODB.Recordset
    cnBanco.BeginTrans
    If MsgBox("Confirma exclusão do cargo selecionado?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
        SqlExcColaborador = "Delete from tbColaboradores where cpf= " & Val(varGlobal) & "' and tipo= 'candidato'"
        rsExcColaborador.Open SqlExcCargo, cnBanco
        MsgBox "Registro excluido com sucesso", vbInformation, "Ok!"
    End If
    cnBanco.CommitTrans
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
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
    MsgBox "Nenhum Cargo cadastrado ou selecionado", vbInformation, "SGCH"
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


