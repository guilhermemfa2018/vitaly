VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmItemVerif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de itens de verificação"
   ClientHeight    =   8280
   ClientLeft      =   2640
   ClientTop       =   1815
   ClientWidth     =   9105
   Icon            =   "frmItemVerif.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Grupos"
      TabPicture(0)   =   "frmItemVerif.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Itens"
      TabPicture(1)   =   "frmItemVerif.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Grupo "
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   8655
         Begin VB.TextBox txtCadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   7
            Top             =   480
            Width           =   6855
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   6
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmItemVerif.frx":0D02
            TabIndex        =   24
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmItemVerif.frx":0D74
            TabIndex        =   23
            Top             =   240
            Width           =   735
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   255
            Index           =   10
            Left            =   8160
            TabIndex        =   8
            Top             =   480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BTYPE           =   4
            TX              =   "..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmItemVerif.frx":0DE0
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
      Begin VB.Frame Frame2 
         Caption         =   "Dados do item "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   8655
         Begin VB.TextBox txtCadastro 
            Height          =   285
            Index           =   3
            Left            =   7680
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   7680
            OleObjectBlob   =   "frmItemVerif.frx":0DFC
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtCadastro 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   480
            Width           =   6375
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   9
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmItemVerif.frx":0E66
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmItemVerif.frx":0ED8
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4095
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   7223
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   7
            Left            =   1320
            TabIndex        =   13
            Top             =   840
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
            MICON           =   "frmItemVerif.frx":0F44
            PICN            =   "frmItemVerif.frx":0F60
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   6
            Left            =   720
            TabIndex        =   12
            Top             =   840
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
            MICON           =   "frmItemVerif.frx":1C3A
            PICN            =   "frmItemVerif.frx":1C56
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   840
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
            MICON           =   "frmItemVerif.frx":2930
            PICN            =   "frmItemVerif.frx":294C
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
      Begin VB.Frame Frame1 
         Caption         =   "Dados do grupo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   8655
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmItemVerif.frx":3626
            Left            =   5160
            List            =   "frmItemVerif.frx":3633
            TabIndex        =   28
            Text            =   "-"
            Top             =   480
            Width           =   2895
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   5160
            OleObjectBlob   =   "frmItemVerif.frx":364F
            TabIndex        =   27
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtCadastro 
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   1
            Top             =   480
            Width           =   4215
         End
         Begin MSMask.MaskEdBox mskCadastro 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmItemVerif.frx":36E1
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmItemVerif.frx":3753
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   2
            Left            =   1320
            TabIndex        =   4
            Top             =   960
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
            MICON           =   "frmItemVerif.frx":37BF
            PICN            =   "frmItemVerif.frx":37DB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   1
            Left            =   720
            TabIndex        =   3
            Top             =   960
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
            MICON           =   "frmItemVerif.frx":44B5
            PICN            =   "frmItemVerif.frx":44D1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdcadastro 
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   960
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
            MICON           =   "frmItemVerif.frx":51AB
            PICN            =   "frmItemVerif.frx":51C7
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
            Height          =   5055
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8916
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
   Begin ZEUS.chameleonButton cmdcadastro 
      Height          =   615
      Index           =   11
      Left            =   720
      TabIndex        =   19
      Top             =   7560
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
      MICON           =   "frmItemVerif.frx":5EA1
      PICN            =   "frmItemVerif.frx":5EBD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdcadastro 
      Height          =   615
      Index           =   12
      Left            =   120
      TabIndex        =   20
      Top             =   7560
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
      MICON           =   "frmItemVerif.frx":6B97
      PICN            =   "frmItemVerif.frx":6BB3
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
Attribute VB_Name = "frmItemVerif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsGrupo As New ADODB.Recordset
Private rsItem As New ADODB.Recordset
Private rsLocal As New ADODB.Recordset
Private SqlGrupo As String
Private SqlItem As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirItemGrupo
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
        ChamaGridGrupo
        CarregaGrupo
    Case 11
        Unload Me
    Case 12
        Bot_salvar
    End Select
End Sub

Private Sub Form_Load()
    AbrirListaVer
    frmItemVerif.Left = 2710
    frmItemVerif.Top = 0
    SSTab1.Tab = 0
    listview_cabecalho1
    Compoe_Listview1
    mskCadastro(0).PromptInclude = False
    mskCadastro(0).Text = Format(GeraCodigo, "000")
    mskCadastro(0).PromptInclude = True
    FecharListaVer
    CompoeTreeview
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub AbrirListaVer()
On Error GoTo Err
    SqlGrupo = "Select * from tbVerifGrupo Order by codgrupo"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlItem = "Select * from tbVerifItem Order by codgrupo,coditem"
    rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
End Sub

Private Sub FecharListaVer()
    rsGrupo.Close
    Set rsGrupo = Nothing
    
    rsItem.Close
    Set rsItem = Nothing
End Sub

Private Sub listview_cabecalho1()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delas e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.8
    ListView1.ColumnHeaders.Add , , "Aplicação", ListView1.Width / 6
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub Compoe_Listview1()
    ' Declaração de variaveis
    Dim X As Integer
    If rsGrupo.RecordCount > 0 Then Principal.ProgressBar1.Max = rsGrupo.RecordCount
    X = 0
    While Not rsGrupo.EOF
        Principal.ProgressBar1.Value = X
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsGrupo(0), "000"))
        ItemLst.SubItems(1) = "" & rsGrupo.Fields(1)
        ItemLst.SubItems(2) = "" & rsGrupo.Fields(2)
        rsGrupo.MoveNext
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

Private Sub IncluirItemGrupo()
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
                ListView1.SelectedItem.ListSubItems.Item(2) = Combo1.Text
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
    ItemLst.SubItems(2) = Combo1.Text
    mskCadastro(0) = Format(Val(ListView1.ListItems.Item(Y)) + 1, "000")
    txtCadastro(0) = ""
    Combo1.Text = "-"
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
    Me.Combo1.Text = ListView1.SelectedItem.ListSubItems.Item(2)
End Sub

Private Sub ExcluirItem()
On Error GoTo Err
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
            mobjMsg.Abrir "Existem  itens cadastrados para esse Grupo. O Grupo não pode ser excluido", Ok, informacao, "ZEUS"
            Exit Sub
        End If
    Next
    ListView1.ListItems.Remove (X)
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

Private Sub Bot_salvar()
On Error GoTo Err
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
10  cnBanco.BeginTrans
    SqlSalvar = "Delete from tbVerifGrupo"
    rsSalvar.Open SqlSalvar, cnBanco

    SqlSalvar = "Select * from tbVerifGrupo"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = ListView1.ListItems.Item(X)
        rsSalvar.Fields(1) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
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
        mobjMsg.Abrir "OOcorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "ZEUS"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Function GeraCodigo()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from tbVerifGrupo order by codgrupo Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGrupo.RecordCount > 0 Then
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
    Dim X As Integer, Y As Integer
    SqlTree = "Select a.codgrupo,a.descricao,b.coditem,b.descricao,b.sigla from tbVerifGrupo as a Inner join tbVerifItem as b on a.codgrupo = b.codgrupo Order by b.codgrupo,b.coditem"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockReadOnly
    
    TreeView1.Nodes.Clear
    For X = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & X, Format(rsTree.Fields(0), "000") & "-" & rsTree.Fields(1))
        Y = rsTree.Fields(0)
        While Y = rsTree.Fields(0)
            TreeView1.Nodes.Add "no" & X, tvwChild, , Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3) & "(" & rsTree.Fields(4) & ")"
            rsTree.MoveNext
            If rsTree.EOF Then Exit Sub
        Wend
        'TreeView1.Nodes(X).Expanded = True
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
    SqlItem = "Select * from tbVerifItem where tbVerifitem.codgrupo =" & " '" & Val(Me.mskCadastro(1)) & "'" & _
    "and tbVerifItem.coditem=" & " '" & Val(mskCadastro(2)) & "'"
    rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
    If rsItem.RecordCount = 0 Then
        rsItem.AddNew
        rsItem.Fields(0) = Val(mskCadastro(1))
        rsItem.Fields(1) = Val(mskCadastro(2))
        mskCadastro(1).SetFocus
    End If
    rsItem.Fields(2) = txtCadastro(2).Text
    rsItem.Fields(3) = txtCadastro(3).Text
    rsItem.Update
    Set rsItem = Nothing
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
                'MsgBox "Subitem"
                mskCadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)
                mskCadastro(2) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)
                
                txtCadastro(3) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "("), 5)
                txtCadastro(3) = RemoveMask2(txtCadastro(3), "(")
                txtCadastro(3) = RemoveMask2(txtCadastro(3), ")")
                
                
                
                txtCadastro(2) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 5, 45)
                txtCadastro(2) = RemoveMask2(txtCadastro(2), "(")
                txtCadastro(2) = RemoveMask2(txtCadastro(2), ")")
                txtCadastro(2) = RemoveMask2(txtCadastro(2), txtCadastro(3))
                
                mskCadastro_KeyDown 1, 13, 1
            Else
                'MsgBox "Grupo"
                mskCadastro(1) = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)
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
                SqlItem = "Delete from tbVerifItem where tbVerifitem.codgrupo =" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)) & "'" & _
                "and tbVerifItem.coditem=" & " '" & Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)) & "'"
                rsItem.Open SqlItem, cnBanco, adOpenKeyset, adLockOptimistic
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
            Msgbox "Favor preencher o campo Descrição!", vbInformation, "Atenção"
            Me.txtCadastro(0).SetFocus
            Exit Function
        End If
    End If
    If SSTab1.Tab = 1 Then
        mskCadastro(1).PromptInclude = False
        mskCadastro(2).PromptInclude = False
        If Me.mskCadastro(1) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Código do Grupo", Ok, critico, "Atenção"
            Me.mskCadastro(1).SetFocus
            mskCadastro(1).PromptInclude = True
            mskCadastro(2).PromptInclude = True
            Exit Function
        ElseIf Me.mskCadastro(2) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Código do Item", Ok, critico, "Atenção"
            Me.mskCadastro(1).SetFocus
            mskCadastro(1).PromptInclude = True
            mskCadastro(2).PromptInclude = True
            Exit Function
        ElseIf Me.txtCadastro(2) = "" Then
            mobjMsg.Abrir "Favor preencher o campo Descrição do Item", Ok, critico, "Atenção"
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
            CarregaGrupo
        End If
    End Select

End Sub

Private Sub CarregaGrupo()
    SqlGrupo = "Select tbVerifGrupo.*, tbVerifItem.coditem, tbVerifItem.descricao from tbVerifGrupo left join tbVerifItem on tbVerifItem.codgrupo = tbVerifGrupo.codgrupo where tbVerifGrupo.codgrupo = '" & Val(Me.mskCadastro(1)) & "'"
    rsGrupo.Open SqlGrupo, cnBanco, adOpenKeyset, adLockOptimistic
    mskCadastro(1).PromptInclude = False
    mskCadastro(2).PromptInclude = False
        
    If rsGrupo.RecordCount = 0 Then
        mskCadastro(1).Text = Format(mskCadastro(1), "000") & ""
        txtCadastro(1).Text = ""
        Msgbox "Grupo não cadastrado"
        mskCadastro(1).SetFocus
    Else
        mskCadastro(1).Text = Format(rsGrupo.Fields(0), "000") & ""
        txtCadastro(1).Text = rsGrupo.Fields(1)
        rsGrupo.MoveLast
        If rsGrupo.Fields(3) <> "Null" Then
            If mskCadastro(2).Text = "" Then Me.mskCadastro(2).Text = Format(rsGrupo.Fields(3) + 1, "000")
        Else
            If mskCadastro(2).Text = "" Then Me.mskCadastro(2).Text = Format(1, "000")
        End If
        txtCadastro(2).SetFocus
    End If
    mskCadastro(1).PromptInclude = True
    mskCadastro(2).PromptInclude = True
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub

Private Sub ChamaGridGrupo()
On Error GoTo Err
    Dim Iposicao As Variant
    Dim F As New frmpesqger
    Sqlp = "Select * from tbVerifGrupo order by descricao"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Grupos"
    Pesquisa = frmItemVerif.Tag
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
            mobjMsg.Abrir "Grupo não cadastrado", Ok, critico, "Atenção"
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
    AlteraTreeview
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
    Dim X As Integer
    For X = 1 To txtCadastro.Count - 1
        txtCadastro(X).SelStart = 0
        txtCadastro(X).SelLength = Len(txtCadastro(X).Text)
    Next
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            IncluirItemGrupo
        End If
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            IncluiTreeview
            txtCadastro(2).SetFocus
        End If
    End Select

End Sub
