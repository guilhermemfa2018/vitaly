VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFCECons 
   Caption         =   "FCE - Ficha de Controle de Encomenda (Planejamento)"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFCECons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9960
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   60
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   2400
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.CommandButton cmdCadastro 
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
      Index           =   9
      Left            =   120
      Picture         =   "frmFCECons.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   59
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   9360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da FCE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   345
         Left            =   2880
         TabIndex        =   1
         Tag             =   "Nº da carta proposta"
         Top             =   480
         Width           =   7815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1200
         TabIndex        =   32
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
         Format          =   163577857
         CurrentDate     =   40449
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFCECons.frx":1994
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   2880
         OleObjectBlob   =   "frmFCECons.frx":19F8
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmFCECons.frx":1A6E
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFCECons.frx":1AE2
         TabIndex        =   56
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   14208
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
      TabCaption(0)   =   "Cliente"
      TabPicture(0)   =   "frmFCECons.frx":1B48
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmFCECons.frx":1B64
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame13"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "FO's"
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
         TabIndex        =   30
         Top             =   6600
         Width           =   5655
         Begin MSComctlLib.ListView ListView1 
            Height          =   855
            Left            =   120
            TabIndex        =   31
            Tag             =   "FO(s) que compõe(m) a FCE"
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   1508
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
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
      End
      Begin VB.Frame Frame13 
         Caption         =   "Escopo de Fornecimento "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   -74880
         TabIndex        =   28
         Top             =   420
         Width           =   10575
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   6015
            Left            =   120
            TabIndex        =   29
            Tag             =   "Escopo de fornecimento"
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   10610
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
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
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Observações Técnicas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   5880
         TabIndex        =   26
         Top             =   3120
         Width           =   4815
         Begin VB.TextBox Text18 
            Enabled         =   0   'False
            Height          =   4335
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Tag             =   "Observações técnicas"
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente "
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
         TabIndex        =   14
         Top             =   480
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   9
            Left            =   120
            TabIndex        =   16
            Tag             =   "Email"
            Top             =   3360
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   10
            Left            =   2880
            TabIndex        =   15
            Tag             =   "Site"
            Top             =   3360
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   8
            Left            =   3240
            TabIndex        =   17
            Tag             =   "Fax"
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   7
            Left            =   960
            TabIndex        =   18
            Tag             =   "Telefone"
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Tag             =   "Estado"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Tag             =   "Bairro"
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   5
            Left            =   2880
            TabIndex        =   20
            Tag             =   "Cidade"
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   3
            Left            =   4560
            TabIndex        =   22
            Tag             =   "CEP"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Tag             =   "Endereço"
            Top             =   1200
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            Tag             =   "Nome do cliente"
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Tag             =   "Código do Cliente"
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmFCECons.frx":1B80
            TabIndex        =   45
            Top             =   3120
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":1BE2
            TabIndex        =   44
            Top             =   3120
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "frmFCECons.frx":1C46
            TabIndex        =   43
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmFCECons.frx":1CA6
            TabIndex        =   42
            Top             =   2400
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":1D10
            TabIndex        =   41
            Top             =   2400
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmFCECons.frx":1D76
            TabIndex        =   40
            Top             =   1680
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":1DDC
            TabIndex        =   39
            Top             =   1680
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   4560
            OleObjectBlob   =   "frmFCECons.frx":1E42
            TabIndex        =   38
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":1EA2
            TabIndex        =   37
            Top             =   960
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmFCECons.frx":1F0C
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":1F6E
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Dados do Contato "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   14
            Left            =   2040
            TabIndex        =   10
            Tag             =   "Email"
            Top             =   1200
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "frmFCECons.frx":1FD4
            TabIndex        =   49
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   13
            Left            =   120
            TabIndex        =   11
            Tag             =   "Telefone"
            Top             =   1200
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":2038
            TabIndex        =   48
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   12
            Left            =   1200
            TabIndex        =   12
            Tag             =   "Nome do contato"
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   11
            Left            =   120
            TabIndex        =   13
            Tag             =   "Código do Contato"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFCECons.frx":20A2
            TabIndex        =   47
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":2104
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Escopo de fornecimento "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   4815
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   19
            Left            =   2400
            TabIndex        =   8
            Tag             =   "Responsável pela pintura"
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   18
            Left            =   120
            TabIndex        =   7
            Tag             =   "Responsável pelo transporte"
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   17
            Left            =   2400
            TabIndex        =   6
            Tag             =   "Responsável pelo fornecimento da matéria prima"
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   16
            Left            =   120
            TabIndex        =   5
            Tag             =   "Responsável peloreparo"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   15
            Left            =   1680
            TabIndex        =   4
            Tag             =   "Responsável pela fabricação"
            Top             =   480
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   345
            Left            =   120
            TabIndex        =   33
            Tag             =   "Data de entrega do escopo de fornecimento"
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   141099009
            CurrentDate     =   40449
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   2400
            OleObjectBlob   =   "frmFCECons.frx":216A
            TabIndex        =   55
            Top             =   1680
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":21D2
            TabIndex        =   54
            Top             =   1680
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   2400
            OleObjectBlob   =   "frmFCECons.frx":2240
            TabIndex        =   53
            Top             =   960
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":22B2
            TabIndex        =   52
            Top             =   960
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "frmFCECons.frx":2318
            TabIndex        =   51
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCECons.frx":2386
            TabIndex        =   50
            Top             =   240
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmFCECons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTreeview As New ADODB.Recordset

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 9
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    inicializa_tabs SSTab1, Picture1
    'If varGlobal = "-" Or varGlobal = "" Then
    '    GoTo ErrHandler
    'End If
    SSTab1.Tab = 0
    DTPicker1 = Date
    DTPicker2 = Date
    Label2 = varGlobal
    CompoeTreeview
    CompoeControles
    carregarIconBotao
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadastro(9), 9, 34 'Sair
End Sub

Private Sub CompoeTreeview()
On Error GoTo Err
    Dim rsTree As New ADODB.Recordset
    Dim SqlTree
    Dim no As Node
    Dim x As Integer, y As Integer
    SqlTree = "Select tbVerifGrupo.codgrupo, tbVerifGrupo.descricao, tbVerifItem.coditem, tbVerifItem.descricao from tbVerifGrupo,tbVerifItem where tbVerifItem.codgrupo=tbVerifGrupo.codgrupo Order by tbVerifItem.codgrupo,tbVerifItem.coditem"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockOptimistic
    
    TreeView1.Nodes.Clear
    For x = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & x, Format(rsTree.Fields(0), "000") & "-" & rsTree.Fields(1))
        y = rsTree.Fields(0)
        While y = rsTree.Fields(0)
            TreeView1.Nodes.Add "no" & x, tvwChild, , Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3)
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

Private Sub TreeView1_Click()
    Dim i As Integer
    With TreeView1
        For i = 1 To .Nodes.Count
            If .Nodes(i).Selected = True Then
                If .Nodes(i).Checked = True Then
                    .Nodes(i).Checked = True
                ElseIf .Nodes(i).Checked = False Then
                    .Nodes(i).Checked = False
                End If
            End If
        Next i
    End With
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    With TreeView1
        For i = 1 To .Nodes.Count
            If Not .Nodes(i).Parent Is Nothing Then
                If .Nodes(i).Parent.Key = Node.Key Then
                    .Nodes(i).Checked = Node.Checked
                End If
            End If
        Next i
    End With
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim llng_Contador As Long
    Dim SqlTreeview As String
    Dim y As Integer, x As Integer, i As Integer
    
    Dim rsFCE As New ADODB.Recordset
    Dim rsClientes As New ADODB.Recordset
    Dim rsContatos As New ADODB.Recordset
    Dim sqlFCE As String
    Dim sqlClientes As String
    Dim sqlContatos As String

    sqlFCE = "select a.fce,a.dataabertura,a.cartaproposta,a.observacao,a.obscomercial,a.obsfinanceira,a.dataentrega,a.fabricacao,a.reparo, a.materiaprima, " & _
    "a.transporte,a.pintura,a.databook,b.codclifor,b.codcontato from tbFCE as a, tbfo as b where a.fce = b.fce and a.FCE = '" & Val(varGlobal) & "'"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsFCE.RecordCount > 0 Then
        txtcadastro(0) = rsFCE.Fields(13)
        If Not IsNull(rsFCE.Fields(14)) Then txtcadastro(11) = rsFCE.Fields(14)
        DTPicker1 = rsFCE.Fields(1)
        DTPicker2 = rsFCE.Fields(6)
        txtcadastro(15) = rsFCE.Fields(7)
        txtcadastro(16) = rsFCE.Fields(8)
        txtcadastro(17) = rsFCE.Fields(9)
        txtcadastro(18) = rsFCE.Fields(10)
        txtcadastro(19) = rsFCE.Fields(11)
        Text18 = rsFCE.Fields(5)
        Text1 = rsFCE.Fields(2)
    End If
    CarregaCli
    CarregaContato
    ContFOSel
    
    SqlTreeview = "Select * from tbListaVerif where tbListaVerif.fce = '" & Val(Me.Label2) & "'"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount > 0 Then
        While Not rsTreeview.EOF
            For llng_Contador = 1 To TreeView1.Nodes.Count
                TreeView1.Nodes(llng_Contador).Expanded = True
                If rsTreeview.Fields(1) = Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)) And rsTreeview.Fields(2) = Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)) Then
                    TreeView1.Nodes(llng_Contador).Checked = True
                End If
            Next
            rsTreeview.MoveNext
        Wend
    End If
    rsTreeview.Close
    Set rsTreeview = Nothing
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

Private Sub CarregaCli()
On Error GoTo Err
    Dim rsCli As New ADODB.Recordset
    Dim SqlCli As String
    SqlCli = "Select * from tbclifor where tbclifor.codclifor = '" & Val(txtcadastro(0)) & "'"
    rsCli.Open SqlCli, cnBanco, adOpenKeyset, adLockOptimistic
    If rsCli.EOF Then
        'Msgbox "Cliente não cadastrado", vbInformation, "Zeus"
        rsCli.Close
        Set rsCli = Nothing
        Exit Sub
    End If
    txtcadastro(0).Text = Format(rsCli.Fields(0), "000000")
    txtcadastro(1).Text = rsCli.Fields(13)
    txtcadastro(2).Text = rsCli.Fields(1)
    txtcadastro(3).Text = rsCli.Fields(2)
    txtcadastro(4).Text = rsCli.Fields(3)
    txtcadastro(5).Text = rsCli.Fields(4)
    txtcadastro(6).Text = rsCli.Fields(5)
    txtcadastro(7).Text = Format(rsCli.Fields(6), "(##)####-####")
    txtcadastro(8).Text = Format(rsCli.Fields(7), "(##)####-####")
    txtcadastro(9).Text = rsCli.Fields(8)
    txtcadastro(10).Text = rsCli.Fields(9)
    rsCli.Close
    Set rsCli = Nothing
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

Private Sub CarregaContato()
On Error GoTo Err
    Dim rsContato As New ADODB.Recordset
    Dim SqlContato As String
    
    SqlContato = "Select * from tbcontatos where tbcontatos.codclifor= '" & Val(txtcadastro(0)) & "'" & _
    "and tbcontatos.codcontato=" & " '" & Val(txtcadastro(11)) & "'order by nome"
    
    rsContato.Open SqlContato, cnBanco, adOpenKeyset, adLockOptimistic
    If rsContato.EOF Then
'        MsgBox "Contato não cadastrado", vbInformation, "Zeus"
        rsContato.Close
        Set rsContato = Nothing
        Exit Sub
    End If
    txtcadastro(11).Text = Format(rsContato.Fields(1), "000000")
    txtcadastro(12).Text = rsContato.Fields(2)
    txtcadastro(13).Text = Format(rsContato.Fields(6), "(##)####-####")
    txtcadastro(14).Text = rsContato.Fields(9)
    rsContato.Close
    Set rsContato = Nothing
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

Private Sub ContFOSel()
On Error GoTo Err
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "FO", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.3
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview

    Dim rsLV As New ADODB.Recordset
    Dim SqlLV As String
    Dim y As Integer, codfornec As Integer
    Dim numFCE As String
    y = vListViewPrincipal.ListItems.Count
    
    For x = 1 To y
        If vListViewPrincipal.ListItems.Item(x).Selected = True Then
            numFCE = vListViewPrincipal.SelectedItem.ListSubItems.Item(1)
            Exit For
        End If
    Next
    If numFCE = "-" Then
        mobjMsg.Abrir "Nenhuma FCE Selecionada", Ok, critico, "Atenção"
        Exit Sub
    End If
    SqlLV = "select codfo,fce,descricao from tbfo where fce = '" & numFCE & "'" '& "'order by codfo"
    rsLV.Open SqlLV, cnBanco, adOpenKeyset, adLockOptimistic
    
    While Not rsLV.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsLV(0), "000000"))
        ItemLst.SubItems(1) = "" & rsLV.Fields(2)
        rsLV.MoveNext
    Wend
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

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
End Sub
