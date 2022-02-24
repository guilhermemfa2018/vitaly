VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmRequisicao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisição de pessoal"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frmRequisicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Observação "
      Height          =   1575
      Left            =   6720
      TabIndex        =   42
      Top             =   1200
      Width           =   3735
      Begin VB.TextBox txtCadReq 
         Height          =   1215
         Index           =   13
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9360
      TabIndex        =   29
      Top             =   8640
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Tag             =   "Status do curso/treinamento"
         ToolTipText     =   "Status do curso/treinamento"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Cargos requisitados"
      TabPicture(0)   =   "frmRequisicao.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label25"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label28"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label30"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label11"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label31"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCadastro(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCadastro(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCadastro(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdCadastro(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ListView1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCadReq(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCadReq(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCadReq(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCadReq(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCadReq(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCadReq(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCadReq(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DTPicker2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCadReq(16)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdCad(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Aprovação"
      TabPicture(1)   =   "frmRequisicao.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCad(2)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "cboCadRequisicao(2)"
      Tab(1).Control(3)=   "ListView2"
      Tab(1).Control(4)=   "txtCadReq(15)"
      Tab(1).Control(5)=   "txtCadReq(14)"
      Tab(1).Control(6)=   "cmdCadastro(9)"
      Tab(1).Control(7)=   "cmdCadastro(8)"
      Tab(1).Control(8)=   "cmdCadastro(7)"
      Tab(1).Control(9)=   "cmdCadastro(6)"
      Tab(1).Control(10)=   "Label16"
      Tab(1).Control(11)=   "Label14"
      Tab(1).Control(12)=   "Label13"
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   45
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   57
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificador"
         Height          =   615
         Left            =   -71880
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   7800
         TabIndex        =   13
         ToolTipText     =   "Nível do cargo"
         Top             =   720
         Width           =   495
      End
      Begin VB.ComboBox cboCadRequisicao 
         Height          =   315
         Index           =   2
         ItemData        =   "frmRequisicao.frx":0D02
         Left            =   -74880
         List            =   "frmRequisicao.frx":0D12
         TabIndex        =   20
         Text            =   "Colaborador"
         Top             =   720
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   24
         Top             =   1920
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6376
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   -71520
         TabIndex        =   22
         Tag             =   "Nome do responsável pelo setor"
         ToolTipText     =   "Nome do responsável pelo setor"
         Top             =   720
         Width           =   5895
      End
      Begin VB.TextBox txtCadReq 
         Height          =   285
         Index           =   14
         Left            =   -72960
         TabIndex        =   21
         Tag             =   "Código do responsável pelo setor"
         ToolTipText     =   "Código do responsável pelo setor"
         Top             =   720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1560
         TabIndex        =   16
         ToolTipText     =   "Data de previsão para admissão"
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   40560
      End
      Begin VB.TextBox txtCadReq 
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Tag             =   "Nº de vagas solicitadas"
         ToolTipText     =   "Nº de vagas solicitadas"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   8400
         TabIndex        =   14
         ToolTipText     =   "Quantidade de colaboradores que ocupam atualmente o cargo"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtCadReq 
         Height          =   285
         Index           =   12
         Left            =   6600
         TabIndex        =   18
         ToolTipText     =   "Breve observação"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtCadReq 
         Height          =   285
         Index           =   11
         Left            =   3120
         TabIndex        =   17
         ToolTipText     =   "Motivo da admissão"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtCadReq 
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
         HelpContextID   =   1110
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Tag             =   "Nº da matriz do cargo solicitado"
         ToolTipText     =   "Nº da matriz do cargo solicitado"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   11
         Tag             =   "Código do cargo"
         ToolTipText     =   "Código do cargo"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   3000
         TabIndex        =   12
         Tag             =   "Nome do cargo"
         ToolTipText     =   "Nome do cargo"
         Top             =   720
         Width           =   4695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5530
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1920
         TabIndex        =   47
         Tag             =   "Excluir escolaridade"
         ToolTipText     =   "Excluir escolaridade"
         Top             =   1680
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
         MICON           =   "frmRequisicao.frx":0D4C
         PICN            =   "frmRequisicao.frx":0D68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   48
         Tag             =   "Editar escolaridade"
         ToolTipText     =   "Editar escolaridade"
         Top             =   1680
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
         MICON           =   "frmRequisicao.frx":1A42
         PICN            =   "frmRequisicao.frx":1A5E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   49
         Tag             =   "Novo escolaridade"
         ToolTipText     =   "Novo escolaridade"
         Top             =   1680
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
         MICON           =   "frmRequisicao.frx":2738
         PICN            =   "frmRequisicao.frx":2754
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Tag             =   "Incluir escolaridade"
         ToolTipText     =   "Incluir escolaridade"
         Top             =   1680
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
         MICON           =   "frmRequisicao.frx":342E
         PICN            =   "frmRequisicao.frx":344A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   9
         Left            =   -73080
         TabIndex        =   51
         Tag             =   "Excluir histórico"
         ToolTipText     =   "Excluir histórico"
         Top             =   1200
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
         MICON           =   "frmRequisicao.frx":4124
         PICN            =   "frmRequisicao.frx":4140
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   8
         Left            =   -73680
         TabIndex        =   52
         Tag             =   "Editar histórico"
         ToolTipText     =   "Editar histórico"
         Top             =   1200
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
         MICON           =   "frmRequisicao.frx":4E1A
         PICN            =   "frmRequisicao.frx":4E36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   7
         Left            =   -74280
         TabIndex        =   53
         Tag             =   "Novo histórico"
         ToolTipText     =   "Novo histórico"
         Top             =   1200
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
         MICON           =   "frmRequisicao.frx":5B10
         PICN            =   "frmRequisicao.frx":5B2C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   6
         Left            =   -74880
         TabIndex        =   54
         Tag             =   "Incluir histórico"
         ToolTipText     =   "Incluir histórico"
         Top             =   1200
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
         MICON           =   "frmRequisicao.frx":6806
         PICN            =   "frmRequisicao.frx":6822
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label31 
         Caption         =   "Nível:"
         Height          =   255
         Left            =   7800
         TabIndex        =   41
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   -71520
         TabIndex        =   39
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   -72960
         TabIndex        =   38
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Previsão para adm:"
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Nº de vagas:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Qtd. Colaboradores:"
         Height          =   255
         Left            =   8400
         TabIndex        =   35
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "Observação:"
         Height          =   255
         Left            =   6600
         TabIndex        =   34
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   3120
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Matriz nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Código cargo:"
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Nome cargo:"
         Height          =   255
         Left            =   3000
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de recrutamento "
      Height          =   975
      Left            =   3600
      TabIndex        =   27
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   5415
      End
      Begin VB.ComboBox cboCadRequisicao 
         Height          =   315
         Index           =   1
         ItemData        =   "frmRequisicao.frx":74FC
         Left            =   120
         List            =   "frmRequisicao.frx":7506
         TabIndex        =   2
         Text            =   "Interno"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmRequisicao.frx":751C
         TabIndex        =   65
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRequisicao.frx":758A
         TabIndex        =   63
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Requisitante"
      Height          =   1575
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   6495
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2640
         TabIndex        =   8
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtCadReq 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Tag             =   "Identificação do instrutor  do curso/treinamento"
         ToolTipText     =   "Identificação do instrutor do curso/treinamento"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   6
         Tag             =   "Nome do instrutor do curso/treinamento"
         ToolTipText     =   "nome do instrutor do curso/treinamento"
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cboCadRequisicao 
         Height          =   315
         Index           =   0
         ItemData        =   "frmRequisicao.frx":75F6
         Left            =   120
         List            =   "frmRequisicao.frx":7606
         TabIndex        =   4
         Text            =   "Colaborador"
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "frmRequisicao.frx":7640
         TabIndex        =   66
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRequisicao.frx":76AA
         TabIndex        =   64
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "frmRequisicao.frx":7722
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmRequisicao.frx":778A
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRequisicao.frx":77FA
         TabIndex        =   58
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   46
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da requisição"
      Height          =   975
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3375
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   40560
      End
      Begin VB.TextBox txtCadReq 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRequisicao.frx":7862
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmRequisicao.frx":78CE
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
   End
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   720
      TabIndex        =   55
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8640
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
      MICON           =   "frmRequisicao.frx":7936
      PICN            =   "frmRequisicao.frx":7952
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   56
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   8640
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
      MICON           =   "frmRequisicao.frx":862C
      PICN            =   "frmRequisicao.frx":8648
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
Attribute VB_Name = "frmRequisicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsRequisicoes As New ADODB.Recordset
Private sqlRequisicoes As String
Private rsReqCargos As New ADODB.Recordset
Private sqlReqCargos As String

Private rsCargoReq As New ADODB.Recordset
Private sqlCargoReq As String

Private Status As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsReqAprovador As New ADODB.Recordset
Private SqlReqAprovador As String
Private rsLocal As New ADODB.Recordset

Private Sub cboCadTreinamento_LostFocus(Index As Integer)
    Select Case Index
    Case 4
        MontaMascara 0
    End Select
End Sub

Private Sub cboCadRequisicao_Click(Index As Integer)
    txtCadReq(5).Enabled = True
    Select Case Index
    Case 0
        MontaMascara 0
    Case 1
        If cboCadRequisicao(1).Text = "Interno" Then txtCadReq(5).Enabled = False
        If cboCadRequisicao(1).Text = "Externo" Then txtCadReq(5).Enabled = True
    Case 2
        MontaMascara 2
    End Select
End Sub

Private Sub cboCadRequisicao_LostFocus(Index As Integer)
    txtCadReq(5).Enabled = True
    Select Case Index
    Case 0
        MontaMascara 0
    Case 1
        If cboCadRequisicao(1).Text = "Interno" Then txtCadReq(5).Enabled = False
        If cboCadRequisicao(1).Text = "Externo" Then txtCadReq(5).Enabled = True
    End Select
End Sub

Private Sub cmdCad_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridColaborador 0
        CarregaColaborador 0
    Case 1
        ChamaGridCargoReq
        CarregaCargoReq
    Case 2
        ChamaGridColaborador 2
        CarregaColaborador 2
    End Select
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirCargoReq
        LimpaControlesCargoReq
    Case 1
        LimpaControlesCargoReq
    Case 2
        AlteraCargoReq
    Case 3
        mobjMsg.Abrir "Deseja EXCLUIR esse cargo da Requisição?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirCargoReq
            LimpaControlesCargoReq
        End If
    Case 6
        IncluirAprovadorReq
        LimpaControlesAprovadorReq
    Case 7
        LimpaControlesAprovadorReq
    Case 8
        AlteraAprovadorReq
    Case 9
        mobjMsg.Abrir "Deseja EXCLUIR esse aprovador da Requisição?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView2
            LimpaControlesAprovadorReq
        End If
    Case 11
        mobjMsg.Abrir "Deseja salvar os dados de Requisição?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            gravaLog "Código req: " & txtCadReq(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
        End If
    Case 12
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Requisição?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = "0"
            Unload Me
            Set frmRequisicao = Nothing
        End If
    End Select
End Sub

Private Sub IncluirCargoReq()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaCampo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadReq(6) Then
                Me.txtCadReq(6) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadReq(8)
                ListView1.SelectedItem.ListSubItems.Item(2) = txtCadReq(16)
                ListView1.SelectedItem.ListSubItems.Item(3) = txtCadReq(9)
                ListView1.SelectedItem.ListSubItems.Item(4) = txtCadReq(10)
                ListView1.SelectedItem.ListSubItems.Item(5) = DTPicker2
                ListView1.SelectedItem.ListSubItems.Item(6) = txtCadReq(11)
                ListView1.SelectedItem.ListSubItems.Item(7) = txtCadReq(12)
                If Not IsNull(ListView1.SelectedItem.ListSubItems.Item(8)) Then
                    If Val(ListView1.SelectedItem.ListSubItems.Item(8)) < Val(ListView1.SelectedItem.ListSubItems.Item(4)) Then
                        ListView1.SelectedItem.ListSubItems.Item(9) = "Aberto"
                    Else
                        ListView1.SelectedItem.ListSubItems.Item(9) = "Fechado"
                    End If
                Else
                    ListView1.SelectedItem.ListSubItems.Item(9) = "Aberto"
                End If
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadReq(6))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadReq(6))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadReq(8)
    ItemLst.SubItems(2) = txtCadReq(16)
    ItemLst.SubItems(3) = txtCadReq(9)
    ItemLst.SubItems(4) = txtCadReq(10)
    ItemLst.SubItems(5) = DTPicker2
    ItemLst.SubItems(6) = txtCadReq(11)
    ItemLst.SubItems(7) = txtCadReq(12)
    ItemLst.SubItems(8) = 0
    ItemLst.SubItems(9) = "Aberto"
    LimpaControlesCargoReq
    txtCadReq(6).SetFocus
End Sub

Private Sub AlteraCargoReq()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadReq(6).Text = ListView1.ListItems.Item(X)
    Me.txtCadReq(10).Text = ListView1.SelectedItem.ListSubItems.Item(4)
    DTPicker2 = ListView1.SelectedItem.ListSubItems.Item(5)
    Me.txtCadReq(11).Text = ListView1.SelectedItem.ListSubItems.Item(6)
    Me.txtCadReq(12).Text = ListView1.SelectedItem.ListSubItems.Item(7)
    CarregaCargoReq
    Me.txtCadReq(9).Text = ListView1.SelectedItem.ListSubItems.Item(3)
End Sub

Private Sub ExcluirCargoReq()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    Dim llng_Contador As Long
    
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If ListView1.SelectedItem.ListSubItems.Item(9) <> "Fechado" Then
        ListView1.ListItems.Remove (X)
    Else
        mobjMsg.Abrir "Este cargo já está ocupado, não pode ser excluido", Ok, critico, "Atenção"
    End If
End Sub

'--------------------------------
Private Sub IncluirAprovadorReq()
'    If ValidaCampoItem = False Then Exit Sub
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True
            If ListView2.ListItems.Item(X) = Me.Label17.Caption Then
                Label17.Caption = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = cboCadRequisicao(2).Text
                ListView2.SelectedItem.ListSubItems.Item(2) = txtCadReq(14).Text
                ListView2.SelectedItem.ListSubItems.Item(3) = txtCadReq(15).Text
                Y = ListView2.ListItems.Count
                Me.ListView2.Sorted = True
                Me.ListView2.SortKey = 0
                Me.ListView2.SortOrder = lvwAscending
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , Label17)
        Y = ListView2.ListItems.Count
    Else
        Set ItemLst = ListView2.ListItems.Add(, , Label17)
        Y = ListView2.ListItems.Count
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 0
        Me.ListView2.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = cboCadRequisicao(2).Text
    ItemLst.SubItems(2) = txtCadReq(14).Text
    ItemLst.SubItems(3) = txtCadReq(15).Text
    Me.ListView2.SortOrder = lvwAscending
    cboCadRequisicao(2).SetFocus
End Sub

Private Sub AlteraAprovadorReq()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.Label17.Caption = ListView2.ListItems.Item(X)
    Me.cboCadRequisicao(2).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.txtCadReq(14).Text = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.txtCadReq(15).Text = ListView2.SelectedItem.ListSubItems.Item(3)
    MontaMascara 2
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    Legenda = cmdCadastro(Index).ToolTipText
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    listview_cabecalho
    'Status = "novo"
    SSTab1.Tab = 0
    If Status = "novo" Then
        LimpaControles
        Label17.Caption = "000001"
    ElseIf Status = "editar" Then
        ResultPesq
    End If
    configControles
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
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Matriz", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome do cargo", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nível", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "Qtd. Colaboradores", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Nº vagas solicitadas", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Prev. adm", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Motivo", ListView1.Width / 1.5
    ListView1.ColumnHeaders.Add , , "Observação", ListView1.Width / 1.5
    ListView1.ColumnHeaders.Add , , "Vagas ocupadas", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "status", ListView1.Width / 8
    
    ListView2.ColumnHeaders.Add , , "ID", ListView1.Width / 11
    ListView2.ColumnHeaders.Add , , "Tipo", ListView1.Width / 6
    ListView2.ColumnHeaders.Add , , "Registro", ListView1.Width / 8
    ListView2.ColumnHeaders.Add , , "Nome", ListView1.Width / 2
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampoSalvar = False Then Exit Sub
    Dim rsSalvarRequisicao As New ADODB.Recordset
    Dim SqlSalvarRequisicao As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarRequisicao = "select * from tbrequisicoes where codcoligada = '" & vCodcoligada & "' and codrequisicao = '" & txtCadReq(0) & "'"
    rsSalvarRequisicao.Open SqlSalvarRequisicao, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarRequisicao.EOF Then rsSalvarRequisicao.AddNew
    rsSalvarRequisicao.Fields(0) = Val(txtCadReq(0)) 'codigo da requisicao
    rsSalvarRequisicao.Fields(1) = DTPicker1 'Data da requisicao
    rsSalvarRequisicao.Fields(2) = cboCadRequisicao(0) 'Tipo
    rsSalvarRequisicao.Fields(3) = txtCadReq(1) 'codigo do colaborador
    rsSalvarRequisicao.Fields(4) = txtCadReq(2) 'nome do requisitante
    rsSalvarRequisicao.Fields(5) = txtCadReq(3) 'departamento do requisitante
    rsSalvarRequisicao.Fields(6) = txtCadReq(4) 'setor do requisitante
    rsSalvarRequisicao.Fields(7) = cboCadRequisicao(1) 'origem
    rsSalvarRequisicao.Fields(8) = txtCadReq(5) 'nome da empresa
    If Check1.Value = 1 Then rsSalvarRequisicao.Fields(9) = "S" Else rsSalvarRequisicao.Fields(9) = "N" 'ativo
    rsSalvarRequisicao.Fields(10) = txtCadReq(13) 'observação
    rsSalvarRequisicao.Fields(11) = vCodcoligada 'Codigo da coligada
    rsSalvarRequisicao.Update
    
    'SALVAR CARGOS REQUISITADOR - LISTVIEW1
    sqlDeletar = "Delete from tbRequisicoesCargos where tbRequisicoesCargos.codcoligada = '" & vCodcoligada & "' and tbRequisicoesCargos.codrequisicao = '" & Val(txtCadReq(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbRequisicoesCargos where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadReq(0).Text)
        rsSalvar.Fields(1) = ListView1.ListItems.Item(X)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(4)
        rsSalvar.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(5)
        rsSalvar.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(6)
        rsSalvar.Fields(5) = ListView1.SelectedItem.ListSubItems.Item(7)
        rsSalvar.Fields(6) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsSalvar.Fields(7) = ListView1.SelectedItem.ListSubItems.Item(8)
        rsSalvar.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(9)
        rsSalvar.Fields(9) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    'SALVAR APROVADORES DA REQUISIÇÃO - LISTVIEW2
    sqlDeletar = "Delete from tbRequisicoesAprovadores where tbRequisicoesAprovadores.codcoligada = '" & vCodcoligada & "' and tbRequisicoesAprovadores.codrequisicao = '" & Val(txtCadReq(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbRequisicoesAprovadores where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadReq(0).Text)
        rsSalvar.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsSalvar.Fields(4) = ListView2.ListItems.Item(X)
        rsSalvar.Fields(5) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    
    cnBanco.CommitTrans
    rsSalvarRequisicao.Close
    Set rsSalvarRequisicao = Nothing
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    mobjMsg.Abrir "Os dados da Requisição foram salvos com sucesso", Ok, informacao, "Atenção"
    AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    For X = 0 To txtCadReq.Count - 1
        txtCadReq(X) = ""
    Next
    txtCadReq(13) = ""
    cboCadRequisicao(0).Text = "Colaborador"
    cboCadRequisicao(1).Text = "Interno"
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    txtCadReq(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub LimpaControlesCargoReq()
    Dim X As Integer
    For X = 6 To 12
        txtCadReq(X) = ""
    Next
    DTPicker2 = Date
End Sub

Private Sub LimpaControlesAprovadorReq()
    Dim X As Integer
    For X = 14 To 15
        txtCadReq(X) = ""
    Next
    txtCadReq(16) = ""
    cboCadRequisicao(2).Text = "Colaborador"
    MontaMascara 2
    If ListView2.ListItems.Count <= 0 Then
        Label17.Caption = Format(GeraCodigo1, "000000")
    Else
        ListView2.ListItems.Item(ListView2.ListItems.Count).Selected = True
        Label17.Caption = Format(Val(Label17) + 1, "000000")
    End If
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtCadReq(0).Text = Format(rsRequisicoes.Fields(0), "000000") 'código da requisição
    DTPicker1 = rsRequisicoes.Fields(1) 'Data da requisição
    cboCadRequisicao(0).Text = rsRequisicoes.Fields(2) 'Tipo do requisitante (colaborador/contratado/diretoria/superintendencia)
    txtCadReq(1).Text = rsRequisicoes.Fields(3) 'codigo do colaborador (se for colaborador)
    txtCadReq(2).Text = rsRequisicoes.Fields(4) 'nome do requisitante (se não for colaborador)
    txtCadReq(3).Text = rsRequisicoes.Fields(5) 'departamento do requisitantes (se não for colaborador)
    txtCadReq(4).Text = rsRequisicoes.Fields(6) 'setor do requisitante (se não for colaborador)
    cboCadRequisicao(1).Text = rsRequisicoes.Fields(7) 'Origem do recrutamento
    txtCadReq(5).Text = rsRequisicoes.Fields(8) 'nome da empresa q irá realizar o recrutamento (se a origem for externa)
    If rsRequisicoes.Fields(9) = "S" Then Check1.Value = 1 Else Check1.Value = 0  'Informa se a requisição esta ativa ou nao
    txtCadReq(13).Text = rsRequisicoes.Fields(10) 'setor do requisitante (se não for colaborador)
'    MontaMascara
End Sub

Private Sub Compoe_Listview1()
    Dim ItemLst As ListItem
    Dim X As Integer
    sqlReqCargos = "Select a.codmatriz,c.nomecargo,b.nivel,a.qtdcolaboradores,a.numvagas,a.dataprevisaoadm,a.motivo,a.observacao,a.qtdocupada,a.status from tbrequisicoescargos as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = b.codmatriz inner join tbcargos as c on b.codcargo=c.codcargo where a.codrequisicao = '" & Val(txtCadReq(0)) & "'Order by a.codrequisicao"
    rsReqCargos.Open sqlReqCargos, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsReqCargos.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsReqCargos.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsReqCargos.Fields(1)
        ItemLst.SubItems(2) = "" & rsReqCargos.Fields(2)
        ItemLst.SubItems(3) = "" & rsReqCargos.Fields(3)
        ItemLst.SubItems(4) = "" & rsReqCargos.Fields(4)
        ItemLst.SubItems(5) = "" & rsReqCargos.Fields(5)
        ItemLst.SubItems(6) = "" & rsReqCargos.Fields(6)
        ItemLst.SubItems(7) = "" & rsReqCargos.Fields(7)
        If IsNull(rsReqCargos.Fields(8)) Then
            ItemLst.SubItems(8) = 0
        Else
            ItemLst.SubItems(8) = rsReqCargos.Fields(8)
        End If
        ItemLst.SubItems(9) = "" & rsReqCargos.Fields(9)
        rsReqCargos.MoveNext
        X = X + 1
    Wend
    rsReqCargos.Close
    Set rsReqCargos = Nothing
    'Legenda = ""
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
End Sub

Private Sub Compoe_Listview2()
    Dim ItemLst As ListItem
    Dim X As Integer
    SqlReqAprovador = "Select a.sequencia,a.tipo,a.codcolaborador,a.nomeaprovador,b.nomecolaborador from tbrequisicoesaprovadores as a left join tbcolaboradores as b on a.codcolaborador=b.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and a.codrequisicao = '" & Val(txtCadReq(0)) & "'Order by a.sequencia"
    rsReqAprovador.Open SqlReqAprovador, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsReqAprovador.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsReqAprovador.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsReqAprovador.Fields(1)
        ItemLst.SubItems(2) = "" & rsReqAprovador.Fields(2)
        ItemLst.SubItems(3) = "" & rsReqAprovador.Fields(3)
        rsReqAprovador.MoveNext
        X = X + 1
    Wend
    rsReqAprovador.Close
    Set rsReqAprovador = Nothing
    'Legenda = ""
    'Ao preencher todo Listview, ele é ordenado pela coluna zero de forma ascendente
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadReq(6).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadReq(6).Tag, Ok, critico, "Atenção"
        Me.txtCadReq(6).SetFocus
        Exit Function
    End If
    If txtCadReq(10).Text = "" Or txtCadReq(10).Text = 0 Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadReq(10).Tag, Ok, critico, "Atenção"
        Me.txtCadReq(10).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function ValidaCampoSalvar()
    ValidaCampoSalvar = False
    If txtCadReq(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadReq(1).Tag, Ok, critico, "Atenção"
        Me.txtCadReq(1).SetFocus
        Exit Function
    End If
    If ListView1.ListItems.Count = 0 Then
        mobjMsg.Abrir "Favor informar Cargos requisitados", Ok, critico, "Atenção"
        Exit Function
    End If
    If ListView2.ListItems.Count = 0 Then
        mobjMsg.Abrir "Favor informar aprovador da requisição", Ok, critico, "Atenção"
        Exit Function
    End If
    ValidaCampoSalvar = True
End Function

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirRequisicao
    SqlGera = "Select top 1 * from tbRequisicoes where codcoligada = '" & vCodcoligada & "' order by codrequisicao Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsRequisicoes.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadReq(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharRequisicao
End Function

Private Function GeraCodigo1()
'On Error GoTo ERR
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from tbRequisicoesAprovadores where codcoligada = '" & vCodcoligada & "' and codrequisicao = '" & Val(txtCadReq(0)) & "' order by codrequisicao Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    AbrirRequisicaoAprovadores
    If rsReqAprovador.RecordCount > 0 Then
        GeraCodigo1 = rsGeraCodigo.Fields(4) + 1
    Else
        GeraCodigo1 = 1
    End If
    Label17.Caption = Format(GeraCodigo1, "000")
    FecharRequisicaoAprovadores
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    Exit Function
Err:
    Exit Function
End Function

Private Sub AbrirRequisicaoAprovadores()
    SqlReqAprovador = "Select * from tbRequisicoesaprovadores where codcoligada = '" & vCodcoligada & "' Order by codrequisicao"
    rsReqAprovador.Open SqlReqAprovador, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharRequisicaoAprovadores()
    rsReqAprovador.Close
    Set rsReqAprovador = Nothing
End Sub

Private Sub AbrirRequisicao()
    sqlRequisicoes = "Select * from tbRequisicoes where codcoligada = '" & vCodcoligada & "' Order by codrequisicao"
    rsRequisicoes.Open sqlRequisicoes, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharRequisicao()
    rsRequisicoes.Close
    Set rsRequisicoes = Nothing
End Sub

Private Sub ResultPesq()
    sqlRequisicoes = "Select * from tbRequisicoes Where codcoligada = '" & vCodcoligada & "' and tbRequisicoes.codrequisicao= '" & Val(varGlobal) & "' order by tbRequisicoes.codrequisicao"
    rsRequisicoes.Open sqlRequisicoes, cnBanco, adOpenKeyset, adLockReadOnly
    If rsRequisicoes.RecordCount > 0 Then
        CompoeControles
        Compoe_Listview1
        Compoe_Listview2
        If rsRequisicoes.Fields(9) = "N" Then BloqueiaControles
    Else
        mobjMsg.Abrir "Requisição não encontrada", Ok, critico, "Atenção"
    End If
    rsRequisicoes.Close
    Set rsRequisicoes = Nothing
    Label17.Caption = Format(GeraCodigo1, "000")
End Sub

Private Sub BloqueiaControles()
    For X = 0 To 15
        txtCadReq(X).Enabled = False
    Next
    For X = 0 To 11
        cmdCadastro(X).Enabled = False
    Next
    cboCadRequisicao(0).Enabled = False
    cboCadRequisicao(1).Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    Check1.Enabled = False
End Sub

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadReq(0), "000000"))
        ItemLst.SubItems(1) = DTPicker1
        ItemLst.SubItems(2) = cboCadRequisicao(1).Text
        ItemLst.SubItems(3) = txtCadReq(2).Text
        If Check1.Value = 0 Then
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(4) = ""
            ItemLst.ListSubItems.Item(4).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = DTPicker1
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = cboCadRequisicao(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadReq(2).Text
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        If cmdCadastro(2).Enabled = True Then AlteraCargoReq
    End If
End Sub

Private Sub ListView2_DblClick()
    If vEdi <> "N" Then
        If cmdCadastro(8).Enabled = True Then AlteraAprovadorReq
    End If
End Sub

Private Sub txtCadReq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 0
        End If
    Case 6
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCargoReq
        End If
    Case 14
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 2
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaColaborador(indice As Integer)
    Dim X As Integer
    If indice = 0 Then
        SqlColaborador = "select a.codcolaborador,a.nomecolaborador,d.nomedepartamento,e.nomesetor from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor where a.codcolaborador = '" & txtCadReq(1) & "'"
        rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    If indice = 2 Then
        SqlColaborador = "select a.codcolaborador,a.nomecolaborador,d.nomedepartamento,e.nomesetor from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor where a.codcolaborador = '" & txtCadReq(14) & "'"
        rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    
    If indice = 0 Then
        If rsColaborador.RecordCount <= 0 Then
            If txtCadReq(1).Text <> "000000" And txtCadReq(1).Text <> "" Then mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtCadReq(2) = ""
            txtCadReq(3) = ""
            txtCadReq(4) = ""
        Else
            txtCadReq(1).Text = rsColaborador.Fields(0)
            txtCadReq(2).Text = rsColaborador.Fields(1)
            txtCadReq(3).Text = rsColaborador.Fields(2)
            txtCadReq(4).Text = rsColaborador.Fields(3)
        End If
    End If
    If indice = 2 Then
        If rsColaborador.RecordCount <= 0 Then
            If txtCadReq(14).Text <> "000000" And txtCadReq(14).Text <> "" Then mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtCadReq(15) = ""
        Else
            txtCadReq(14).Text = rsColaborador.Fields(0)
            txtCadReq(15).Text = rsColaborador.Fields(1)
        End If
    End If
    rsColaborador.Close
    Set rsColaborador = Nothing
    
End Sub

Private Sub ChamaGridColaborador(indice As Integer)
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' order by nomecolaborador"
    procnom = "nomecolaborador"
    campo = 3
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Colaborador"
    Pesquisa = frmRequisicao.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If indice = 0 Then
                txtCadReq(1).Text = rsLocal.Fields(1)
            End If
            If indice = 2 Then
                txtCadReq(14).Text = rsLocal.Fields(1)
            End If
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaCargoReq()
    Dim X As Integer
    sqlCargoReq = "Select tbMatriz.codmatriz,tbMatriz.codcargo,tbMatriz.nivel,tbcargos.nomecargo from tbMatriz,tbcargos where tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo and tbmatriz.ativo = 'S'  and tbmatriz.codmatriz = '" & Val(txtCadReq(6)) & "' order by tbMatriz.codmatriz"
    rsCargoReq.Open sqlCargoReq, cnBanco, adOpenKeyset, adLockOptimistic
    If rsCargoReq.RecordCount <= 0 Then
        txtCadReq(6).Text = Format(txtCadReq(6), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Matriz não cadastrada", Ok, critico, "Atenção"
            txtCadReq(7) = ""
            txtCadReq(8) = ""
            txtCadReq(16) = ""
        End If
    Else
        txtCadReq(6).Text = Format(rsCargoReq.Fields(0), "000000") & ""
        txtCadReq(7).Text = Format(rsCargoReq.Fields(1), "000000") & ""
        txtCadReq(8).Text = rsCargoReq.Fields(3)
        txtCadReq(16).Text = rsCargoReq.Fields(2)
    End If
    rsCargoReq.Close
    Set rsCargoReq = Nothing

    sqlCargoReq = "Select * from tbColaboradoresHist where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and codmatriz = '" & Val(txtCadReq(6)) & "'"
    rsCargoReq.Open sqlCargoReq, cnBanco, adOpenKeyset, adLockReadOnly
    txtCadReq(9).Text = rsCargoReq.RecordCount
    
    rsCargoReq.Close
    Set rsCargoReq = Nothing
End Sub

Private Sub ChamaGridCargoReq()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select tbmatriz.codmatriz,tbcargos.nomecargo,tbmatriz.nivel,tbdepartamentos.nomedepartamento,tbsetores.nomesetor from tbmatriz,tbdepartamentos,tbsetores,tbcargos where tbmatriz.codcoligada = '" & vCodcoligada & "' and tbmatriz.coddepartamento = tbdepartamentos.coddepartamento and tbmatriz.codsetor = tbsetores.codsetor and tbmatriz.codcargo = tbcargos.codcargo and tbmatriz.ativo = 'S' order by tbcargos.nomecargo,tbMatriz.nivel"
    procnom = "codmatriz"
    procnom1 = "nomecargo"
    campo = 0
    Campo1 = 1
    campo2 = 2
    campo3 = 3
    Campo4 = 4
    Pesquisa = "Histórico"
    Load F
    F.Caption = "Pesquisa de Matrizes"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "codmatriz=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadReq(6).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub MontaMascara(indice As Integer)
    If indice = 0 Then
        If cboCadRequisicao(0) <> "Colaborador" Then
            txtCadReq(1) = Format(0, "000000")
            txtCadReq(1).Enabled = False
            txtCadReq(2).Enabled = True
            txtCadReq(3).Enabled = True
            txtCadReq(4).Enabled = True
            txtCadReq(2).BackColor = &H80000018
            txtCadReq(3).BackColor = &H80000018
            txtCadReq(4).BackColor = &H80000018
            If txtCadReq(2) = "" Then txtCadReq(2).Text = "Digite o nome do requisitante"
            If txtCadReq(3) = "" Then txtCadReq(3).Text = "Digite o departamento do requisitante"
            If txtCadReq(4) = "" Then txtCadReq(4).Text = "Digite o setor do requisitante"
            cmdCadastro(4).Enabled = False
        ElseIf cboCadRequisicao(0) = "Colaborador" Then
            txtCadReq(1).Enabled = True
            txtCadReq(2).Enabled = False
            txtCadReq(3).Enabled = False
            txtCadReq(4).Enabled = False
            txtCadReq(2).BackColor = &H80000005
            txtCadReq(3).BackColor = &H80000005
            txtCadReq(4).BackColor = &H80000005
            txtCadReq(1).Text = ""
            txtCadReq(2).Text = ""
            txtCadReq(3).Text = ""
            txtCadReq(4).Text = ""
            cmdCadastro(4).Enabled = True
            CarregaColaborador 0
        End If
    End If
    If indice = 2 Then
        If cboCadRequisicao(2) <> "Colaborador" Then
            txtCadReq(14) = Format(0, "000000")
            txtCadReq(14).Enabled = False
            txtCadReq(15).Enabled = True
            txtCadReq(15).BackColor = &H80000018
            If txtCadReq(15) = "" Then txtCadReq(15).Text = "Digite o nome do aprovador"
            cmdCadastro(5).Enabled = False
        ElseIf cboCadRequisicao(2) = "Colaborador" Then
            txtCadReq(14).Enabled = True
            txtCadReq(15).Enabled = False
            txtCadReq(15).BackColor = &H80000005
            If txtCadReq(14) <> "000000" And txtCadReq(14) = "" Then
                txtCadReq(14).Text = ""
                txtCadReq(15).Text = ""
            End If
            cmdCadastro(5).Enabled = True
            CarregaColaborador 2
        End If
    End If
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    
        cmdCadastro(6).UseGreyscale = True
        cmdCadastro(6).DragMode = 1
        cmdCadastro(6).SpecialEffect = cbEngraved
        
        cmdCadastro(7).UseGreyscale = True
        cmdCadastro(7).DragMode = 1
        cmdCadastro(7).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    
        cmdCadastro(8).UseGreyscale = True
        cmdCadastro(8).DragMode = 1
        cmdCadastro(8).SpecialEffect = cbEngraved
    
    End If
    If vSal = "N" Then
        cmdCadastro(11).UseGreyscale = True
        cmdCadastro(11).DragMode = 1
        cmdCadastro(11).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    
        cmdCadastro(9).UseGreyscale = True
        cmdCadastro(9).DragMode = 1
        cmdCadastro(9).SpecialEffect = cbEngraved
    End If
End Sub
