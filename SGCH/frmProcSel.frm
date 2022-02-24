VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcSel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processo Seletivo"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   Icon            =   "frmProcSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLvw 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   52
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Parâmetros do Módulo Avaliador"
      Height          =   1695
      Left            =   3960
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   5160
         TabIndex        =   48
         Top             =   240
         Width           =   2175
      End
      Begin VB.Frame Frame10 
         Caption         =   "Média geral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   735
         Left            =   2880
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         Begin VB.Label Label41 
            Caption         =   "Label41"
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Formação escolar:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Cursos/treinamentos:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Habilidades:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkAvaliador 
         Caption         =   "Experiência:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskCadMatriz 
         Height          =   285
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         Height          =   255
         Left            =   2040
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label39 
         Caption         =   "Label39"
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label38 
         Caption         =   "Label38"
         Height          =   255
         Left            =   2040
         TabIndex        =   39
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10560
      TabIndex        =   34
      Top             =   7320
      Width           =   1095
      Begin VB.CheckBox Check3 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "Status do curso/treinamento"
         ToolTipText     =   "Status do curso/treinamento"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Cargos/requisição"
      TabPicture(0)   =   "frmProcSel.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(1)=   "cmdCadastro(1)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Filtro"
      TabPicture(1)   =   "frmProcSel.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ListView2"
      Tab(1).Control(1)=   "cmdCadastro(2)"
      Tab(1).Control(2)=   "Check4"
      Tab(1).Control(3)=   "Frame7"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Selecionados"
      TabPicture(2)   =   "frmProcSel.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCadastro(9)"
      Tab(2).Control(1)=   "ListView3"
      Tab(2).Control(2)=   "cmdCadastro(3)"
      Tab(2).Control(3)=   "Check5"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Admitidos"
      TabPicture(3)   =   "frmProcSel.frx":0D1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "ListView4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdCadastro(4)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdCadastro(6)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdCadastro(5)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   5
         Left            =   840
         TabIndex        =   54
         Tag             =   "Adicionar dados complementares"
         ToolTipText     =   "Adicionar dados complementares"
         Top             =   720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":0D3A
         PICN            =   "frmProcSel.frx":0D56
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Top             =   720
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":1A30
         PICN            =   "frmProcSel.frx":1A4C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   4
         Left            =   10680
         TabIndex        =   53
         Tag             =   "Concluir Processo Seletivo"
         ToolTipText     =   "Concluir Processo Seletivo"
         Top             =   720
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":2726
         PICN            =   "frmProcSel.frx":2742
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame7 
         Caption         =   "Pesquisa"
         Height          =   735
         Left            =   -73560
         TabIndex        =   51
         Top             =   600
         Width           =   6135
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmProcSel.frx":341C
            Left            =   120
            List            =   "frmProcSel.frx":341E
            TabIndex        =   14
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.CheckBox Check5 
         Height          =   255
         Left            =   -74760
         TabIndex        =   50
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   1110
         Width           =   195
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   3255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
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
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   -73920
         TabIndex        =   18
         Tag             =   "Admitir candidato selecionado"
         ToolTipText     =   "Admitir candidato selecionado"
         Top             =   720
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
         BCOL            =   10862530
         BCOLO           =   10862530
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":3420
         PICN            =   "frmProcSel.frx":343C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   -74520
         TabIndex        =   13
         Tag             =   "Incluir no Processo Seletivo"
         ToolTipText     =   "Incluir no Processo Seletivo"
         Top             =   720
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
         BCOL            =   10862530
         BCOLO           =   10862530
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":4116
         PICN            =   "frmProcSel.frx":4132
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   16
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   -74760
         TabIndex        =   11
         Top             =   720
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
         BCOL            =   10862530
         BCOLO           =   10862530
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":4E0C
         PICN            =   "frmProcSel.frx":4E28
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
         Height          =   3255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   9
         Left            =   -74520
         TabIndex        =   17
         Tag             =   "Excluir do Proceso Seletivo"
         ToolTipText     =   "Excluir do Proceso Seletivo"
         Top             =   720
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
         MICON           =   "frmProcSel.frx":5B02
         PICN            =   "frmProcSel.frx":5B1E
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
   Begin VB.Frame Frame4 
      Caption         =   "Dados da Requisição "
      Height          =   975
      Left            =   120
      TabIndex        =   30
      Top             =   1200
      Width           =   8535
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   6960
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   156434433
         CurrentDate     =   40686
      End
      Begin VB.TextBox txtProcesso 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtProcesso 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Tag             =   "Código da requisição"
         ToolTipText     =   "Código da requisição"
         Top             =   480
         Width           =   1455
      End
      Begin SGCH.chameleonButton cmdCadastro 
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
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
         BCOL            =   10862530
         BCOLO           =   10862530
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcSel.frx":67F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "Data da requisição:"
         Height          =   255
         Left            =   6960
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Requisitante:"
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Exibir linhas/marcação"
      Height          =   975
      Left            =   6720
      TabIndex        =   29
      Top             =   120
      Width           =   1935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmProcSel.frx":6814
         Left            =   600
         List            =   "frmProcSel.frx":6824
         TabIndex        =   5
         Text            =   "2"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listar "
      Height          =   975
      Left            =   4440
      TabIndex        =   28
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox Check2 
         Caption         =   "Candidatos"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Colaboradores"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do Processo Seletivo"
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Format          =   156434433
         CurrentDate     =   40686
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   156434433
         CurrentDate     =   40686
      End
      Begin VB.TextBox txtProcesso 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Data término:"
         Height          =   255
         Left            =   3000
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data início:"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   720
      TabIndex        =   23
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7320
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
      MICON           =   "frmProcSel.frx":6836
      PICN            =   "frmProcSel.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7320
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
      MICON           =   "frmProcSel.frx":752C
      PICN            =   "frmProcSel.frx":7548
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Caption         =   "Foi atingido a data de término do Processo Seletivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   56
      Top             =   7560
      Visible         =   0   'False
      Width           =   6975
   End
End
Attribute VB_Name = "frmProcSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Abaixo - usado poder editar o listview --------------------
'straight from the standard API Viewver
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Declare Function GetScrollInfo Lib "user32" (ByVal HWnd As Long, ByVal fnBar As Long, lpScrollInfo As SCROLLINFO) As Long
 
'interestingly, API Viewer doesn't have these constants, translating from Windows.h is straight forward
Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
  
'my declarations
Private Const c_EntryTxt = ""
Private m_ColIndex As Long 'listview col index
Private m_RowIndex As Long 'listview row index
'-------------------------------------------------------------

Private rsProcesso As New ADODB.Recordset
Private sqlProcesso As String

Private rsRequisicoes As New ADODB.Recordset
Private sqlRequisicoes As String

Private rsProcCargos As New ADODB.Recordset
Private sqlProcCargos As String

Private rsCandidatos As New ADODB.Recordset
Private sqlCandidatos As String

Private rscriaTabTemp As New ADODB.Recordset
Private SqlcriaTabTemp

Private rsSalvarNovoCol As New ADODB.Recordset
Private SqlSalvarNovoCol As String
Private rsDeletar As New ADODB.Recordset
Private sqlDeletar As String

Private rsLocal As New ADODB.Recordset

Private Sub Check4_Click()
    MarcaDesmarca ListView2
End Sub

Private Sub Check5_Click()
    MarcaDesmarca ListView3
End Sub

Private Sub MarcaDesmarca(LV As ListView)
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems(X).Selected = True
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        Else
            LV.ListItems.Item(X).Checked = True
        End If
    Next
End Sub

Private Sub Desmarca(LV As ListView)
    Dim Y As Integer, X As Integer
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems.Item(X).Checked = False
    Next
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridProc
        CarregaProc
        Compoe_Listview1
        ListView2.ListItems.Clear
        ListView3.ListItems.Clear
        SSTab1.Tab = 0
    Case 1
        ListView2.ListItems.Clear
        filtrarCandidatos
        SSTab1.Tab = 1
        Combo2.Clear
        CompoeComboLV2 Combo2
        MudaCorLV ListView2
    Case 2
        incluiPC
        Desmarca ListView2
        SSTab1.Tab = 2
        MudaCorLV ListView3
    Case 3
        admCand
        Desmarca ListView3
        SSTab1.Tab = 3
        MudaCorLV ListView4
    Case 4
        'RESERVAR VAGAS PARA O PROCESSO SELETIVO
        If MsgBox("Deseja iniciar a conclusão do Processo Seletivo?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravaPS
            If reservaVagaPS = False Then Exit Sub
            'gravaLog "Código PS: " & txtProcesso(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
            carregaADP "TODOS"
            Unload Me
        End If
        Unload Me
        Set frmProcSel = Nothing
    Case 5
        AchaItem
        Set chamaForm = New frmProcSelAddDados
        chamaForm.Show 1
        If Sqlp = True Then
            ListView4.SelectedItem.ListSubItems.Item(8) = AddDadosGeral(8)
            ListView4.SelectedItem.ListSubItems.Item(10) = AddDadosGeral(9)
            If vIntegra = "S" Then
                ListView4.SelectedItem.ListSubItems.Item(11) = AddDadosGeral(7)
            End If
            MsgBox "Dados atualizados com sucesso!"
        End If
    Case 6
        ExcluirItemLV ListView4
    Case 9
        ExcluirLV3
    Case 11
        If MsgBox("Deseja salvar os dados do Processo Seletivo?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravaPS
        End If
    Case 12
        If MsgBox("Deseja sair da tela de cadastro de Processo Seletivo?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = "0"
            Unload Me
        End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    listview_cabecalho
    SSTab1.Tab = 0
    DTPicker1 = Date
    DTPicker2 = Date
    If Status = "novo" Then
        LimpaControles
'        Label17.Caption = "000001"
    ElseIf Status = "editar" Then
        ResultPesq
    End If
'    configControles
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Matriz", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Nome do cargo", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nível", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Qtd. requi.", ListView1.Width / 10.5
    ListView1.ColumnHeaders.Add , , "Qtd. aprov.", ListView1.Width / 10.5
    ListView1.ColumnHeaders.Add , , "Prev. adm", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Motivo", ListView1.Width / 1.5
    
    ListView2.ColumnHeaders.Add , , "CPF", ListView2.Width / 6.5
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 5
    ListView2.ColumnHeaders.Add , , "Matriz encontrada", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "Cargo", ListView2.Width / 5
    ListView2.ColumnHeaders.Add , , "Tipo", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Cargo Requisitado", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Nota", ListView2.Width / 10
    ListView2.ColumnHeaders.Add , , "Matriz Requisitada", ListView2.Width / 10000
    
    ListView3.ColumnHeaders.Add , , "CPF", ListView3.Width / 6.5
    ListView3.ColumnHeaders.Add , , "Nome", ListView3.Width / 5
    ListView3.ColumnHeaders.Add , , "Matriz encontrada", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Cargo", ListView3.Width / 5
    ListView3.ColumnHeaders.Add , , "Tipo", ListView3.Width / 8
    ListView3.ColumnHeaders.Add , , "Cargo Requisitado", ListView3.Width / 6
    ListView3.ColumnHeaders.Add , , "Nota", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Ordem", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Matriz Requisitada", ListView3.Width / 10000
    
    ListView4.ColumnHeaders.Add , , "CPF", ListView4.Width / 9.5
    ListView4.ColumnHeaders.Add , , "Nome", ListView4.Width / 5
    ListView4.ColumnHeaders.Add , , "Matriz encontrada", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Cargo", ListView4.Width / 5
    ListView4.ColumnHeaders.Add , , "Tipo", ListView4.Width / 8
    ListView4.ColumnHeaders.Add , , "Cargo Requisitado", ListView4.Width / 6
    ListView4.ColumnHeaders.Add , , "Nota", ListView4.Width / 11
    ListView4.ColumnHeaders.Add , , "Ordem", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Observação", ListView4.Width / 10
    ListView4.ColumnHeaders.Add , , "Matriz Requisitada", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Registro nº", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Cód Func Totvs", ListView4.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    For X = 0 To txtProcesso.Count - 1
        txtProcesso(X) = ""
    Next
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    ListView3.ListItems.Clear
    ListView4.ListItems.Clear
    txtProcesso(0) = Format(GeraCodigo, "000000")
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirProcesso
    SqlGera = "Select top 1 * from tbProcessos where codcoligada = '" & vCodcoligada & "' order by codprocesso Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProcesso.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtProcesso(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharProcesso
End Function

Private Sub AbrirProcesso()
    sqlProcesso = "Select * from tbProcessos where codcoligada = '" & vCodcoligada & "' Order by codProcesso"
    rsProcesso.Open sqlProcesso, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharProcesso()
    rsProcesso.Close
    Set rsProcesso = Nothing
End Sub

Private Sub ChamaGridProc()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select a.codrequisicao,b.codmatriz,d.nomecargo,b.numvagas,b.qtdocupada from tbrequisicoes as a inner join tbRequisicoesCargos as b on a.codcoligada = '" & vCodcoligada & "' and b.codrequisicao = a.codrequisicao inner join tbmatriz as c on c.codmatriz = b.codmatriz inner join tbcargos as d on d.codcargo = c.codcargo where b.status = 'Aberto' order by a.codrequisicao"
    procnom = "codrequisicao"
    procnom1 = "codcargo"
    campo = 0
    Campo1 = 1
    campo2 = 2
    Pesquisa = "Admissao"
    Load F
    F.Caption = "Pesquisa de Requisições"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        If Pesquisa <> 0 Then
            txtProcesso(1).Text = Mid$(Pesquisa, 1, 6)
        End If
        txtProcesso(1).SetFocus
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub ListView4_DblClick()
    If cmdCadastro(5).Enabled = True Then
        AchaItem
        frmProcSelAddDados.Show 1
        ListView4.SelectedItem.ListSubItems.Item(8) = AddDadosGeral(8)
        ListView4.SelectedItem.ListSubItems.Item(10) = AddDadosGeral(9)
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ' Ao teclar ENTER no TexBox Text1 chama a Sub Pesquisar
        Pesquisar
    End If
End Sub

Private Sub Pesquisar()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count 'Conta as linhas preenchidas do Listview
    For X = 1 To Y
        ListView2.ListItems(X).Selected = True 'Seleciona a linha de acordo com o valor de "X"
        If UCase(ListView2.SelectedItem.ListSubItems.Item(1)) Like UCase(Me.Text2.Text & "*") And ListView2.SelectedItem.ListSubItems.Item(5) = Combo2.Text Then
            ListView2.ListItems(X).Selected = True
            ListView2.ListItems(X).EnsureVisible
            Exit Sub
        End If
    Next
End Sub

Private Sub txtProcesso_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If CarregaProc = False Then Exit Sub
            Compoe_Listview1
            ListView2.ListItems.Clear
            ListView3.ListItems.Clear
            ListView4.ListItems.Clear
            SSTab1.Tab = 0
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Function CarregaProc()
    CarregaProc = False
    Dim rsProc As New ADODB.Recordset 'OK
    Dim sqlProc As String 'OK
    Dim X As Integer
    sqlProc = "Select a.codrequisicao,a.nomerequisitante,b.status,a.datarequisicao,c.codprocesso from tbrequisicoes as a inner join tbrequisicoescargos as b on a.codcoligada = '" & vCodcoligada & "' and a.codrequisicao = b.codrequisicao left join tbProcessos as c on a.codrequisicao = c.codrequisicao where b.status = 'Aberto' and a.codrequisicao = '" & Val(txtProcesso(1)) & "' order by a.codrequisicao"
    rsProc.Open sqlProc, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProc.RecordCount <= 0 Then
        MsgBox "Requisição não encontrada", vbInformation, "SGCH"
        LimpaControles
    ElseIf Not IsNull(rsProc.Fields(4)) Then
        MsgBox "Requisição esta reservada para outro processo seletivo", vbInformation, "SGCH"
        LimpaControles
    Else
        txtProcesso(1).Text = Format(rsProc.Fields(0), "000000") & ""
        txtProcesso(2).Text = rsProc.Fields(1)
        DTPicker3 = rsProc.Fields(3)
        CarregaProc = True
    End If
    rsProc.Close
    Set rsProc = Nothing
End Function

Private Sub Compoe_Listview1()
    Dim ItemLst As ListItem
    Dim X As Integer
    'sqlProcCargos = "Select b.codmatriz,d.nomecargo,c.nivel,b.numvagas-b.qtdocupada,b.dataprevisaoadm,b.motivo from tbrequisicoes as a inner join tbrequisicoescargos as b on a.codrequisicao = b.codrequisicao inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo where a.codrequisicao = '" & Val(txtProcesso(1)) & "'and b.status = 'Aberto' Order by a.codrequisicao"
    sqlProcCargos = "Select b.codmatriz,d.nomecargo,c.nivel,b.numvagas-b.qtdocupada,b.dataprevisaoadm,b.motivo from tbrequisicoes as a inner join tbrequisicoescargos as b on a.codcoligada = '" & vCodcoligada & "' and a.codrequisicao = b.codrequisicao inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo where a.codrequisicao = '" & Val(txtProcesso(1)) & "' Order by a.codrequisicao"
    rsProcCargos.Open sqlProcCargos, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    ListView1.ListItems.Clear
    While Not rsProcCargos.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsProcCargos.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsProcCargos.Fields(1)
        ItemLst.SubItems(2) = "" & rsProcCargos.Fields(2)
        ItemLst.SubItems(3) = "" & rsProcCargos.Fields(3)
        ItemLst.SubItems(4) = "0"
        ItemLst.SubItems(5) = "" & rsProcCargos.Fields(4)
        ItemLst.SubItems(6) = "" & rsProcCargos.Fields(5)
        rsProcCargos.MoveNext
        X = X + 1
    Wend
    rsProcCargos.Close
    Set rsProcCargos = Nothing
    
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
End Sub

Private Sub filtrarCandidatos()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Checked = True Then
            ListView1.ListItems.Item(X).Selected = True
            txtCadMatriz(4) = ListView1.ListItems.Item(X)
            Text1.Text = ListView1.ListItems.Item(X) & ListView1.SelectedItem.ListSubItems.Item(1)
        
            SqlcriaTabTemp = "Delete from tbProcessoListaTmp where codcoligada = '" & vCodcoligada & "'"
            rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
            selecionaCandidatos 1
            Compoe_Listview2
        
            SqlcriaTabTemp = "Delete from tbProcessoListaTmp where codcoligada = '" & vCodcoligada & "'"
            rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
            selecionaCandidatos 2
            Compoe_Listview2
        End If
    Next
End Sub

Private Sub selecionaCandidatos(filtra As Integer)
    Dim R As Integer
    If filtra = 1 Then
        If Check1.Value = 0 Then Exit Sub
        sqlCandidatos = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo from tbColaboradores as a inner join tbColaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join  tbcargos as d on c.codcargo = d.codcargo where a.tipo = 'colaborador' and b.ativo = 'S' and b.codmatriz <> '" & Val(txtCadMatriz(4)) & "'"
    Else
        If Check2.Value = 0 Then Exit Sub
        sqlCandidatos = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo from tbColaboradores as a inner join tbColaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join  tbcargos as d on c.codcargo = d.codcargo where a.tipo = 'candidato' and b.ativo = 'S'"
    End If
    rsCandidatos.Open sqlCandidatos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCandidatos.RecordCount <> 0 Then frmMenu2.ProgressBar1.Max = rsCandidatos.RecordCount
    R = 0
    Legenda = "Aguarde, Filtrando Colaboradores..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    If Not rsCandidatos.EOF Then
        While Not rsCandidatos.EOF '.Move(Val(Combo1.Text))
            frmMenu2.ProgressBar1.Value = R
            mskCadMatriz = rsCandidatos.Fields(0)
            If filtra = 1 Then
                Avaliador "colaborador"
                criaTabTemp "colaborador"
            Else
                Avaliador "candidato"
                criaTabTemp "candidato"
            End If
            rsCandidatos.MoveNext
            R = R + 1
        Wend
    End If
    frmMenu2.ProgressBar1.Value = 0
    frmMenu2.StatusBar1.Panels(3).Text = "Grupo: " & GrupoUsu
    rsCandidatos.Close
    Set rsCandidatos = Nothing
End Sub

Private Sub criaTabTemp(ColCan As String)
    On Error Resume Next
    SqlcriaTabTemp = "Insert into tbProcessoListaTmp(cpf,nome,matrizcpf,cargocpf,tipo,cargopesq,nota,matrizpesq,codcoligada) Values('" & rsCandidatos.Fields(0) & "','" & rsCandidatos.Fields(1) & "','" & Str(rsCandidatos.Fields(2)) & "','" & rsCandidatos.Fields(3) & "','" & ColCan & "','" & Mid$(Text1.Text, 7, 50) & "','" & RemoveMask(Label41.Caption) & "','" & Mid$(Text1.Text, 1, 6) & "','" & vCodcoligada & "')"
    rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
End Sub

Private Sub Compoe_Listview2()
    Dim ItemLst As ListItem
    Dim X As Integer
    SqlcriaTabTemp = "select top " & Val(Combo1.Text) & " * from tbProcessoListaTmp where codcoligada = '" & vCodcoligada & "' Order by tipo,cast(replace(nota,',','.') as float) desc"
    rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
    While Not rscriaTabTemp.EOF
        Set ItemLst = ListView2.ListItems.Add(, , rscriaTabTemp.Fields(0)) 'CPF
        ItemLst.SubItems(1) = "" & rscriaTabTemp.Fields(1) 'Nome
        ItemLst.SubItems(2) = "" & Format(rscriaTabTemp.Fields(2), "000000")
        ItemLst.SubItems(3) = "" & rscriaTabTemp.Fields(3)
        ItemLst.SubItems(4) = "" & rscriaTabTemp.Fields(4)
        ItemLst.SubItems(5) = "" & rscriaTabTemp.Fields(5)
        ItemLst.SubItems(6) = "" & rscriaTabTemp.Fields(6) & "%"
        ItemLst.SubItems(7) = "" & rscriaTabTemp.Fields(7)
        rscriaTabTemp.MoveNext
    Wend
    rscriaTabTemp.Close
    Set rscriaTabTemp = Nothing
End Sub

Private Sub incluiPC() 'Incluir Filtrado no Processo Seletivo
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Selected = True
            Dim ItemLst As ListItem
            Dim K As Integer, L As Integer, LV3Edit As String
            L = ListView3.ListItems.Count
            LV3Edit = ""
            If L > 0 Then
                For K = 1 To L
                    ListView3.ListItems.Item(K).Selected = True
                    If ListView3.ListItems.Item(K) = ListView2.ListItems.Item(X) Then
                        'A String LV3Edit recebe qualquer valor para informar
                        'se entrou na rotina de edição ou não
                        LV3Edit = ListView2.ListItems.Item(X)
                        ListView3.SelectedItem.ListSubItems.Item(1) = ListView2.SelectedItem.ListSubItems.Item(1)
                        ListView3.SelectedItem.ListSubItems.Item(2) = ListView2.SelectedItem.ListSubItems.Item(2)
                        ListView3.SelectedItem.ListSubItems.Item(3) = ListView2.SelectedItem.ListSubItems.Item(3)
                        ListView3.SelectedItem.ListSubItems.Item(4) = ListView2.SelectedItem.ListSubItems.Item(4)
                        ListView3.SelectedItem.ListSubItems.Item(5) = ListView2.SelectedItem.ListSubItems.Item(5)
                        ListView3.SelectedItem.ListSubItems.Item(6) = ListView2.SelectedItem.ListSubItems.Item(6)
                        ListView3.SelectedItem.ListSubItems.Item(7) = ListView2.SelectedItem.ListSubItems.Item(5) & ListView2.SelectedItem.ListSubItems.Item(6)
                        ListView3.SelectedItem.ListSubItems.Item(8) = ListView2.SelectedItem.ListSubItems.Item(7)
                    End If
                Next
                If LV3Edit = "" Then
                    Set ItemLst = ListView3.ListItems.Add(, , ListView2.ListItems.Item(X))
                    ItemLst.SubItems(1) = ListView2.SelectedItem.ListSubItems.Item(1)
                    ItemLst.SubItems(2) = ListView2.SelectedItem.ListSubItems.Item(2)
                    ItemLst.SubItems(3) = ListView2.SelectedItem.ListSubItems.Item(3)
                    ItemLst.SubItems(4) = ListView2.SelectedItem.ListSubItems.Item(4)
                    ItemLst.SubItems(5) = ListView2.SelectedItem.ListSubItems.Item(5)
                    ItemLst.SubItems(6) = ListView2.SelectedItem.ListSubItems.Item(6)
                    ItemLst.SubItems(7) = ListView2.SelectedItem.ListSubItems.Item(5) & ListView2.SelectedItem.ListSubItems.Item(6)
                    ItemLst.SubItems(8) = ListView2.SelectedItem.ListSubItems.Item(7)
                End If
            Else
                Set ItemLst = ListView3.ListItems.Add(, , ListView2.ListItems.Item(X))
                ItemLst.SubItems(1) = ListView2.SelectedItem.ListSubItems.Item(1)
                ItemLst.SubItems(2) = ListView2.SelectedItem.ListSubItems.Item(2)
                ItemLst.SubItems(3) = ListView2.SelectedItem.ListSubItems.Item(3)
                ItemLst.SubItems(4) = ListView2.SelectedItem.ListSubItems.Item(4)
                ItemLst.SubItems(5) = ListView2.SelectedItem.ListSubItems.Item(5)
                ItemLst.SubItems(6) = ListView2.SelectedItem.ListSubItems.Item(6)
                ItemLst.SubItems(7) = ListView2.SelectedItem.ListSubItems.Item(5) & ListView2.SelectedItem.ListSubItems.Item(6)
                ItemLst.SubItems(8) = ListView2.SelectedItem.ListSubItems.Item(7)
            End If
        End If
    Next
    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 7
    Me.ListView3.SortOrder = lvwDescending
End Sub

Private Sub admCand() 'Incluir Filtrado no Processo Seletivo
    Dim Y As Integer, X As Integer, vID As Integer
    Y = ListView3.ListItems.Count
    
    For X = 1 To Y
        If ListView3.ListItems.Item(X).Checked = True Then
            ListView3.ListItems.Item(X).Selected = True
            Dim ItemLst As ListItem
            Dim K As Integer, L As Integer, LV3Edit As String
            L = ListView4.ListItems.Count
            LV3Edit = ""
            
            SqlcriaTabTemp = "Select a.id from tbcolaboradores as a where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & ListView3.ListItems.Item(X) & "'"
            rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
            vID = rscriaTabTemp.Fields(0)
            rscriaTabTemp.Close
            Set rscriaTabTemp = Nothing
            
            If L > 0 Then
                For K = 1 To L
                    ListView4.ListItems.Item(K).Selected = True
                    If ListView4.ListItems.Item(K) = ListView3.ListItems.Item(X) Then
                        'A String LV3Edit recebe qualquer valor para informar
                        'se entrou na rotina de edição ou não
                        LV3Edit = ListView3.ListItems.Item(X)
                        ListView4.SelectedItem.ListSubItems.Item(1) = ListView3.SelectedItem.ListSubItems.Item(1)
                        ListView4.SelectedItem.ListSubItems.Item(2) = ListView3.SelectedItem.ListSubItems.Item(2)
                        ListView4.SelectedItem.ListSubItems.Item(3) = ListView3.SelectedItem.ListSubItems.Item(3)
                        ListView4.SelectedItem.ListSubItems.Item(4) = ListView3.SelectedItem.ListSubItems.Item(4)
                        ListView4.SelectedItem.ListSubItems.Item(5) = ListView3.SelectedItem.ListSubItems.Item(5)
                        ListView4.SelectedItem.ListSubItems.Item(6) = ListView3.SelectedItem.ListSubItems.Item(6)
                        ListView4.SelectedItem.ListSubItems.Item(7) = ListView3.SelectedItem.ListSubItems.Item(5) & ListView3.SelectedItem.ListSubItems.Item(6)
                        ListView4.SelectedItem.ListSubItems.Item(9) = ListView3.SelectedItem.ListSubItems.Item(8)
                        ListView4.SelectedItem.ListSubItems.Item(10) = "-"
                        ListView4.SelectedItem.ListSubItems.Item(11) = comporControlesTotvs(vID)
                        'ListView4.SelectedItem.ListSubItems.Item(11) = "-"
                    End If
                Next
                If LV3Edit = "" Then
                    Set ItemLst = ListView4.ListItems.Add(, , ListView3.ListItems.Item(X))
                    ItemLst.SubItems(1) = ListView3.SelectedItem.ListSubItems.Item(1)
                    ItemLst.SubItems(2) = ListView3.SelectedItem.ListSubItems.Item(2)
                    ItemLst.SubItems(3) = ListView3.SelectedItem.ListSubItems.Item(3)
                    ItemLst.SubItems(4) = ListView3.SelectedItem.ListSubItems.Item(4)
                    ItemLst.SubItems(5) = ListView3.SelectedItem.ListSubItems.Item(5)
                    ItemLst.SubItems(6) = ListView3.SelectedItem.ListSubItems.Item(6)
                    ItemLst.SubItems(7) = ListView3.SelectedItem.ListSubItems.Item(5) & ListView3.SelectedItem.ListSubItems.Item(6)
                    ItemLst.SubItems(9) = ListView3.SelectedItem.ListSubItems.Item(8)
                    ItemLst.SubItems(10) = "-"
                    ItemLst.SubItems(11) = comporControlesTotvs(vID)
                
                End If
            Else
                Set ItemLst = ListView4.ListItems.Add(, , ListView3.ListItems.Item(X))
                ItemLst.SubItems(1) = ListView3.SelectedItem.ListSubItems.Item(1)
                ItemLst.SubItems(2) = ListView3.SelectedItem.ListSubItems.Item(2)
                ItemLst.SubItems(3) = ListView3.SelectedItem.ListSubItems.Item(3)
                ItemLst.SubItems(4) = ListView3.SelectedItem.ListSubItems.Item(4)
                ItemLst.SubItems(5) = ListView3.SelectedItem.ListSubItems.Item(5)
                ItemLst.SubItems(6) = ListView3.SelectedItem.ListSubItems.Item(6)
                ItemLst.SubItems(7) = ListView3.SelectedItem.ListSubItems.Item(5) & ListView3.SelectedItem.ListSubItems.Item(6)
                ItemLst.SubItems(9) = ListView3.SelectedItem.ListSubItems.Item(8)
                ItemLst.SubItems(10) = "-"
                ItemLst.SubItems(11) = comporControlesTotvs(vID)
            End If
        End If
    Next
    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 7
    Me.ListView4.SortOrder = lvwDescending
End Sub

Private Sub ExcluirLV3()
On Error GoTo Err
    Dim X As Integer, Y As Integer
    Y = ListView3.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView3.ListItems.Item(X).Checked = True Then
            ListView3.ListItems.Remove (X)
            Y = ListView3.ListItems.Count
            If Y = 0 Then Exit For
            X = 0
        End If
    Next
Err:
    Exit Sub
End Sub

Public Sub CompoeComboLV2(Combo As ComboBox, Optional Column As ColumnHeader = Nothing)
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Checked = True Then
            ListView1.ListItems.Item(X).Selected = True
            Combo.AddItem ListView1.SelectedItem.ListSubItems.Item(1)
        End If
    Next
End Sub

Private Sub MudaCorLV(LV As ListView)
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        'Verde - Aprovado
        If Val(Mid$(LV.SelectedItem.ListSubItems.Item(6), 1, 5)) > MediaGlobal Then
            LV.SelectedItem.ListSubItems.Item(6).ForeColor = &H8000&
        'Laranja - Aprovador com restrição
        ElseIf Val(Mid$(LV.SelectedItem.ListSubItems.Item(6), 1, 5)) >= vAprovadoRest And Val(Mid$(LV.SelectedItem.ListSubItems.Item(6), 1, 5)) < MediaGlobal Then
            LV.SelectedItem.ListSubItems.Item(6).ForeColor = &H80FF&
        'Vermelho - Reprovado
        ElseIf Val(Mid$(LV.SelectedItem.ListSubItems.Item(6), 1, 5)) < vAprovadoRest Then
            LV.SelectedItem.ListSubItems.Item(6).ForeColor = &HC0&
        End If
        LV.SelectedItem.ListSubItems.Item(6).Bold = True
    Next
End Sub

Private Sub GravaPS()
On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarProcSel As New ADODB.Recordset
    Dim SqlSalvarProcSel As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarProcSel = "select * from tbProcessos where codcoligada = '" & vCodcoligada & "' and codprocesso = '" & txtProcesso(0) & "'"
    rsSalvarProcSel.Open SqlSalvarProcSel, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarProcSel.EOF Then rsSalvarProcSel.AddNew
    rsSalvarProcSel.Fields(0) = Val(txtProcesso(0)) 'codigo do Processo Seletivo
    rsSalvarProcSel.Fields(1) = Val(txtProcesso(1)) 'codigo da Requisição do Processo Seletivo
    rsSalvarProcSel.Fields(2) = DTPicker1 'Data de início do Processo Seletivo
    rsSalvarProcSel.Fields(3) = DTPicker2 'Data de Término do Processo Seletivo
    If Check1.Value = 0 And Check2.Value = 0 Then rsSalvarProcSel.Fields(4) = 0 'Listar nada
    If Check1.Value = 1 And Check2.Value = 0 Then rsSalvarProcSel.Fields(4) = 1 'Listar colaboradores
    If Check1.Value = 0 And Check2.Value = 1 Then rsSalvarProcSel.Fields(4) = 2 'Listar candidatos
    If Check1.Value = 1 And Check2.Value = 1 Then rsSalvarProcSel.Fields(4) = 3 'Listar colaboradores/candidatos
    rsSalvarProcSel.Fields(5) = Combo1 'linhas
    rsSalvarProcSel.Fields(6) = "Aberto" 'Status
    If Check3.Value = 1 Then rsSalvarProcSel.Fields(7) = "S" Else rsSalvarProcSel.Fields(7) = "N" 'ativo
    rsSalvarProcSel.Fields(8) = vCodcoligada ' Codigo da Coligada
    rsSalvarProcSel.Update
    
    'SALVAR CARGOS PROCESSO SELETIVO - LISTVIEW1
    sqlDeletar = "Delete from tbProcessosCargos where tbProcessosCargos.codcoligada = '" & vCodcoligada & "' and tbProcessosCargos.codprocesso = '" & Val(txtProcesso(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbProcessosCargos where codcoligada ='" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtProcesso(0).Text)
        rsSalvar.Fields(1) = ListView1.ListItems.Item(X)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(4)
        rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    'SALVAR SELECIONADOS DO PROCESSO SELETIVO - LISTVIEW3
    sqlDeletar = "Delete from tbProcessosParticipantes where tbProcessosParticipantes.codcoligada = '" & vCodcoligada & "' and tbProcessosParticipantes.codprocesso = '" & Val(txtProcesso(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbProcessosParticipantes where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView3.ListItems.Count
        ListView3.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtProcesso(0).Text) ' Código do Processo Seletivo
        rsSalvar.Fields(1) = ListView3.ListItems.Item(X) 'CPF Sugerido no Processo Seletivo
        rsSalvar.Fields(2) = ListView3.SelectedItem.ListSubItems.Item(8) 'Matriz Pesquisada no Processo Seletivo
        rsSalvar.Fields(3) = ListView3.SelectedItem.ListSubItems.Item(4) ' Tipo
        rsSalvar.Fields(4) = ListView3.SelectedItem.ListSubItems.Item(2) 'Matriz encontrada no Processo Seletivo
        rsSalvar.Fields(5) = RemoveMask(ListView3.SelectedItem.ListSubItems.Item(6)) 'Nota calculada pelo Módulo avaliador para o cargo
        rsSalvar.Fields(6) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    'SALVAR ADMITIDOS DO PROCESSO SELETIVO - LISTVIEW4
    sqlDeletar = "Delete from tbProcessosAdm where tbProcessosAdm.codcoligada = '" & vCodcoligada & "' and tbProcessosAdm.codprocesso = '" & Val(txtProcesso(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbProcessosAdm where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView4.ListItems.Count
        ListView4.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtProcesso(0).Text) ' Código do Processo Seletivo
        rsSalvar.Fields(1) = ListView4.ListItems.Item(X) 'CPF Sugerido no Processo Seletivo
        rsSalvar.Fields(2) = ListView4.SelectedItem.ListSubItems.Item(9) 'Matriz Pesquisada no Processo Seletivo
        rsSalvar.Fields(3) = ListView4.SelectedItem.ListSubItems.Item(4) ' Tipo
        rsSalvar.Fields(4) = ListView4.SelectedItem.ListSubItems.Item(2) 'Matriz encontrada no Processo Seletivo
        rsSalvar.Fields(5) = RemoveMask(ListView4.SelectedItem.ListSubItems.Item(6)) 'Nota calculada pelo Módulo avaliador para o cargo
        rsSalvar.Fields(6) = ListView4.SelectedItem.ListSubItems.Item(8) 'Observação
        rsSalvar.Fields(7) = ListView4.SelectedItem.ListSubItems.Item(10) 'Registro do colaborador
        rsSalvar.Fields(8) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    
    cnBanco.CommitTrans
    rsSalvarProcSel.Close
    Set rsSalvarProcSel = Nothing
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    MsgBox "Os dados do Processo Seletivo foram salvos com sucesso", vbInformation, "SGCH"
    
    'AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtProcesso(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtProcesso(1).Tag, vbInformation, "Atenção"
        Me.txtProcesso(1).SetFocus
        Exit Function
    End If
    If txtProcesso(2).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtProcesso(1).Tag, vbInformation, "Atenção"
        Me.txtProcesso(1).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub ResultPesq()
    sqlProcesso = "Select a.*,b.nomerequisitante,b.datarequisicao from tbProcessos as a inner join tbrequisicoes as b on a.codcoligada = '" & vCodcoligada & "' and a.codrequisicao = b.codrequisicao Where a.codprocesso= '" & Val(varGlobal) & "' order by a.codprocesso"
    rsProcesso.Open sqlProcesso, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProcesso.RecordCount > 0 Then
        CompoeControles
        'Restaura ListView1
        Compoe_Listview1
        Restaura_Listview1
        
        'Restaura ListView2
        filtrarCandidatos
        Combo2.Clear
        CompoeComboLV2 Combo2
        MudaCorLV ListView2
        
        'Restaura ListView3
        Restaura_Listview3
        MudaCorLV ListView3
        
        'Restaura ListView4
        Restaura_Listview4
        MudaCorLV ListView4
        
        If rsProcesso.Fields(6) = "Fechado" Then BloqueiaControles Else lembrete
    Else
        MsgBox "Processo Seletivo não encontrado"
    End If
    
    rsProcesso.Close
    Set rsProcesso = Nothing
End Sub

Private Sub BloqueiaControles()
    'ListView1.Enabled = False
    'ListView2.Enabled = False
    'ListView3.Enabled = False
    'ListView4.Enabled = False
    txtProcesso(0).Enabled = False
    txtProcesso(1).Enabled = False
    txtProcesso(2).Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DTPicker3.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    cmdCadastro(0).Enabled = False
    cmdCadastro(1).Enabled = False
    cmdCadastro(2).Enabled = False
    cmdCadastro(3).Enabled = False
    cmdCadastro(4).Enabled = False
    cmdCadastro(5).Enabled = False
    cmdCadastro(6).Enabled = False
    cmdCadastro(9).Enabled = False
    cmdCadastro(11).Enabled = False
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
    If Status <> "novo" Then
    '    Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtProcesso(0), "000000"))
    '    ItemLst.SubItems(1) = DTPicker1
    '    ItemLst.SubItems(2) = cboCadRequisicao(1).Text
    '    ItemLst.SubItems(3) = txtCadReq(2).Text
    'Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = "Fechado"
    End If
    Exit Sub
Err:
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    txtProcesso(0).Text = Format(rsProcesso.Fields(0), "000000") 'código do Processo Seletivo
    DTPicker1 = rsProcesso.Fields(2) 'Data início do Processo Seletivo
    DTPicker2 = rsProcesso.Fields(3) 'Data fim do Processo Seletivo
    If rsProcesso.Fields(4) = 0 Then
        Check1.Value = 0 ' Listar
        Check2.Value = 0 ' Listar
    ElseIf rsProcesso.Fields(4) = 1 Then
        Check1.Value = 1 ' Listar
        Check2.Value = 0 ' Listar
    ElseIf rsProcesso.Fields(4) = 2 Then
        Check1.Value = 0 ' Listar
        Check2.Value = 1 ' Listar
    ElseIf rsProcesso.Fields(4) = 3 Then
        Check1.Value = 1 ' Listar
        Check2.Value = 1 ' Listar
    End If
    Combo1.Text = rsProcesso.Fields(5) ' Linhas por marcação
    txtProcesso(1).Text = Format(rsProcesso.Fields(1), "000000") 'codigo da Requisição
    txtProcesso(2).Text = rsProcesso.Fields(9) 'nome do requisitante
    DTPicker3 = rsProcesso.Fields(10) 'Data da Requisição
    If rsProcesso.Fields(7) = "S" Then Check3.Value = 1 Else Check3.Value = 0  'Informa se a requisição esta ativa ou nao
End Sub

Private Sub Restaura_Listview1()
    Dim ItemLst As ListItem
    Dim X As Integer
    sqlProcCargos = "Select * from tbProcessosCargos as a Where a.codcoligada = '" & vCodcoligada & "' and a.codprocesso = '" & Val(varGlobal) & "' order by a.codmatriz"
    rsProcCargos.Open sqlProcCargos, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsProcCargos.EOF
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            If Val(ListView1.ListItems.Item(X)) = rsProcCargos.Fields(1) Then
                ListView1.ListItems.Item(X).Checked = True
                ListView1.SelectedItem.ListSubItems.Item(4) = rsProcCargos.Fields(2)
            End If
        Next
        rsProcCargos.MoveNext
    Wend
    rsProcCargos.Close
    Set rsProcCargos = Nothing
End Sub

Private Sub Restaura_Listview3()
    Dim ItemLst As ListItem
    Dim X As Integer
    SqlcriaTabTemp = "Select a.cpf,b.nomecolaborador,a.matrizcargo,d.nomecargo,a.tipo,f.nomecargo,a.nota,f.nomecargo+substring(convert(char,a.nota,103),1,10),a.matrizpesq from tbProcessosParticipantes as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf " & _
    "inner join tbmatriz as c on a.matrizcargo = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo inner join tbmatriz as e on a.matrizpesq = e.codmatriz inner join tbcargos as f on e.codcargo = f.codcargo Where a.codprocesso = '" & Val(varGlobal) & "' order by f.nomecargo,a.tipo desc ,a.nota desc"
    rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
    While Not rscriaTabTemp.EOF
        Set ItemLst = ListView3.ListItems.Add(, , rscriaTabTemp.Fields(0)) 'CPF
        ItemLst.SubItems(1) = "" & rscriaTabTemp.Fields(1) 'Nome
        ItemLst.SubItems(2) = "" & Format(rscriaTabTemp.Fields(2), "000000")
        ItemLst.SubItems(3) = "" & rscriaTabTemp.Fields(3)
        ItemLst.SubItems(4) = "" & rscriaTabTemp.Fields(4)
        ItemLst.SubItems(5) = "" & rscriaTabTemp.Fields(5)
        ItemLst.SubItems(6) = "" & Format(rscriaTabTemp.Fields(6), "#,##0.00;(#,##0.00)") & " %"
        ItemLst.SubItems(7) = "" & rscriaTabTemp.Fields(7)
        ItemLst.SubItems(8) = "" & Format(rscriaTabTemp.Fields(8), "000000")
        rscriaTabTemp.MoveNext
    Wend
    rscriaTabTemp.Close
    Set rscriaTabTemp = Nothing
    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 7
    Me.ListView3.SortOrder = lvwDescending
End Sub

Private Sub Restaura_Listview4()
    Dim ItemLst As ListItem
    Dim X As Integer, vID As Integer
    SqlcriaTabTemp = "Select a.cpf,b.nomecolaborador,a.matrizcargo,d.nomecargo,a.tipo,f.nomecargo,a.nota,f.nomecargo+substring(convert(char,a.nota,103),1,10) as ordem,a.observacao,a.matrizpesq,a.codcolaborador,b.id from tbProcessosAdm as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf " & _
    "inner join tbmatriz as c on a.matrizcargo = c.codmatriz inner join tbcargos as d on c.codcargo = d.codcargo inner join tbmatriz as e on a.matrizpesq = e.codmatriz inner join tbcargos as f on e.codcargo = f.codcargo Where a.codprocesso = '" & Val(varGlobal) & "' order by f.nomecargo,a.tipo desc ,a.nota desc"
    rscriaTabTemp.Open SqlcriaTabTemp, cnBanco
    While Not rscriaTabTemp.EOF
        Set ItemLst = ListView4.ListItems.Add(, , rscriaTabTemp.Fields(0)) 'CPF
        ItemLst.SubItems(1) = "" & rscriaTabTemp.Fields(1) 'Nome
        ItemLst.SubItems(2) = "" & Format(rscriaTabTemp.Fields(2), "000000")
        ItemLst.SubItems(3) = "" & rscriaTabTemp.Fields(3)
        ItemLst.SubItems(4) = "" & rscriaTabTemp.Fields(4)
        ItemLst.SubItems(5) = "" & rscriaTabTemp.Fields(5)
        ItemLst.SubItems(6) = "" & Format(rscriaTabTemp.Fields(6), "#,##0.00;(#,##0.00)") & " %"
        ItemLst.SubItems(7) = "" & rscriaTabTemp.Fields(7)
        ItemLst.SubItems(8) = "" & rscriaTabTemp.Fields(8)
        ItemLst.SubItems(9) = "" & Format(rscriaTabTemp.Fields(9), "000000")
        If Not IsNull(rscriaTabTemp.Fields(10)) Then ItemLst.SubItems(10) = "" & rscriaTabTemp.Fields(10) Else ItemLst.SubItems(10) = "-"
        vID = rscriaTabTemp.Fields(11)
        ItemLst.SubItems(11) = comporControlesTotvs(vID)
        rscriaTabTemp.MoveNext
    Wend
    rscriaTabTemp.Close
    Set rscriaTabTemp = Nothing
    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 7
    Me.ListView4.SortOrder = lvwDescending
End Sub

Private Function comporControlesTotvs(vIdent As Integer)
    On Error Resume Next
    Dim rsContrTotvs As New ADODB.Recordset
    Dim SqlContrTotvs As String
        
    SqlContrTotvs = "select * from tbColaboradoresIntTotvs where codcoligada = '" & vCodcoligada & "'id = '" & vIdent & "'"
    rsContrTotvs.Open SqlContrTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    If rsContrTotvs.RecordCount > 0 Then
        comporControlesTotvs = rsContrTotvs.Fields(10)
    Else
        comporControlesTotvs = "-"
    End If
    rsContrTotvs.Close
    Set rsContrTotvs = Nothing
End Function

Private Function reservaVagaPS()
'On Error GoTo Err
    reservaVagaPS = False
    Dim contaVagas As Integer, X As Integer, Y As Integer
    cnBanco.BeginTrans
    contaVagas = 0
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        For Y = 1 To ListView4.ListItems.Count
            ListView4.ListItems.Item(Y).Selected = True
                        
            If ListView4.SelectedItem.ListSubItems.Item(10) = "" Or ListView4.SelectedItem.ListSubItems.Item(10) = "-" Then
                MsgBox "O Processo Seletivo não pode ser concluido. Existem colaboradores sem registro."
                cnBanco.RollbackTrans
                Exit Function
            End If
            
            If vintegrar = "S" Then
                If ListView4.SelectedItem.ListSubItems.Item(11) = "" Or ListView4.SelectedItem.ListSubItems.Item(11) = "-" Then
                    MsgBox "O Processo Seletivo não pode ser concluido. Dados de integração Totvs não informado."
                    cnBanco.RollbackTrans
                    Exit Function
                End If
            End If
            
            If Val(ListView4.SelectedItem.ListSubItems.Item(6)) < MediaGlobal And ListView4.SelectedItem.ListSubItems.Item(8) = "" Then
                MsgBox "O Processo Seletivo não pode ser concluido. Existem colaboradores que necessitam de justificativa para admissão."
                cnBanco.RollbackTrans
                Exit Function
            End If
            
            If Val(ListView1.ListItems.Item(X)) = Val(ListView4.SelectedItem.ListSubItems.Item(9)) Then
                contaVagas = contaVagas + 1
            End If
        Next
        
        If contaVagas < Val(ListView1.SelectedItem.ListSubItems.Item(4)) Then
            MsgBox "Dados salvos e NÃO concluidos. Faltam Preencher vagas para o cargo de: " & ListView1.SelectedItem.ListSubItems.Item(1)
            cnBanco.RollbackTrans
            Exit Function
        ElseIf contaVagas > Val(ListView1.SelectedItem.ListSubItems.Item(4)) Then
            MsgBox "Dados salvos e NÃO concluidos. Ultrapassou o limite de vagas para o cargo de: " & ListView1.SelectedItem.ListSubItems.Item(1)
            cnBanco.RollbackTrans
            Exit Function
        End If
        
        sqlRequisicoes = "Update tbRequisicoesCargos set status = 'Fechado', qtdocupada = '" & contaVagas & "' Where codrequisicao = '" & Val(txtProcesso(1)) & "' and codmatriz = '" & Val(ListView1.ListItems.Item(X)) & "'"
        rsRequisicoes.Open sqlRequisicoes, cnBanco
        
        sqlProcesso = "Update tbProcessos set status = 'Fechado' Where codprocesso = '" & Val(txtProcesso(0)) & "'"
        rsProcesso.Open sqlProcesso, cnBanco
        
        sqlProcesso = "Update tbRequisicoes set ativo = 'N', observacao = '' Where codrequisicao = '" & Val(txtProcesso(1)) & "'"
        rsProcesso.Open sqlProcesso, cnBanco
        
        contaVagas = 0
    Next
    
    'A ROTINA ABAIXO ADMITE O CANDIDATO OU ALTERA O COLABORADOR DE CARGO
    '-----------------------------
    For Y = 1 To ListView4.ListItems.Count
        ListView4.ListItems.Item(Y).Selected = True
            
        SqlSalvarNovoCol = "Update tbcolaboradores set ativo = 'S', tipo = 'colaborador', codrequisicao = '" & Val(txtProcesso(1).Text) & "', codcolaborador = '" & ListView4.SelectedItem.ListSubItems.Item(10) & "', obsadm = '" & ListView4.SelectedItem.ListSubItems.Item(8) & "' Where cpf = '" & ListView4.ListItems.Item(Y) & "'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
            
        SqlSalvarNovoCol = "Update tbcolaboradorescur set tipo = 'colaborador' Where cpf = '" & ListView4.ListItems.Item(Y) & "'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
            
        SqlSalvarNovoCol = "Update tbcolaboradoresesc set tipo = 'colaborador' Where cpf = '" & ListView4.ListItems.Item(Y) & "'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
            
        SqlSalvarNovoCol = "Update tbcolaboradoresexp set tipo = 'colaborador' Where cpf = '" & ListView4.ListItems.Item(Y) & "'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
            
        SqlSalvarNovoCol = "Update tbcolaboradoreshab set tipo = 'colaborador' Where cpf = '" & ListView4.ListItems.Item(Y) & "'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
        SqlSalvarNovoCol = "Update tbColaboradoreshist set ativo = 'N', datasai = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))) Where cpf = '" & ListView4.ListItems.Item(Y) & "' and tipo = 'colaborador' and ativo = 'S'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
        SqlSalvarNovoCol = "Update tbColaboradoreshist set ativo = 'N', datasai = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))) Where cpf = '" & ListView4.ListItems.Item(Y) & "' and tipo = 'candidato' and ativo = 'S'"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
        '-- Historico Funcional
        sqlDeletar = "Delete from tbColaboradoreshist where codcoligada = '" & vCodcoligada & "' and cpf = '" & ListView4.ListItems.Item(Y) & "' and ativo <> 'S' and tipo = 'candidato'"
        rsDeletar.Open sqlDeletar, cnBanco
        
        SqlSalvarNovoCol = "Insert into tbColaboradoreshist(cpf,codmatriz,data,ativo,sequencia,tipo,codrequisicao,codcoligada) Values(" & ListView4.ListItems.Item(Y) & ", " & Val(ListView4.SelectedItem.ListSubItems.Item(9)) & ", CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))), 'S', 1, 'colaborador', " & Val(txtProcesso(1)) & ",'" & vCodcoligada & "')"
        rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
        'Abaixo: Rotinas de gravacao de treinamentos para o novo cargo
        excluiProgramacao ListView4.ListItems.Item(Y), Val(ListView4.SelectedItem.ListSubItems.Item(9))
        GravaTreiPen ListView4.ListItems.Item(Y), Val(ListView4.SelectedItem.ListSubItems.Item(9))
        If GeraIntr = "S" Then GravaTreiIntrodutorio ListView4.ListItems.Item(Y), Val(ListView4.SelectedItem.ListSubItems.Item(9))
        If GeraObri = "S" Then GravaTreiObrigatorio ListView4.ListItems.Item(Y), Val(ListView4.SelectedItem.ListSubItems.Item(9))
        '-------------------------
        
        'VERIFICAR SE HÁ INTEGRAÇÃO TOTVS
        If vIntegra = "S" Then
            If ListView4.SelectedItem.ListSubItems.Item(11) = "" Or ListView4.SelectedItem.ListSubItems.Item(11) = "-" Then
                checkListINTD = False
                Exit Function
            Else
                
                Dim rsDadosTotvs As New ADODB.Recordset
                Dim SqlDBTotvs As String
                
                SqlDBTotvs = "Select a.nomecolaborador,a.datanascimento,a.ctpsnumero,a.foto,b.sexo,b.grauinst,b.tipoadm,b.motadm,b.forreceb,b.situacao,b.tipofunc,b.hortrab,b.funcao,b.secao,b.contsind,b.rais,b.memsind " & _
                "from tbColaboradores as a LEFT join tbColaboradoresIntTotvs as b on a.id = b.id where a.codcoligada = '" & vCodcoligada & " and a.codcolaborador = '" & ListView4.SelectedItem.ListSubItems.Item(10) & "'"
                rsDadosTotvs.Open SqlDBTotvs, cnBanco, adOpenKeyset, adLockReadOnly
                
                vDadosTotvs(0) = ListView4.SelectedItem.ListSubItems.Item(10) 'Chapa
                vDadosTotvs(1) = rsDadosTotvs.Fields(0) 'Nome do colaborador
                vDadosTotvs(2) = rsDadosTotvs.Fields(1) 'Data de nascimento
                vDadosTotvs(3) = rsDadosTotvs.Fields(2) 'Carteira de trabalho
                vDadosTotvs(4) = rsDadosTotvs.Fields(3) 'caminho foto
                vDadosTotvs(5) = rsDadosTotvs.Fields(4) 'sexo
                vDadosTotvs(6) = rsDadosTotvs.Fields(5) 'grau de instrução
                vDadosTotvs(7) = rsDadosTotvs.Fields(6) 'tipo de admissão
                vDadosTotvs(8) = rsDadosTotvs.Fields(7) 'motivo da admissão
                vDadosTotvs(9) = rsDadosTotvs.Fields(8) 'forma de recebimento
                vDadosTotvs(10) = rsDadosTotvs.Fields(9) 'Situação
                vDadosTotvs(11) = rsDadosTotvs.Fields(10) 'Tipo de funcionário
                vDadosTotvs(12) = rsDadosTotvs.Fields(11) 'horário de trabalho
                vDadosTotvs(13) = rsDadosTotvs.Fields(12) 'função
                vDadosTotvs(14) = rsDadosTotvs.Fields(13) 'seção
                vDadosTotvs(15) = rsDadosTotvs.Fields(14) 'contribuição sindical
                vDadosTotvs(16) = rsDadosTotvs.Fields(15) 'situação/rais
                vDadosTotvs(17) = rsDadosTotvs.Fields(16) 'membro sindicato
                For X = 0 To 17
                    If vDadosTotvs(X) = "" Then
                        MsgBox "Dados incompletos para exportar para o Corpore RM Labore. Favor conferir no cadastro", vbCritical, "SGCH"
                        GoTo Err
                    End If
                Next
                GravaDadosDBTotvs ListView4.SelectedItem.ListSubItems.Item(10)
                rsDadosTotvs.Close
            End If
        End If
    Next
    AtualizaListview
    '-----------------------------
    reservaVagaPS = True
    cnBanco.CommitTrans
    Exit Function
Err:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
End Function

'AS 4 ROTINAS ABAIXO SAO RESPONSAVEIS POR GRAVAR TODOS OS TREINAMENTOS
'DO NOVO COLABORADOR NA TABELA TBPENDENTESCUR
'TAIS ROTINAS DEVERAO SER GLOBALIZADAS
'Deixar GLOBAL as seguintes rotinas listadas abaixo:
'excluirProgramacao
'GravarTreiPen
'ApagaExcesso
'GeraIntr
'GeraObri
Private Sub excluiProgramacao(vCPF As String, vMatriz As Integer)
    ' Rotina deleta toda a programação "Agendada ou Pendente" se o
    ' colaborador sofrer alteração de cargo
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    sqlDeletar = "Delete from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and tbPendentesCur.cpf = '" & vCPF & "' and status = 'Pendente' and codmatriz <> '" & vMatriz & "' or tbPendentesCur.cpf = '" & vCPF & "' and status = 'Agendado' and codmatriz <> '" & vMatriz & "'"
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

Private Sub GravaTreiPen(vCPF As String, vMatriz As Integer)
    On Error Resume Next
    Dim rsGravaTreiPen As New ADODB.Recordset
    Dim SqlGravaTreiPen As String
    Dim rsPendentesCur As New ADODB.Recordset
    Dim SqlPendentesCur As String
    Dim contaID As Integer

    SqlGravaTreiPen = "Select a.codmatriz,a.codtreinamento,b.codtreinamento,b.cpf,a.codnivel from tbmatrizcur as a left join tbcolaboradorescur as b on a.codtreinamento = b.codtreinamento and b.codnivel >= a.codnivel and b.tipo = 'colaborador' and b.cpf = '" & vCPF & "' where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & vMatriz & "' order by a.codtreinamento"
    rsGravaTreiPen.Open SqlGravaTreiPen, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlPendentesCur = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id"
    rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not rsPendentesCur.EOF Then
        rsPendentesCur.MoveLast
        contaID = rsPendentesCur.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsPendentesCur.Close
    Set rsPendentesCur = Nothing
    
    While Not rsGravaTreiPen.EOF
        If IsNull(rsGravaTreiPen.Fields(2)) Then
            SqlPendentesCur = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and codtreinamento= '" & rsGravaTreiPen.Fields(1) & "' order by id"
            rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockOptimistic
            If rsPendentesCur.RecordCount = 0 Then
                rsPendentesCur.AddNew
                rsPendentesCur.Fields(0) = vCPF
                rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                rsPendentesCur.Fields(4) = "S"
                rsPendentesCur.Fields(5) = contaID
                rsPendentesCur.Fields(6) = "Pendente"
                rsPendentesCur.Fields(7) = 0
                If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(12) = 0 Else rsPendentesCur.Fields(12) = rsGravaTreiPen.Fields(4)
                rsPendentesCur.Fields(14) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            Else
                If rsPendentesCur.Fields(4) = "N" Then
                    rsPendentesCur.AddNew
                    rsPendentesCur.Fields(0) = vCPF
                    rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                    rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                    rsPendentesCur.Fields(4) = "S"
                    rsPendentesCur.Fields(5) = contaID
                    rsPendentesCur.Fields(6) = "Pendente"
                    rsPendentesCur.Fields(7) = 0
                    If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(12) = 0 Else rsPendentesCur.Fields(12) = rsGravaTreiPen.Fields(4)
                    rsPendentesCur.Fields(14) = vCodcoligada 'Codigo da coligada
                    contaID = contaID + 1
                Else
                End If
            End If
            rsPendentesCur.Update
            rsPendentesCur.Close
        End If
        rsGravaTreiPen.MoveNext
    Wend
    Set rsPendentesCur = Nothing
    
    rsGravaTreiPen.Close
    Set rsGravaTreiPen = Nothing
End Sub

Private Sub GravaTreiIntrodutorio(vCPF As String, vMatriz As Integer)
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiInt As New ADODB.Recordset
    Dim SqlSelecionaTreiInt As String
    Dim rsGravaTreiInt As New ADODB.Recordset
    Dim SqlGravaTreiInt As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & vMatriz & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
    
    If ListView4.ListItems.Count > 1 Then
        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    Else
        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    End If
    
    rsSelecionaTreiInt.Open SqlSelecionaTreiInt, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao from tbPendentesCur"
    rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiInt.EOF Then
        rsGravaTreiInt.MoveLast
        contaID = rsGravaTreiInt.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiInt.Close
    Set rsGravaTreiInt = Nothing
    
    While Not rsSelecionaTreiInt.EOF
        SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and codtreinamento ='" & rsSelecionaTreiInt.Fields(0) & "'"
        rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiInt.RecordCount = 0 Then
            rsGravaTreiInt.AddNew
            rsGravaTreiInt.Fields(0) = vCPF
            rsGravaTreiInt.Fields(1) = vMatriz
            rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
            rsGravaTreiInt.Fields(4) = "S"
            rsGravaTreiInt.Fields(5) = contaID
            rsGravaTreiInt.Fields(6) = "Pendente"
            rsGravaTreiInt.Fields(7) = 0
            rsGravaTreiInt.Fields(8) = 0
            rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiInt.Fields(4) = "N" Then
                rsGravaTreiInt.AddNew
                rsGravaTreiInt.Fields(0) = vCPF
                rsGravaTreiInt.Fields(1) = vMatriz
                rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
                rsGravaTreiInt.Fields(4) = "S"
                rsGravaTreiInt.Fields(5) = contaID
                rsGravaTreiInt.Fields(6) = "Pendente"
                rsGravaTreiInt.Fields(7) = 0
                rsGravaTreiInt.Fields(8) = 0
                rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiInt.Update
        rsGravaTreiInt.Close
        rsSelecionaTreiInt.MoveNext
    Wend
    Set rsGravaTreiInt = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiInt.Close
    Set rsSelecionaTreiInt = Nothing
End Sub

Private Sub GravaTreiObrigatorio(vCPF As String, vMatriz As Integer)
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiObr As New ADODB.Recordset
    Dim SqlSelecionaTreiObr As String
    Dim rsGravaTreiObr As New ADODB.Recordset
    Dim SqlGravaTreiObr As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & vMatriz & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiObr = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel from tbPendentesCur"
    rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiObr.EOF Then
        rsGravaTreiObr.MoveLast
        contaID = rsGravaTreiObr.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiObr.Close
    Set rsGravaTreiObr = Nothing
    
    While Not rsSelecionaTreiObr.EOF
        SqlGravaTreiObr = "Select a.cpf,a.codmatriz,a.codtreinamento,a.codprogramacao,a.ativo,a.id,a.status,a.tipoprogramacao,a.codnivel,a.codcoligada from tbPendentesCur as a  left join tbTreinamentosNiv as b on a.codnivel = b.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & vCPF & "' and a.codtreinamento ='" & rsSelecionaTreiObr.Fields(0) & "'"
        rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiObr.RecordCount = 0 Then
            rsGravaTreiObr.AddNew
            rsGravaTreiObr.Fields(0) = vCPF
            rsGravaTreiObr.Fields(1) = vMatriz
            rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
            rsGravaTreiObr.Fields(4) = "S"
            rsGravaTreiObr.Fields(5) = contaID
            rsGravaTreiObr.Fields(6) = "Pendente"
            rsGravaTreiObr.Fields(7) = 0
            rsGravaTreiObr.Fields(8) = 0
            rsGravaTreiObr.Fields(9) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiObr.Fields(4) = "N" Then
                rsGravaTreiObr.AddNew
                rsGravaTreiObr.Fields(0) = vCPF
                rsGravaTreiObr.Fields(1) = vMatriz
                rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
                rsGravaTreiObr.Fields(4) = "S"
                rsGravaTreiObr.Fields(5) = contaID
                rsGravaTreiObr.Fields(6) = "Pendente"
                rsGravaTreiObr.Fields(7) = 0
                rsGravaTreiObr.Fields(8) = 0
                rsGravaTreiObr.Fields(9) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiObr.Update
        rsGravaTreiObr.Close
        rsSelecionaTreiObr.MoveNext
    Wend
    Set rsGravaTreiObr = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiObr.Close
    Set rsSelecionaTreiObr = Nothing
End Sub

Private Sub AchaItem()
On Error GoTo Err
    Dim X As Integer, Y As Integer
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        If ListView4.ListItems.Item(X).Selected = True Then
            AddDadosGeral(0) = ListView4.ListItems.Item(X) 'CPF
            AddDadosGeral(1) = ListView4.SelectedItem.ListSubItems.Item(1) 'Nome
            AddDadosGeral(2) = ListView4.SelectedItem.ListSubItems.Item(2) 'Matriz encontrada do candidato
            AddDadosGeral(3) = ListView4.SelectedItem.ListSubItems.Item(3) 'Cargo encontrado do candidato
            AddDadosGeral(4) = ListView4.SelectedItem.ListSubItems.Item(4) 'Tipo (colaborador/candidato)
            AddDadosGeral(5) = ListView4.SelectedItem.ListSubItems.Item(5) 'cargo requisitador
            AddDadosGeral(6) = ListView4.SelectedItem.ListSubItems.Item(6) 'Nota
            AddDadosGeral(7) = ListView4.SelectedItem.ListSubItems.Item(8) 'observação
            AddDadosGeral(9) = ListView4.SelectedItem.ListSubItems.Item(10) 'observação
        End If
    Next
Err:
    Exit Sub
End Sub

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate
If Button = vbLeftButton Then
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 5 To 5
            leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView1.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView1.SelectedItem.Text
                    Else
                        .Text = ListView1.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
                Exit For
            End If
        Next i
    End If
End If
End Sub

Function GetLvwDeltaX() As Single
    Dim si As SCROLLINFO, maxScrollPos As Long
    Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView1.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView1.Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView1.Left Then txtLeft = ListView1.Left + 48
    
        txtRightMax = ListView1.Left + ListView1.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView1.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView1.Left + ListView1.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView1.Top
        If Not ListView1.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView1.Top + ListView1.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView1.Top + ListView1.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
    
        With txtLvw '.move produces runtime error with -ve values
            If txtLeft < 11000 Then .Left = txtLeft + 620 Else .Left = txtLeft - 140
            .Top = txtTop + 2415
            .Width = txtWidth - 530
            .Height = ListView1.SelectedItem.Height - 50
            
        End With
    End If
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView1 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
        'other keys can be used for navigation
    End Select
    If txtLvw.Text = "-" Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw <> "" And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLvw_LostFocus()
'On Error GoTo TrataErro
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "0"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView1.ListItems(ListView1.ListItems.Count) <> c_EntryTxt Then ListView1.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    If ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) = "-" And ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) < ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
        Exit Sub
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
    End If
    
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    ListView1.SetFocus
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub lembrete()
    If Date > DTPicker2 Then
        Label7.Visible = True
    End If
End Sub

Private Sub configControles()
    If vInc = "N" Then 'Incluir
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then 'Editar
        cmdCadastro(5).UseGreyscale = True
        cmdCadastro(5).DragMode = 1
        cmdCadastro(5).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then 'Salvar
        cmdCadastro(11).UseGreyscale = True
        cmdCadastro(11).DragMode = 1
        cmdCadastro(11).SpecialEffect = cbEngraved
    End If
    If vAdi = "N" Then 'Salvar
        cmdCadastro(4).UseGreyscale = True
        cmdCadastro(4).DragMode = 1
        cmdCadastro(4).SpecialEffect = cbEngraved
    End If
    
    If vFil = "N" Then 'Filtrar
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    End If
    
    If vExc = "N" Then 'Excluir
        cmdCadastro(6).UseGreyscale = True
        cmdCadastro(6).DragMode = 1
        cmdCadastro(6).SpecialEffect = cbEngraved
    
        cmdCadastro(9).UseGreyscale = True
        cmdCadastro(9).DragMode = 1
        cmdCadastro(9).SpecialEffect = cbEngraved
    End If
End Sub

