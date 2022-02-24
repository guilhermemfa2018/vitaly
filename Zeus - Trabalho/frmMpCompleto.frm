VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmMPCompleto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Metodos e Processos"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19545
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMpCompleto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   19545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame24 
      Caption         =   "Clonar OS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   16320
      TabIndex        =   129
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "Clonar"
         Height          =   375
         Index           =   19
         Left            =   1800
         TabIndex        =   131
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   130
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMpCompleto.frx":0CCA
         TabIndex        =   132
         Top             =   240
         Width           =   1455
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   375
      Left            =   5160
      OleObjectBlob   =   "frmMpCompleto.frx":0D2E
      TabIndex        =   111
      Top             =   9720
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Frame Frame18 
      Caption         =   "Peso Posição (APENAS TESTE)"
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
      Left            =   14520
      TabIndex        =   109
      Top             =   9600
      Visible         =   0   'False
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMpCompleto.frx":0DFC
         TabIndex        =   110
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2040
      TabIndex        =   106
      Top             =   9720
      Width           =   2655
   End
   Begin VB.TextBox label53 
      Height          =   330
      Left            =   7080
      TabIndex        =   70
      Text            =   "-"
      Top             =   9960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sequencial "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14640
      TabIndex        =   68
      Top             =   120
      Width           =   1335
      Begin VB.TextBox txtformula 
         Alignment       =   2  'Center
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         HideSelection   =   0   'False
         Index           =   15
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "-"
         Top             =   360
         Width           =   1095
      End
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
      Index           =   13
      Left            =   720
      Picture         =   "frmMpCompleto.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   9600
      Width           =   615
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
      Index           =   12
      Left            =   120
      Picture         =   "frmMpCompleto.frx":1B20
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9600
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Width           =   19335
      _ExtentX        =   34105
      _ExtentY        =   14631
      _Version        =   393216
      Tab             =   1
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
      TabCaption(0)   =   "Desenhos"
      TabPicture(0)   =   "frmMpCompleto.frx":27EA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(1)=   "Frame31"
      Tab(0).Control(2)=   "cmdCadastro(7)"
      Tab(0).Control(3)=   "cmdCadastro(8)"
      Tab(0).Control(4)=   "cmdCadastro(4)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Recursos"
      TabPicture(1)   =   "frmMpCompleto.frx":2806
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "aicAlphaImage2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "ListView1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdCadastro(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdCadastro(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdCadastro(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdCadastro(3)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame6"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame7"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "ScriptControl1"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Frame13"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Frame14"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "SkinLabel6"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "SkinLabel7"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtformula(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtformula(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cmdCadastro(10)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtformula(25)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtDB"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtLV"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "SkinLabel18"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtformula(26)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "SkinLabel16"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Combo1"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Frame19"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmdCadastro(17)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Frame20"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Timer2"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Frame25"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Frame26"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).ControlCount=   37
      TabCaption(2)   =   "Ordem de Serviço"
      TabPicture(2)   =   "frmMpCompleto.frx":2822
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "Frame16"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame26 
         Caption         =   "Tempo Orçado (OP)"
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
         Left            =   14880
         TabIndex        =   136
         Top             =   7560
         Width           =   1935
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":283E
            TabIndex        =   137
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Tempo apropriado (OP)"
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
         Left            =   16920
         TabIndex        =   134
         Top             =   7560
         Width           =   2295
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":2898
            TabIndex        =   135
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   2520
         Top             =   360
      End
      Begin VB.Frame Frame20 
         Caption         =   "Histórico "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   6840
         TabIndex        =   122
         Top             =   960
         Width           =   3855
         Begin VB.Frame Frame27 
            Caption         =   "Peso "
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
            Left            =   840
            TabIndex        =   139
            Top             =   3000
            Width           =   1215
            Begin VB.TextBox Text11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               BorderStyle     =   0  'None
               ForeColor       =   &H000000C0&
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Tempo Total"
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
            TabIndex        =   127
            Top             =   3000
            Width           =   1575
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000A&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   120
               TabIndex        =   128
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Sequencial"
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
            Left            =   1560
            TabIndex        =   125
            Top             =   3000
            Visible         =   0   'False
            Width           =   1215
            Begin VB.TextBox txtformula 
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   28
               Left            =   120
               TabIndex        =   126
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   18
            Left            =   120
            Picture         =   "frmMpCompleto.frx":28F2
            Style           =   1  'Graphical
            TabIndex        =   124
            Top             =   3120
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView4 
            Height          =   2775
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   4895
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
         Caption         =   "Agregar"
         Height          =   495
         Index           =   17
         Left            =   8280
         TabIndex        =   121
         Tag             =   "Inclui um item selecionado no LV à uma OS"
         ToolTipText     =   "Inclui um item selecionado no LV à uma OS"
         Top             =   7680
         Width           =   1455
      End
      Begin VB.Frame Frame19 
         Caption         =   "Legenda: "
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
         Left            =   14640
         TabIndex        =   112
         Top             =   3120
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "Cinza:"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   330
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   119
            Text            =   "Verde:"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   118
            Text            =   "Preto:"
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   495
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":35BC
            TabIndex        =   113
            Top             =   360
            Width           =   3015
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmMpCompleto.frx":3670
            TabIndex        =   114
            Top             =   840
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmMpCompleto.frx":36EA
            TabIndex        =   115
            Top             =   1080
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmMpCompleto.frx":3764
            TabIndex        =   116
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "frmMpCompleto.frx":37D0
         Left            =   18000
         List            =   "frmMpCompleto.frx":37F2
         TabIndex        =   105
         Tag             =   "Operação nº"
         ToolTipText     =   "Operação nº"
         Top             =   600
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   18000
         OleObjectBlob   =   "frmMpCompleto.frx":381F
         TabIndex        =   104
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtformula 
         Height          =   330
         Index           =   26
         Left            =   9000
         TabIndex        =   15
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   600
         Width           =   8895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   9000
         OleObjectBlob   =   "frmMpCompleto.frx":3895
         TabIndex        =   102
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtLV 
         Height          =   330
         Left            =   3000
         TabIndex        =   101
         Text            =   "LV"
         Top             =   7560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtDB 
         Height          =   330
         Left            =   1560
         TabIndex        =   100
         Text            =   "DB"
         Top             =   7920
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtformula 
         Height          =   330
         Index           =   25
         Left            =   1560
         TabIndex        =   99
         Text            =   "ID OS"
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame16 
         Caption         =   "Serviços de Terceiros "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7695
         Left            =   -64680
         TabIndex        =   93
         Top             =   480
         Width           =   8895
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   16
            Left            =   6720
            Picture         =   "frmMpCompleto.frx":3903
            Style           =   1  'Graphical
            TabIndex        =   103
            Tag             =   "Cadastrar Serviços de Terceiros"
            ToolTipText     =   "Cadastrar Serviços de Terceiros"
            Top             =   2760
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   4095
            Left            =   120
            TabIndex        =   33
            Top             =   3480
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   7223
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   15
            Left            =   1320
            Picture         =   "frmMpCompleto.frx":45CD
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   14
            Left            =   720
            Picture         =   "frmMpCompleto.frx":5297
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   11
            Left            =   120
            Picture         =   "frmMpCompleto.frx":5F61
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   24
            Left            =   120
            TabIndex        =   29
            Top             =   2400
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":6C2B
            TabIndex        =   98
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtformula 
            Height          =   855
            Index           =   23
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1200
            Width           =   8655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":6C99
            TabIndex        =   97
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   255
            Left            =   8400
            TabIndex        =   96
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   22
            Left            =   840
            TabIndex        =   27
            Top             =   480
            Width           =   7455
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   21
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmMpCompleto.frx":6D05
            TabIndex        =   95
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":6D6D
            TabIndex        =   94
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Dados da OS - Ordem de Serviço "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7695
         Left            =   -74880
         TabIndex        =   87
         Top             =   480
         Width           =   10095
         Begin VB.TextBox txtformula 
            Height          =   6495
            Index           =   20
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   1080
            Width           =   9855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":6DCD
            TabIndex        =   92
            Top             =   840
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   2760
            TabIndex        =   23
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   172752897
            CurrentDate     =   41568
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmMpCompleto.frx":6E3B
            TabIndex        =   91
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   19
            Left            =   4440
            TabIndex        =   24
            Top             =   480
            Width           =   5535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   4440
            OleObjectBlob   =   "frmMpCompleto.frx":6E9D
            TabIndex        =   90
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   1920
            TabIndex        =   22
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "frmMpCompleto.frx":6F15
            TabIndex        =   89
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":6F7D
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   10
         Left            =   120
         Picture         =   "frmMpCompleto.frx":6FDB
         Style           =   1  'Graphical
         TabIndex        =   86
         Tag             =   "Gerar OS - Ordem de Serviço"
         ToolTipText     =   "Gerar OS - Ordem de Serviço"
         Top             =   7605
         Width           =   615
      End
      Begin VB.TextBox txtformula 
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
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox txtformula 
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Tag             =   "ID Centro de Custo"
         ToolTipText     =   "ID Centro de Custo"
         Top             =   600
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmMpCompleto.frx":7CA5
         TabIndex        =   84
         Top             =   360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMpCompleto.frx":7D11
         TabIndex        =   83
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame Frame14 
         Caption         =   "Tempo total (OS)"
         Height          =   615
         Left            =   11880
         TabIndex        =   71
         Top             =   7560
         Width           =   2895
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   195
            Width           =   1815
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Grupo "
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
         Left            =   2400
         TabIndex        =   66
         Top             =   4080
         Width           =   4335
         Begin ACTIVESKINLibCtl.SkinLabel Label8 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmMpCompleto.frx":7D79
            TabIndex        =   77
            Top             =   240
            Width           =   4095
         End
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   10560
         Top             =   4200
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Frame Frame7 
         Caption         =   "Semana Programada"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4680
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
         Begin VB.TextBox Text9 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   1440
            TabIndex        =   133
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   405
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   172752897
            CurrentDate     =   41556
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tempo calculado (Min)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         TabIndex        =   50
         Top             =   3240
         Width           =   2175
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   120
            Top             =   240
         End
         Begin VB.TextBox txtResultado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   525
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   51
            Tag             =   "Tempo calculado"
            ToolTipText     =   "Tempo calculado"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   4
         Left            =   -65640
         Picture         =   "frmMpCompleto.frx":7DD3
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "Limpar Controles"
         ToolTipText     =   "Limpar Controles"
         Top             =   7200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1320
         Picture         =   "frmMpCompleto.frx":8A9D
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   720
         Picture         =   "frmMpCompleto.frx":9767
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   13
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmMpCompleto.frx":A431
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4200
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Figura "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   10800
         TabIndex        =   46
         Top             =   960
         Width           =   4455
         Begin VB.PictureBox Picture1 
            Height          =   3495
            Left            =   120
            ScaleHeight     =   3435
            ScaleWidth      =   4155
            TabIndex        =   76
            Top             =   240
            Width           =   4215
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   3495
               Left            =   0
               Top             =   0
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   6165
               Image           =   "frmMpCompleto.frx":B0FB
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fórmulas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   15360
         TabIndex        =   44
         Top             =   960
         Width           =   3855
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   45
            Top             =   2880
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComctlLib.TreeView TreeView3 
            Height          =   3495
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6165
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "DIcas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   6615
         Begin VB.TextBox txtformula 
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H000000C0&
            Height          =   1815
            Index           =   6
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   43
            Tag             =   "Fórmula"
            ToolTipText     =   "Fórmula"
            Top             =   240
            Width           =   6375
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Variáveis "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   3240
         Width           =   2175
         Begin VB.TextBox txtformula 
            Height          =   330
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Tag             =   "Insira as variáveis de acordo com a Observação acima"
            ToolTipText     =   "Insira as variáveis de acordo com a Observação acima"
            Top             =   360
            Width           =   1935
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Top             =   4920
         Width           =   19095
         _ExtentX        =   33681
         _ExtentY        =   4683
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "<"
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
         Index           =   8
         Left            =   -65760
         TabIndex        =   10
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   ">"
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
         Index           =   7
         Left            =   -65760
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.Frame Frame31 
         Caption         =   "Itens selecionados "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -64920
         TabIndex        =   39
         Top             =   480
         Width           =   9015
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   6240
            Visible         =   0   'False
            Width           =   8775
         End
         Begin VB.Frame Frame3 
            Caption         =   "Peso Total Selecionado"
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
            TabIndex        =   49
            Top             =   6600
            Width           =   8775
            Begin ACTIVESKINLibCtl.SkinLabel Label3 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmMpCompleto.frx":B113
               TabIndex        =   75
               Top             =   240
               Width           =   8535
            End
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   6255
            Left            =   120
            TabIndex        =   9
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   11033
            _Version        =   393217
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
      Begin VB.Frame Frame11 
         Caption         =   "Desenhos disponíveis "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   9015
         Begin VB.Frame Frame41 
            Caption         =   "Peso Total Selecionado"
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
            TabIndex        =   40
            Top             =   6600
            Width           =   8775
            Begin ACTIVESKINLibCtl.SkinLabel Label6 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmMpCompleto.frx":B16D
               TabIndex        =   74
               Top             =   240
               Width           =   8535
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   6255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   11033
            _Version        =   393217
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
      Begin VB.Frame Frame12 
         Caption         =   "Constantes "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   11400
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
         Begin MSComctlLib.ListView ListView2 
            Height          =   2895
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   5106
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   4210752
            BackColor       =   16777215
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
      Begin VB.Frame Frame17 
         Caption         =   "Status (0-nada/1-aberta/2-andamento/3-fechada)"
         Height          =   615
         Left            =   3600
         TabIndex        =   107
         Top             =   7560
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox txtformula 
            Height          =   330
            Index           =   27
            Left            =   120
            TabIndex        =   108
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Decoder "
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
         TabIndex        =   62
         Top             =   1920
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtDecoder 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   240
            Width           =   7335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Referências "
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
         TabIndex        =   53
         Top             =   960
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   10
            Left            =   2160
            TabIndex        =   59
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   7
            Left            =   2160
            TabIndex        =   56
            Top             =   600
            Visible         =   0   'False
            Width           =   5415
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtformula 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2160
            TabIndex        =   54
            Top             =   480
            Width           =   5415
         End
         Begin VB.Label Label7 
            Caption         =   "Parâmetros:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Fórmula:"
            Height          =   255
            Left            =   2160
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
         Height          =   435
         Left            =   9960
         Top             =   7680
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         Image           =   "frmMpCompleto.frx":B1C7
         Props           =   5
      End
   End
   Begin VB.Frame Frame21 
      Caption         =   "Dados Método e Processo "
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
      TabIndex        =   36
      Top             =   120
      Width           =   14415
      Begin VB.ComboBox Combo2 
         Height          =   345
         ItemData        =   "frmMpCompleto.frx":BF74
         Left            =   12600
         List            =   "frmMpCompleto.frx":BF81
         TabIndex        =   142
         Text            =   "Fabricação"
         Top             =   480
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
         Height          =   255
         Left            =   12600
         OleObjectBlob   =   "frmMpCompleto.frx":BFA7
         TabIndex        =   141
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   9960
         TabIndex        =   138
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtformula 
         Height          =   330
         Index           =   16
         Left            =   6120
         TabIndex        =   85
         Text            =   "ID Projeto"
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtformula 
         Enabled         =   0   'False
         Height          =   345
         Index           =   14
         Left            =   8760
         TabIndex        =   5
         Tag             =   "Nome do Responsável"
         ToolTipText     =   "Nome do Responsável"
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtformula 
         Height          =   345
         Index           =   13
         Left            =   4560
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtformula 
         Height          =   345
         Index           =   12
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1680
         TabIndex        =   1
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
         Format          =   169213953
         CurrentDate     =   41554
      End
      Begin VB.TextBox txtformula 
         Enabled         =   0   'False
         Height          =   345
         HelpContextID   =   1
         Index           =   11
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   8760
         OleObjectBlob   =   "frmMpCompleto.frx":C00F
         TabIndex        =   82
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   4560
         OleObjectBlob   =   "frmMpCompleto.frx":C07F
         TabIndex        =   81
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmMpCompleto.frx":C0E7
         TabIndex        =   80
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmMpCompleto.frx":C147
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMpCompleto.frx":C1BD
         TabIndex        =   78
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   8160
         TabIndex        =   73
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   11760
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMPCompleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'my declarations
Private Const c_EntryTxt = ""
Private m_ColIndex As Long 'listview col index
Private m_RowIndex As Long 'listview row index

'Variaveis que irao receber os valores referente aos parametros das formulas
'para localizar os dados na tabela de classificação

'Variaveis que irão receber os dados da tabela de classificação após a localizacao
Private vTMedio As Double '
Private vFFadiga As Double
Private vOrganiza As Double
Private vSomaTempo As Double

'Variáveis que irão receber os dados do textBox de parametro para realizar a localização na
'tabela de parametros
Private vGrupo As String
Private vDimTipo As String
Private vDimValor As String
Private vInterTipo As String
Private vInterValor As String
Private vStatus As Double
Private Status As String

Private var(50) As Double
Private cons(50) As Double
'---------------------------------------------------

Private vNomeA As String
Private vNomeB As String
Private vNomeC As String
Private vJuntaNome As String
Private vPesoTotal1 As Double
Private vPesoTotal2 As Double
Private vPesoPosicao As Double
Private vAcumulaTempo As Double

Private vAcumula As String
Private vNmNo As String
Private vPAutomatico As String

Private vPonte1 As TextBox
Private vPonte2 As TextBox
Private vPonte3 As TextBox
Private vPonte4 As TextBox
Private vPonte5 As TextBox

Private rsFCE As New ADODB.Recordset
Private sqlFCE As String
Private rsProjeto As New ADODB.Recordset
Private SqlProjeto As String
Private rsProg As New ADODB.Recordset
Private SqlProg As String
'Private vTime As String

Private Sub cmdCadastro_Click(Index As Integer)
On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Select Case Index
    Case 0
        'Chama CC - Centro de Custo
        Label8 = "-"
        ChamaGrid "CORPORERM.dbo.GCCUSTO", "codreduzido", txtformula(0), frmMPCompleto, "codreduzido", "nome"
        CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtformula(0), txtformula(1), 7, 2, txtformula(0), "S", txtformula(1), "1"
        montaEstrutTreeview
        LimpaVariaveis
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
        
        'Procurar onde esta limpando a tabela "tbMPDesSel" & vTime
        compoeDadosLVs
    Case 1
        Status = "novo"
        'Primeiro gera o ID da Programação para lançar no LV
        Dim CodID As String
        Dim rsGeraID As New ADODB.Recordset
        Dim sqlGeraID As String
        Dim X As Integer
        
        ' CHAMA FUNÇÃO QUE CONVERTE SEMANA DO ANO PARA DATA
        DTPicker2.Value = ""
        converteSemana Val(Text9.Text), DTPicker2, ""
        If DTPicker2.Value = "" Then
            mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
            Exit Sub
        End If
        
        sqlGeraID = "Select * from tbMP where tbMP.idprogramacao= " & Val(Me.txtformula(11))
        rsGeraID.Open sqlGeraID, cnBanco, adOpenKeyset, adLockOptimistic
        CodID = 0
        If txtformula(11).Text = "" Then 'Código do Cliente/Comitente
            rsGeraID.AddNew
            CodID = Format(GeraCodigoTB("tbMP", "idprogramacao", "", ""), "000000")
            rsGeraID.Fields(0) = CodID
            txtformula(11) = CodID
            'rsGeraID.Close
            'Set rsGeraID = Nothing
        Else
            CodID = txtformula(11)
            txtformula(11) = CodID
        End If
        
        If Text8 <> "" And txtformula(11) <> "" Then
            ClonaHist
        End If
        
        'salvar_Dados
        
        'Cria TextBox em tempo de Execução
        txtLV = Val(txtformula(11).Text) & Val(vPonte1) & Val(txtformula(15)) & Val(Combo1.Text)
        If txtformula(26) = "" Then txtformula(26) = "-"
        If vPonte1 = "" Then vPonte1 = 0
        
        'If DTPicker2.Value <> "" Then
        '    vPonte2.Text = DTPicker2.Value
        'Else
        '    vPonte2.Text = "-"
        'End If
        If Not IsNull(rsGeraID.Fields(1)) Then
            vPonte2.Text = "-"
        End If
        
        vPonte3.Text = Label8.Caption
        vPonte4.Text = Combo1.Text
        vPonte5.Text = vPonte1 & Format(Combo1.Text, "000")
        txtformula(5).Text = "-"
        If ValidaCampos(ListView1, txtformula(0), txtformula(6), Text1, vPonte3, txtformula(5), vPonte4, txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0)) = False Then Exit Sub
        If IncluirLV(ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), vPonte5) = False Then
            Exit Sub
        End If
        
        
        ListView1.Sorted = True
        ListView1.SortKey = 11
        ListView1.SortOrder = lvwAscending
        
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.SelectedItem.ListSubItems.Item(5) = "" Then
                ListView1.SelectedItem.ListSubItems.Item(5) = vPonte2
            End If
        Next
        ListView1.ListItems.Item(1).Selected = True
        
        
        vPonte1.Text = "0"
        txtformula(27) = "0"
        LimpaVariaveis
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(26), txtformula(0)
        LimpaControles txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1)
        Combo1.Text = ""
        txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        SomaLV ListView1, 6, Text2
        TreeView3.Nodes.Clear
        aicAlphaImage1.ClearImage
        
        'SALVA OS DADOS A CADA VEZ QUE UM ITEM É INCLUIDO NO LISTVIEW
        'DEIXA O SISTEMA BEM MAIS LENTO
        
            salvar_Dados
        If ListView4.ListItems.Count > 0 Then salvar_dados_hist
        Text7.Text = ""
        ListView4.ListItems.Clear
    Case 2
        EditaLVMP
    Case 3
        Status = "excluir"
        Dim G As Integer, i As Integer
        i = ListView1.ListItems.Count
        For G = 1 To i
            If ListView1.ListItems.Item(G).Selected = True And ListView1.SelectedItem.ListSubItems.Item(5) <> "" And ListView1.SelectedItem.ListSubItems.Item(5) <> "-" Then
                mobjMsg.Abrir "Operação já possui programação, não pode ser EXCLUIDA", Ok, critico, "ZEUS"
                Exit For
            End If
            If ListView1.ListItems.Item(G).Selected = True And SkinLabel27 = "0000:00" Then
                ExcluirItemLV ListView1
                LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(26), txtformula(2)
                txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
                SomaLV ListView1, 6, Text2
                Exit For
            ElseIf ListView1.ListItems.Item(G).Selected = True And SkinLabel27 <> "0000:00" Then
                mobjMsg.Abrir "Operação já encontra-se em apropriação, não pode ser EXCLUIDA", Ok, critico, "ZEUS"
                Exit For
            End If
        Next
    Case 4
        'Em teste criado agora
        'Esvazia a tabela tbMPDesSel
        'Dim rsDeletar As New ADODB.Recordset
        'Dim sqlDeletar As String
        sqlDeletar = "Delete from tbMPDesSel" & vTime
        rsDeletar.Open sqlDeletar, cnBanco
        
        
        TreeView2.Nodes.Clear
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        Label3 = "-"
    Case 5
        ChamaGridFCE
        CarregaFCE
    Case 6
        If txtformula(12).Text <> "" Then
            ChamaGridProjeto
            CarregaProjeto
            mostraDesenhos "tbitemlm", TreeView1
            txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        End If
    Case 7
        'teste
        Status = "novo"
        
        'sqlDeletar = "Delete from tbMPDesSel"
        'rsDeletar.Open sqlDeletar, cnBanco
        
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        vAcumulaTempo = 0
        buscaChecado2 TreeView1
        mostraDesenhos "tbMPDesSel" & vTime, TreeView2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
        
        If vPAutomatico = "S" Then
            'Se o calculo for automatico, irá desativar o textbox txtformula(5)
            txtformula(5).Enabled = False
            'Exibe o resultado dos calculos no textbox txtResultado
            txtResultado = Format(vAcumulaTempo, "#,##0.00;(#,##0.00)")
            'Limpa o textbox txtformula(5) após realizar todos os cálculos e exibir o resultado
            'no textbox txtResultado
            'txtformula(5).Text = ""
            txtformula(5).Text = "AUTOMÁTICO"
        Else
            'Se o calculo NÃO for automatico, irá ativar o textbox txtformula(5)
            If vStatus <= 1 Then
                txtformula(5).Enabled = True
            End If
        End If
    Case 8
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        buscaChecado2 TreeView2
        mostraDesenhos "tbMPDesSel" & vTime, TreeView2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 9
        ChamaGrid "tbUsuarios", "nome", txtformula(14), frmMPCompleto, "codigo", "nome"
        txtformula(14) = Mid$(Pesquisa, 1, 6) & " - " & Mid$(Pesquisa, 7, 20)
    Case 10
        'mobjMsg.Abrir "Rotina que irá gerar número da OS (Em desenvolvimento)", Ok, informacao, "ZEUS"
        LimpaControles txtformula(17), txtformula(18), txtformula(19), txtformula(20), txtformula(17), txtformula(17), txtformula(17), txtformula(17), txtformula(17), txtformula(17)
        txtDB = ""
        txtDB = Format(GeraCodigoTB("tbOS", "idos", "", ""), "000000000")
        txtLV = ""
        txtLV = Format(GeraOSLV(ListView1), "000000000")
        txtformula(25).Text = ""
        If Val(txtDB) = Val(txtLV) Then
            txtformula(25).Text = Format(txtDB, "000000000")
        ElseIf Val(txtDB) > Val(txtLV) Then
            txtformula(25).Text = Format(txtDB, "000000000")
        ElseIf Val(txtDB) < Val(txtLV) Then
            txtformula(25).Text = Format(txtLV, "000000000")
        End If
        If MarcaOS = False Then Exit Sub
        txtformula(17).Text = txtformula(25).Text
        txtformula(18).Text = 0
        verOS 'Verifica se há alguma OS ativa
    Case 11 'Incluir Serviços de terceiros
        If ValidaCampos(ListView3, txtformula(21), txtformula(22), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)) = False Then Exit Sub
        IncluirLV ListView3, txtformula(21), txtformula(22), txtformula(23), txtformula(24), txtformula(17), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)
        LimpaControles txtformula(21), txtformula(22), txtformula(23), txtformula(24), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)
    Case 12
        If salvar_Dados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
        Else
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ZEUS"
        End If
    Case 13 'Sair do formulário
        excluiTabela
        Unload Me
    Case 14
        AlteraLV ListView3, txtformula(21), txtformula(22), txtformula(23), txtformula(24), txtformula(17), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)
    Case 15
        ExcluirItemLV ListView3
    Case 16
        frmServTerc.Show 1
    Case 17
        'Rotina para agregar
        AgregarOS
        If AgregarOS = True Then
            Msgbox "Procedimento realizado com sucesso!", vbInformation, "ZEUS"
        '    mobjMsg.Abrir "Procedimento realizado com sucesso!", , informacao
        Else
        '    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
        End If
        'Exit Sub
    Case 18
        Text7 = ""
        ExcluirItemLV ListView4
        SomaLV ListView4, 2, Text7
        pesoTempo
        Timer1.Enabled = True
    Case 19
        If Text8.Text = "" Then
            mobjMsg.Abrir "Digite o número da OS que deseja CLONAR", , informacao
            Exit Sub
        End If
        Text8.Text = Format(Text8.Text, "000000")
        clonarOS
    End Select
    'If Index = 0 Then Msgbox "Procedimento realizado com sucesso!", vbInformation, "ZEUS"
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

Private Sub salvar_dados_hist()
On Error GoTo Err
    Dim rsGravaHist As New ADODB.Recordset
    Dim sqlGravaHist As String
    
    Dim rsExcHist As New ADODB.Recordset
    Dim sqlExcHist As String
    
    sqlGravaHist = "Select * from tbMPHist"
    rsGravaHist.Open sqlGravaHist, cnBanco, adOpenKeyset, adLockOptimistic
    
    ListView4.ListItems.Item(1).Selected = True
    
    sqlExcHist = "Delete from tbMPHist where programacao = '" & Val(txtformula(11)) & "' and seqprog = '" & Val(ListView4.SelectedItem.ListSubItems.Item(5)) & "'"
    rsExcHist.Open sqlExcHist, cnBanco
    
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        ListView4.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaHist.AddNew
        rsGravaHist(0) = Val(ListView4.ListItems.Item(X))
        rsGravaHist(1) = ListView4.SelectedItem.ListSubItems.Item(1)
        rsGravaHist(2) = ListView4.SelectedItem.ListSubItems.Item(2)
        rsGravaHist(3) = ListView4.SelectedItem.ListSubItems.Item(3)
        rsGravaHist(4) = Val(txtformula(11))
        rsGravaHist(5) = ListView4.SelectedItem.ListSubItems.Item(5)
        rsGravaHist(6) = ListView4.SelectedItem.ListSubItems.Item(6)
        rsGravaHist(7) = ListView4.SelectedItem.ListSubItems.Item(7)
    Next
    If Not rsGravaHist.EOF Then rsGravaHist.Update
    rsGravaHist.Close
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

Private Function salvar_Dados()
On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
        
    salvar_Dados = True
    
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(1, 1) = txtformula(11).Text
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = DTPicker1.Value
    vQualquerDado(2, 2) = "D"
    vQualquerDado(3, 1) = txtformula(16).Text
    vQualquerDado(3, 2) = "I"
    vQualquerDado(4, 1) = txtformula(14).Text
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = "S"
    vQualquerDado(5, 2) = "S"
    vQualquerDado(6, 1) = ""
    vQualquerDado(6, 2) = ""
    vQualquerDado(7, 1) = Text10.Text
    vQualquerDado(7, 2) = "S"
    
    GravaDados "tbMP", "idprogramacao", "I", txtformula(11), 7, "", "", txtformula(11)
        
    'TROCAR AS FUNÇÕES ABAIXO PELAS OPERAÇÕES PADRAO
    GravaLV1
    'Grava dados ListView1
    'limpaQualquerDado
    'ordenaLVArray ListView1, "8", "0", "2", "3", "4", "5", "6", "7", "9", "1", "10", "11", "12", "13", "", ""
    'GravaDadosLV "tbMPItens", "idprogramacao", "I", txtformula(11)
    
    
    Dim rsGravaDataProgramacao As New ADODB.Recordset
    Dim sqlGravaDataProgramacao As String
    If DatePart("ww", CDate(DTPicker1.Value), vbMonday, vbFirstFourDays) = DatePart("ww", CDate(Date), vbMonday, vbFirstFourDays) Then
        sqlGravaDataProgramacao = "update tbMPItens set dataprogramacao = '" & Format(DTPicker1.Value, "yyyy-mm-dd") & "' where idprogramacao = '" & Val(txtformula(11).Text) & "'"
    Else
        sqlGravaDataProgramacao = "update tbMPItens set dataprogramacao = '" & Format(Date, "yyyy-mm-dd") & "' where idprogramacao = '" & Val(txtformula(11).Text) & "' and dataprogramacao is null and dataprevista is null"
    End If
    rsGravaDataProgramacao.Open sqlGravaDataProgramacao, cnBanco
        
        
    'Grava dados ListView3
    limpaQualquerDado
    ordenaLVArray ListView3, "0", "4", "2", "3", "", "", "", "", "", "", "", "", "", "", "", ""
    GravaDadosLV "tbServTercOS", "idos", "I", txtformula(17)
        
    If txtformula(17).Text <> "" Then
        'Limpa dados da Matriz vQualquerDado
        limpaQualquerDado
        'Grava dados do OS na tabela tbOS
        'O 1º parametro é o valor que sera gravado no campo
        'O 2º parametro é o tipo de dado que o campo armazena
            
        'O STATUS da OS não pode ser alterado toda vez que ela sofrer uma alteração
        'A rotina abaixo localiza a OS, grava o status da OS em uma variavel para que possa
        'preservar o status atual
        Dim rsOSstatus As New ADODB.Recordset
        Dim slqOSstatus As String
        Dim statusOS As Integer
        sqlOSstatus = "SELECT IDOS,RASTREABILIDADE,OBSERVACAO,DATAOS,REVISAO,STATUS,TIPOOS FROM TBOS as a where a.idos = '" & Val(txtformula(17).Text) & "' and revisao = '" & Val(txtformula(18).Text) & "'"
        rsOSstatus.Open sqlOSstatus, cnBanco, adOpenKeyset, adLockReadOnly
        If Not rsOSstatus.EOF Then
            statusOS = rsOSstatus.Fields(5)
        End If
        '------------------------------------------------------------------
            
        vQualquerDado(1, 1) = txtformula(17).Text
        vQualquerDado(1, 2) = "I"
        vQualquerDado(2, 1) = txtformula(19).Text
        vQualquerDado(2, 2) = "S"
        vQualquerDado(3, 1) = txtformula(20).Text
        vQualquerDado(3, 2) = "S"
        vQualquerDado(4, 1) = DTPicker3.Value
        vQualquerDado(4, 2) = "D"
        If txtformula(18).Text = "" Then
            vQualquerDado(5, 1) = 0
        Else
            vQualquerDado(5, 1) = txtformula(18).Text
        End If
        vQualquerDado(5, 2) = "S"
        vQualquerDado(6, 1) = statusOS
        vQualquerDado(6, 2) = "I"
        
        If Combo2.Text = "Fabricação" Then vQualquerDado(7, 1) = 0
        If Combo2.Text = "Manutenção" Then vQualquerDado(7, 1) = 1
        If Combo2.Text = "Usinagem" Then vQualquerDado(7, 1) = 2
        vQualquerDado(7, 2) = "I"
        GravaDados "tbOS", "idos", "I", txtformula(17), 7, "", "", txtformula(17)
        sqlDeletar = "delete from tbositens where idos = '" & Val(txtformula(17).Text) & "' and revisao = '" & txtformula(18).Text & "'"
        rsDeletar.Open sqlDeletar, cnBanco
        gravaItensOS
    End If
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        salvar_Dados = False
        Msgbox Err.Number & " - " & Err.Description
        Exit Function
    End If
End Function

Private Function AgregarOS()
    AgregarOS = False
    If ListView1.ListItems.Count < 1 Then Exit Function
    Dim Y As Integer, X As Integer, vConta As Integer
    Y = ListView1.ListItems.Count
    vConta = 0
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            vConta = vConta + 1
            'Captura o número da OS a qual o item selecionado fará parte
            If vConta = 1 Then
                txtformula(25).Text = ListView1.SelectedItem.ListSubItems.Item(1)
            ElseIf vConta > 1 Then
                If Val(ListView1.SelectedItem.ListSubItems.Item(1)) <> 0 Then
                    mobjMsg.Abrir "Itens selecionado já em outra OS!", Ok, critico, "ZEUS"
                    'A linha abaixo foi adicionada para corrigir códigos de barra qdo forem gerados errado
                    'ListView1.SelectedItem.ListSubItems.Item(12) = Val(txtformula(11).Text) & Val(txtformula(25).Text) & Val(ListView1.ListItems.Item(X)) & Val(ListView1.SelectedItem.ListSubItems.Item(11))
                    Exit Function
                End If
                ListView1.SelectedItem.ListSubItems.Item(1) = txtformula(25).Text
                ListView1.SelectedItem.ListSubItems.Item(12) = Val(txtformula(11).Text) & Val(txtformula(25).Text) & Val(ListView1.ListItems.Item(X)) & Val(ListView1.SelectedItem.ListSubItems.Item(11))
                ListView1.SelectedItem.ListSubItems.Item(13) = "1"
                ListView1.SelectedItem.ListSubItems.Item(14) = Format(ListView1.SelectedItem.ListSubItems.Item(14), "000000000000")
                ListView1.ListItems.Item(X).Checked = False
            End If
        End If
    Next
    AgregarOS = True
End Function

Private Function MarcaOS()
    MarcaOS = False
    If ListView1.ListItems.Count < 1 Then Exit Function
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            If Val(ListView1.SelectedItem.ListSubItems.Item(1)) <> 0 Then
                mobjMsg.Abrir "Itens selecionado já em outra OS!", Ok, critico, "ZEUS"
                'A linha abaixo foi adicionada para corrigir códigos de barra qdo forem gerados errado
                'ListView1.SelectedItem.ListSubItems.Item(12) = Val(txtformula(11).Text) & Val(txtformula(25).Text) & Val(ListView1.ListItems.Item(X)) & Val(ListView1.SelectedItem.ListSubItems.Item(11))
                Exit Function
            End If
            ListView1.SelectedItem.ListSubItems.Item(1) = txtformula(25).Text & "/" & Val(txtformula(18))
            ListView1.SelectedItem.ListSubItems.Item(12) = Val(txtformula(11).Text) & Val(txtformula(25).Text) & Val(ListView1.ListItems.Item(X)) & Val(ListView1.SelectedItem.ListSubItems.Item(11))
            ListView1.SelectedItem.ListSubItems.Item(13) = "1"
            ListView1.ListItems.Item(X).Checked = False
        End If
    Next
    MarcaOS = True
End Function

Private Sub EditaLVHist()
On Error Resume Next
    Dim Y As Integer, X As Integer
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        If ListView4.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtformula(28).Text = ListView4.ListItems.Item(X)
    Me.txtformula(0).Text = ListView4.SelectedItem.ListSubItems.Item(6)
    Me.txtformula(4).Text = ListView4.SelectedItem.ListSubItems.Item(7)
    Me.txtformula(5).Text = ListView4.SelectedItem.ListSubItems.Item(1)
    Me.Label8.Caption = ListView4.SelectedItem.ListSubItems.Item(3)
End Sub

Private Sub EditaLVMP()
        LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(26), txtformula(2)
        If Text8.Text = "" Then
            AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), txtLV
        Else
            AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), vPonte2, vPonte2, txtResultado, vPonte3, txtformula(5), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), txtLV
        End If
        montaEstrutTreeview
        compoeDadosLVs
        txtformula(17).Text = Mid$(vPonte1.Text, 1, 9)
        If vPonte2.Text <> "-" And vPonte2.Text <> "009999" Then DTPicker2.Value = vPonte2.Text
        
        'Text9.Text = DatePart("ww", CDate(DTPicker2.Value))
        
        Label8.Caption = vPonte3.Text
        Combo1.Text = vPonte4.Text
        EditaTreeview
        CompoeControles
        separaDadosText1 Text1
        vPesoTotal2 = 0
        'Text1 = ""
        
        'Calcula o tempo automaticamente se vPAutomatico for igual a "S"
        'Caso contrário o procedimento requer entrada manual de parâmetros
        If vPAutomatico = "S" Then
            vAcumula = ""
            vAcumulaTempo = 0
            If Text1 <> "009999" Then
                mostraDesenhos "tbMPDesSel" & vTime, TreeView2
            End If
            txtResultado = Format(vAcumulaTempo, "#,##0.00;(#,##0.00)")
            'txtformula(5).Text = ""
            txtformula(5).Text = "AUTOMÁTICO"
        Else
            If Text1 <> "009999" Then
                mostraDesenhos "tbMPDesSel" & vTime, TreeView2
            End If
        End If
        '-------------------------------------------------------
        
        
        verOS 'Verifica se há alguma OS ativa
        compoeControlesOS
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"

'---------------------
        LimpaLV ListView3
        chamaSQL "select a.idservterc,b.nmserv,a.observacao,a.quantidade,a.idos from tbServTercOS as a inner join tbServTerc as b on a.idservterc = b.idservterc where a.idos = '" & Val(vPonte1.Text) & "'"
        Compoe_Listview ListView3, Sqlp, "00"
'---------------------
'---------------------
        If txtformula(11).Text <> "" Then
            LimpaLV ListView4
            chamaSQL "select a.* from tbMPHist as a where a.programacao = '" & Val(txtformula(11)) & "' and a.seqprog = '" & Val(txtformula(15)) & "'"
            Compoe_Listview ListView4, Sqlp, "00"
            SomaLV ListView4, 2, Text7
            pesoTempo
        End If
'---------------------
        calculaTempoApropriado ListView1.SelectedItem.ListSubItems.Item(12), ListView1.SelectedItem.ListSubItems.Item(6)
        txtformula(11) = Format(txtformula(11), "000000")
        
        'A ROTINA ABAIXO RODA CASO A OPERAÇÃO SELECIONADA NÃO POSSA SER MAIS ALTERADA DEVIDO A SEMANA PROGRAMADA
        If SkinLabel20.Visible = True Then
            'vPonte1.Text = "0"
            'txtformula(27) = "0"
            'LimpaVariaveis
            'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(26), txtformula(0)
            'LimpaControles txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1)
            'Combo1.Text = ""
            Dim vPos As Integer, r As Integer, S As Integer
            r = ListView1.ListItems.Count
            For S = 1 To r
                If ListView1.ListItems.Item(S).Selected = True Then vPos = S
            Next
            If Status <> "editar" Then
                txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
            End If
            ListView1.ListItems.Item(vPos).Selected = True
'            SomaLV ListView1, 6, Text2
            'TreeView3.Nodes.Clear
            'aicAlphaImage1.ClearImage
'            Exit Sub
        End If
End Sub

Private Sub verOS()
    'Temporariamente
    SSTab1.TabEnabled(2) = True
    Exit Sub
    
    If txtformula(17).Text = "" Or txtformula(17).Text = "0" Then
        SSTab1.TabEnabled(2) = False
    Else
        'verifica se a OS ja esta sendo apropriada. Se estiver o sistema não deixa editar
        '1 - Não houve apropriacao
        '2 - houve apropriação
        '3 - OS fechada
        If vStatus <= 1 Then
            SSTab1.TabEnabled(2) = True
        End If
    End If
End Sub

Private Function GeraOSLV(LV As Listview)
    If LV.ListItems.Count > 0 Then
        Dim X As Integer
        X = 1
        LV.Sorted = True
        LV.SortKey = 1
        LV.SortOrder = lvwDescending
        LV.ListItems.Item(X).Selected = True
        GeraOSLV = Val(LV.SelectedItem.ListSubItems.Item(1)) + 1
        LV.SortKey = 11
        LV.SortOrder = lvwAscending
        Exit Function
    Else
        GeraOSLV = 1
    End If
End Function

Private Sub gravaItensOS()
    If ListView1.ListItems.Count < 1 Then Exit Sub
    'Label36.Caption = "Alteração"
    Dim Y As Integer, Z As Integer
    Y = ListView1.ListItems.Count
    For Z = 1 To Y
        ListView1.ListItems.Item(Z).Selected = True
        If Val(ListView1.SelectedItem.ListSubItems.Item(1)) = Val(txtformula(17).Text) Then
            separaDesLv ListView1.SelectedItem.ListSubItems.Item(4)
        End If
    Next
End Sub

Private Sub separaDesLv(vTxtForm As String)
On Error GoTo Err
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Len(RECEBE) = 5 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 3)
            Else
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 4)
            End If
            SqlTransf = "Insert into tbOSItens(idos,revisao,fce,projeto,codlm,codseq,idcc,idprogramacao,status,codigobarra,idoperacao) Values('" & Val(txtformula(17)) & "','" & txtformula(18) & "','" & Val(txtformula(12)) & "','" & txtformula(13) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "','" & ListView1.SelectedItem.ListSubItems.Item(2) & "','" & Val(txtformula(11)) & "',1,'" & ListView1.SelectedItem.ListSubItems.Item(12) & "','" & ListView1.SelectedItem.ListSubItems.Item(11) & "')"
            rsTransf.Open SqlTransf, cnBanco
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Len(RECEBE) = 5 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
        Else
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 4)
        End If
        SqlTransf = "Insert into tbOSItens(idos,revisao,fce,projeto,codlm,codseq,idcc,idprogramacao,status,codigobarra,idoperacao) Values('" & Val(txtformula(17).Text) & "','" & txtformula(18) & "','" & Val(txtformula(12)) & "','" & Mid$(txtformula(13), 1, 50) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "','" & ListView1.SelectedItem.ListSubItems.Item(2) & "','" & Val(txtformula(11)) & "',1,'" & ListView1.SelectedItem.ListSubItems.Item(12) & "','" & ListView1.SelectedItem.ListSubItems.Item(11) & "')"
        rsTransf.Open SqlTransf, cnBanco
    End If
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

Private Sub CarregaFCE()
On Error GoTo Err
    Dim X As Integer
    sqlFCE = "Select a.*,b.status from tbprojetos as a inner join tbFCE as b on a.fce = b.fce where a.fce = '" & txtformula(12) & "' and b.status <> 1 order by a.fce"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFCE.EOF Then
        txtformula(12).Text = txtformula(12)
        mobjMsg.Abrir "FCE não cadastrada", Ok, critico, "Atenção"
    Else
        txtformula(12).Text = rsFCE.Fields(1)
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
    Sqlp = "Select a.fce,MAX(a.oc) from tbprojetos as a inner join tbFCE as b on a.fce = b.fce where b.status <> 1 group by a.FCE order by a.fce"
    procnom = "fce"
    campo = 0
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de FCE"
'    Pesquisa = frmMPCompleto.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "fce=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtformula(12).Text = rsLocal.Fields(0)
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
        Resume Next
    End If
End Sub

Private Sub CarregaProjeto()
On Error GoTo Err
    Dim X As Integer
    SqlProjeto = "Select * from tbprojetos where fce = '" & txtformula(12) & "' order by fce"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsProjeto.EOF Then rsProjeto.MoveFirst
    rsProjeto.Find "projeto=" & "'" & Me.txtformula(13) & "'"
    If rsProjeto.EOF Then
        txtformula(13).Text = txtformula(13)
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Projeto não cadastrado", Ok, critico, "Atenção"
        End If
    Else
        txtformula(13).Text = rsProjeto.Fields(2)
        txtformula(16).Text = rsProjeto.Fields(0)
        'txtDesenho(1).Text = Format(rsProjeto.Fields(0), "000000")
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
    Sqlp = "Select * from tbprojetos where fce = '" & txtformula(12) & "' order by fce,Projeto"
    procnom = "projeto"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Projetos"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "projeto=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtformula(13).Text = rsLocal.Fields(2)
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

Private Sub Command1_Click()
    vFCE = txtformula(12).Text
    varGlobal = Val(txtformula(11))
    FCROrdemServico.Show 1
End Sub

Private Sub Form_Activate()
    'vTime = Time
    'vTime = RemoveMask(vTime)
    excluiTabela
    criaTabela
    verOS 'Verifica se há alguma OS ativa
    listview_cabecalho
    Status = Pesquisa
    If Status = "novo" Then
        txtformula(11) = ""
    ElseIf Status = "editar" Then
        ResultPesq
    End If
    ListView4.ListItems.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    Set vPonte3 = Me.Controls.Add("VB.TextBox", "vPonte3")
    Set vPonte4 = Me.Controls.Add("VB.TextBox", "vPonte4")
    Set vPonte5 = Me.Controls.Add("VB.TextBox", "vPonte5")
    DTPicker1 = Date
    DTPicker2.Value = Date
    SSTab1.Tab = 0

    verOS 'Verifica se há alguma OS ativa
    listview_cabecalho
    vPonte1.Text = ""
    Status = Pesquisa
    If Status = "novo" Then
        Text9.Text = DatePart("ww", (DTPicker2.Value), vbMonday, vbFirstFourDays)
        txtformula(11) = ""
    ElseIf Status = "editar" Then
        ResultPesq
        compoeControlesOS
        vPonte1.Text = ""
    End If
    
    ListView4.ListItems.Clear
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub mostraDesenhos(vTabela As String, TV As TreeView)
On Error GoTo Err
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    '17/04
    TV.Nodes.Clear

    If vTabela = "tbitemlm" Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(g.idos) as OS " & _
        "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbProjetos as f on f.codprojeto = c.codprojeto left join tbositens as g on a.fce = g.fce and a.codlm = g.codlm and a.codseq = g.codseq Where a.fce = '" & Val(txtformula(12)) & "' and f.projeto = '" & txtformula(13) & "'" & _
        "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    
    ElseIf vTabela = "tbMPDesSel" & vTime Then
    
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(h.idos) as OS " & _
        "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel" & RemoveMask(vTime) & " as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq inner join tbProjetos as g on g.codprojeto = c.codprojeto left join tbositens as h on a.fce = h.fce and a.codlm = h.codlm and a.codseq = h.codseq Where a.fce = '" & Val(txtformula(12)) & "'" & _
        "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    
    End If
    
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    If Text10.Text = "" Then Text10.Text = rsTreeview.Fields(0)
    
    vJuntaNome = ""
              vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & " - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
    separaDadosTree vJuntaNome
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TV.Nodes.Add(, , vNome1, vNome1)
        'TESTE DE COR --------------
        If Not IsNull(rsTreeview.Fields(13)) Then
            nd.ForeColor = &H8000&
        End If
        '----------------------------
        
        'If Mid$(vNome1, 1, 14) = "K2171-119-0207" Then
        '    Msgbox "Aki"
        'End If
        
        
        Do While Mid$(vNome1, 1, Len(vNome1) - 1) = Mid$(vNomeA, 1, Len(vNome1) - 1) And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TV.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                'TESTE DE COR --------------
                If Not IsNull(rsTreeview.Fields(13)) Then
                    nd.ForeColor = &H8000&
                End If
                '----------------------------
                
                
                'TEORICAMENTE INICIALIZAÇÃO DA VARIAVEL QUE RECEBERA O SOMATORIO DO VALOR DA POSIÇÃO DO DESENHO IRÁ
                'FICAR NESSE LOCAL
                vPesoPosicao = 0
                
                
                'Teste
                If Mid$(Right(vNome2, 14), 1, 3) = "OS:" Then
                    Dim vTamanho1 As Integer
                    vTamanho1 = Len(vNome2) - 11
                    vNome2 = Mid$(vNome2, 1, vTamanho1) & ")"
                End If
                
                
                
                
                Do While Mid$(vNome1, 1, Len(vNome1) - 1) = Mid$(vNomeA, 1, Len(vNome1) - 1) And Mid(vNome2, 1, Len(vNome2) - 1) = Mid$(vNomeB, 1, Len(vNome2) - 1) And vNomeC <> "" And Not rsTreeview.EOF
'                Do While vNome1 = vNomeA And vNome2 = vNomeB And vNomeC <> "" And Not rsTreeview.EOF
                
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    If TV.Name = "TreeView2" Then
                        
                        'Abaixo é calculado o peso de cada posicao de cada desenho e realizado a classificação
                        'dentro da formula
                        'vPesoPosicao = vPesoPosicao + (rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        vPesoPosicao = vPesoPosicao + (rsTreeview.Fields(6) * rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        
                        
                        vPesoTotal2 = vPesoTotal2 + (rsTreeview.Fields(6) * rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        
                        If Status = "novo" Then
                            If Text1.Text = "" Then
                                Text1 = Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                            Else
                                Text1 = Text1.Text & ";" & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                            End If
                        End If
                        
                        
                        
                    End If
                    Set nd = TV.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    'TESTE DE COR --------------
                    If Not IsNull(rsTreeview.Fields(13)) Then
                        nd.ForeColor = &H8000&
                    End If
                    '----------------------------
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
' ORIGINAL          vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                    vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & " - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                    separaDadosTree vJuntaNome
                    vPula = 1
                Loop
                
                'Utiliza o peso calculado da posição para classificar o tipo de estrutura e calcular o tempo
                'em seguida acumula o tempo encontrado para determinar o tempo real de fabricação
                If vPAutomatico = "S" Then
                    txtformula(5) = Format(vPesoPosicao, "#,##0.00;(#,##0.00)")
                    txtformula_KeyDown 5, 13, 5
                    vAcumulaTempo = vAcumulaTempo + txtResultado
                End If
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
'ORIGINAL       vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & ";" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                separaDadosTree vJuntaNome
            End If
            vPula = 0
            
            If Not rsTreeview.EOF Then
                vNome2 = vNomeB
            End If
        Loop
        If Not rsTreeview.EOF Then
            vNome1 = vNomeA
            vNome2 = vNomeB
            vNome3 = vNomeC
        End If
    Loop
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

Private Sub Form_Unload(Cancel As Integer)
    excluiTabela
    Set frmMPCompleto = Nothing
    
End Sub

Private Sub ListView1_Click()
    If txtformula(13).Text = "" Then
        mobjMsg.Abrir "Antes de selecionar a operação identifique a FCE/Projeto", Ok, critico, "Atenção"
        Exit Sub
    End If
    Status = "editar"
    EditaLVMP
End Sub

Private Sub ListView1_DblClick()
    Status = "editar"
    EditaLVMP
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Status = "editar"
    EditaLVMP
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    Status = "editar"
    EditaLVMP
End Sub

Private Sub ListView3_DblClick()
    AlteraLV ListView3, txtformula(21), txtformula(22), txtformula(23), txtformula(24), txtformula(17), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)
End Sub

Private Sub ListView4_Click()
    EditaLVHist
    CompoeHist
End Sub

Private Sub Text8_LostFocus()
    Text8.Text = Format(Text8.Text, "000000")
End Sub

Private Sub Text9_LostFocus()
    ' CHAMA FUNÇÃO QUE CONVERTE SEMANA DO ANO PARA DATA
    DTPicker2.Value = ""
    converteSemana Val(Text9.Text), DTPicker2, ""
    If DTPicker2.Value = "" Then
        mobjMsg.Abrir "Semana não encontrada", Ok, critico, "ZEUS"
        Exit Sub
    End If
End Sub

Private Sub Timer1_Timer()
    txtResultado = Text7
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Msgbox "ok"
    'compoeDadosLVs
    Timer2.Enabled = False
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As MSComctlLib.Node
    Set aux = Node.Child
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        If Not aux.Child Is Nothing Then
            TreeView1_NodeCheck aux
        End If
        Set aux = aux.Next
    Loop
    Set aux = Node.Parent
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        Set aux = aux.Parent
    Loop
    
    vAcumula = ""
    Label6 = "-"
    vPesoTotal1 = 0
    buscaChecado
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As MSComctlLib.Node
    Set aux = Node.Child
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        If Not aux.Child Is Nothing Then
            TreeView2_NodeCheck aux
        End If
        Set aux = aux.Next
    Loop
    Set aux = Node.Parent
    Do While Not aux Is Nothing
        aux.Checked = Node.Checked
        Set aux = aux.Parent
    Loop
    'vPesoTotal = 0
End Sub

Private Sub buscaChecado()
    Dim X As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    X = 0
    vQtdNos = TreeView1.Nodes.Count
    For X = 1 To vQtdNos
        If TreeView1.Nodes.Item(X).Checked = True Then
            PegaTreeview X
            separaDadosTree vJuntaNome
            buscaPeso
        End If
    Next
End Sub

Private Sub buscaChecado2(vLV As TreeView)
    Dim X As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    X = 0
    vQtdNos = vLV.Nodes.Count
    For X = 1 To vQtdNos
        If vLV.Nodes.Item(X).Checked = True Then
            transfDesenhosSel X, vLV
        End If
    Next
End Sub

Private Sub PegaTreeview(llng_Contador As Integer)
    If TreeView1.Nodes(llng_Contador).Checked = True Then
        vNmNo = TreeView1.Nodes(llng_Contador).FullPath
    End If
    vNmNo = Replace(vNmNo, "\", ";")
    vJuntaNome = vNmNo
End Sub

Private Sub buscaPeso()
On Error GoTo Err
    Dim rsBuscaPeso As New ADODB.Recordset
    Dim SqlBuscaPeso As String
    Dim vCodLM As String, vCodSeq As String
        
    If vNomeC <> "" Then
        If Mid$(Right(vNomeC, 6), 1, 1) = " " Then
            vNomeC = Right(vNomeC, 5)
            vCodLM = Mid$(vNomeC, 1, 2)
            vCodSeq = Mid$(vNomeC, 3, 3)
        Else
            vNomeC = Right(vNomeC, 6)
            vCodLM = Mid$(vNomeC, 1, 2)
            vCodSeq = Mid$(vNomeC, 3, 4)
        End If
        
        
        If vAcumula = vNomeC And Label6 <> "-" Then
            Exit Sub
        Else
            vAcumula = vNomeC
        End If
        
        SqlBuscaPeso = "select a.quantcj*a.quantunit*a.pesounit as PesoTotal from tbItemLM as a where a.fce = '" & Val(txtformula(12)) & "' and a.codlm = '" & Val(vCodLM) & "' and a.codseq = '" & Val(vCodSeq) & "'"
        rsBuscaPeso.Open SqlBuscaPeso, cnBanco, adOpenKeyset, adLockReadOnly
        vPesoTotal1 = vPesoTotal1 + rsBuscaPeso.Fields(0)
    End If
    Label6 = Format(vPesoTotal1, "#,##0.00;(#,##0.00)")
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

'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Seq.", ListView1.Width / 26
    ListView1.ColumnHeaders.Add , , "OS nº/Rev", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "ID. C.Custo", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome C. Custo", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Desenhos/Itens", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Data Prevista", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "T. Calculado", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Grupo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "ID Programação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Variáveis", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Observação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Operação", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Código de Barras", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Ordenação", ListView1.Width / 10000
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Valor constante", ListView2.Width / 2.5
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 3.5
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "ID", ListView3.Width / 6
    ListView3.ColumnHeaders.Add , , "Serviço", ListView3.Width / 2.5
    ListView3.ColumnHeaders.Add , , "Descrição", ListView3.Width / 3.5
    ListView3.ColumnHeaders.Add , , "Qtd.", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "OS", ListView3.Width / 10000
    
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "ID", ListView4.Width / 6
    ListView4.ColumnHeaders.Add , , "Variáveis", ListView4.Width / 2.5
    ListView4.ColumnHeaders.Add , , "Tempo", ListView4.Width / 2.5
    ListView4.ColumnHeaders.Add , , "Grupo", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Programação", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Seq", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "codreduzido", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "idFormula", ListView4.Width / 10000
    Me.ListView4.ColumnHeaders(3).Alignment = lvwColumnRight
    
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    
End Sub

Private Sub compoeAutomatico()
On Error GoTo Err
    Dim rscompoeAutomatico As New ADODB.Recordset
    Dim SqlcompoeAutomatico As String
    SqlcompoeAutomatico = "Select * from tbParametrosAut as a where a.codreduzido = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    rscompoeAutomatico.Open SqlcompoeAutomatico, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rscompoeAutomatico.EOF Then
        vPAutomatico = "S"
        txtformula(5).Enabled = False
    Else
        vPAutomatico = "N"
        txtformula(5).Enabled = True
    End If
    rscompoeAutomatico.Close
    Set rscompoeAutomatico = Nothing
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

Private Sub CompoeControles()
On Error GoTo Err
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    sqlCompoe = "Select a.parametros,a.formula,a.observacao,a.imagem from tbFormula as a where a.codreduzido = '" & txtformula(0) & "' and a.idform = '" & Val(txtformula(4)) & "'"
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        txtformula(2).Text = rsCompoe.Fields(0) 'Parâmetros
        txtformula(3).Text = rsCompoe.Fields(1) 'Formula
        If Not IsNull(rsCompoe.Fields(2)) Then txtformula(6).Text = rsCompoe.Fields(2) 'Observação
        If Not IsNull(rsCompoe.Fields(3)) Then Label53 = rsCompoe.Fields(3) Else Label53 = "-" 'Imagem
    Else
        txtformula(2).Text = "" 'Parâmetros
        txtformula(3).Text = "" 'Formula
        txtformula(6).Text = "" 'Observação
        Label53 = "-" 'Imagem
    End If
    If Mid$(txtformula(2).Text, 1, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 9, 1), 1
    End If
    If Mid$(txtformula(2).Text, 12, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 20, 1), 2
    End If
    
    separaDadosTree vNmNo
    If vNomeC <> "" Then
        Label8 = vNomeA & "/" & vNomeB & "/" & vNomeC
    ElseIf vNomeC = "" And vNomeB <> "" Then
        Label8 = vNomeA & "/" & vNomeB
    ElseIf vNomeB = "" Then
        Label8 = vNomeA
    End If
    
    aicAlphaImage1.ClearImage
    If Label53 <> "" Or Label53 <> "-" Then
        aicAlphaImage1.LoadImage_FromFile (Label53.Text)
    End If
    rsCompoe.Close
    Set rsCompoe = Nothing
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

Private Sub compoeDadosLVs()
    'Faz referências a Funções que estão no: Module1.bas
    'Listview2 - Constantes
    LimpaLV ListView2
    chamaSQL "Select a.idseq,a.valconst,a.descricao from tbconstantesCC as a where a.idprd = '" & txtformula(0) & "'"
    Compoe_Listview ListView2, Sqlp, "000"
End Sub

Private Sub LimpaVariaveis()
    vGrupo = ""
    vDimTipo = ""
    vDimValor = ""
    vInterTipo = ""
    vInterValor = ""
    vSomaTempo = 0
    vTMedio = 0
    vFFadiga = 0
    vOrganiza = 0
    vSomaTempo = 0
End Sub

'As 3 próximas SUBs são referentes a montagem e manipulação do TREEVIEW3
Private Sub montaEstrutTreeview()
On Error GoTo Err
    Dim rsTreeview As New ADODB.Recordset
    Dim SqlTreeview As String
    Dim vNome1 As String, vNome2 As String, vNome3 As String
    Dim nd As Node
    Dim vPula As Integer
    Dim vNo As Integer, vNo2 As Integer
    Dim vNomeNo As String
       
    TreeView3.Nodes.Clear

    SqlTreeview = "Select * from tbFormula as a where a.codreduzido = '" & txtformula(0) & "' order by a.codreduzido,a.nmform"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    separaDadosTree rsTreeview.Fields(2)
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TreeView3.Nodes.Add(, , vNome1, vNome1)
        Do While vNome1 = vNomeA And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TreeView3.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                Do While vNome2 = vNomeB And vNomeC <> "" And Not rsTreeview.EOF
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    Set nd = TreeView3.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    separaDadosTree rsTreeview.Fields(2)
                    vPula = 1
                Loop
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                separaDadosTree rsTreeview.Fields(2)
            End If
            vPula = 0
            
            If Not rsTreeview.EOF Then
                vNome2 = vNomeB
            End If
        Loop
        If Not rsTreeview.EOF Then
            vNome1 = vNomeA
            vNome2 = vNomeB
            vNome3 = vNomeC
        End If
    Loop
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

Private Sub separaDadosTree(vTxtForm As String)
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    vNomeA = ""
    vNomeB = ""
    vNomeC = ""
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Contador = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            Contador = Contador + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Contador = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

'A função abaixo pega os valores dos parâmetro informados no textBox e armazena em variáveis
'específicas para cada valor
Private Sub separaDadosPar(vTxtForm As TextBox)
    Dim RECEBE As String
    Dim Contador As Integer, vNum As Integer
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Contador = 0 And RECEBE <> "-" Then vGrupo = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 1 Then vDimTipo = RECEBE 'Variável vDimTipo receber o valor do segundo parâmetro
            If Contador = 2 Then vDimValor = RECEBE 'Variavel vDimTipo recebe o valor do terceiro parâmetro
            If Contador = 3 Then vInterTipo = RECEBE 'Variável vInterTipo recebe o valor do quarto parâmetro
            If Contador = 4 Then vInterValor = RECEBE 'Variável vInterValor recebe o valor do quinto parâmetro
            Contador = Contador + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If Contador = 0 And RECEBE <> "-" Then vGrupo = RECEBE
    If Contador = 1 Then vDimTipo = RECEBE
    If Contador = 2 Then vDimValor = RECEBE
    If Contador = 3 Then vInterTipo = RECEBE
    If Contador = 4 Then vInterValor = RECEBE
    
    If Mid$(vDimValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vDimValor = var(Val(Mid$(vDimValor, 5, 2)))
        vDimValor = Replace(vDimValor, ",", ".")
    End If
    If Mid$(vInterValor, 1, 3) = "var" Then
        vNum = Val(Mid$(vDimValor, 5, 2))
        vInterValor = var(Val(Mid$(vInterValor, 5, 2)))
        vInterValor = Replace(vInterValor, ",", ".")
    End If
End Sub

'A função abaixo pega os valores das variáveis informados no textBox txtformula(5) e armazena em Arrays: var(?)
'específicas para cada valor
Private Sub separaDadosVar(vTxtForm As TextBox)
    On Error GoTo Err
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            If Contador = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            Contador = Contador + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Contador = 0 Then var(1) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 1 Then var(2) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 2 Then var(3) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 3 Then var(4) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 4 Then var(5) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 5 Then var(6) = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
    Exit Sub
Err:
    'mobjMsg.Abrir "Por favor refaça a operação", Ok, critico, "Atenção"
End Sub

'A função abaixo pega os valores das constantes informados no Listview2 e armazena em Arrays: cons(?)
'específicas para cada valor
Private Sub separaDadosCons()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        If ListView2.ListItems.Item(X).Selected = True Then
            cons(Val(ListView2.ListItems.Item(X))) = ListView2.SelectedItem.ListSubItems.Item(1)
        End If
    Next
End Sub

'A função abaixo separa os valores do texbox TEXT1 e grava na tabela tbMPDesSel
Private Sub separaDadosText1(vTxtForm As TextBox)
On Error GoTo Err
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim vCodLM As String, vCodSeq As String
    
    SqlTransf = "Delete from tbMPDesSel" & vTime & " where fce = '" & Val(txtformula(12)) & "'"
    rsTransf.Open SqlTransf, cnBanco
    
    Dim RECEBE As String
    Dim Contador As Integer, X As Integer
    Contador = 0
    For X = 1 To Len(vTxtForm)
        If Mid(vTxtForm, X, 1) = ";" Then
            'Separa para localizar: codigo da LM e código da sequência da LM
            'Se a variável recebe tiver + de 5 caracteres significa que a sequencia da LM ultrapassou a 999 registros
            'O procedimento para esse caso é diferenciado, por isso utilizasse o IF abaixo
            If Len(RECEBE) = 5 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 3)
            ElseIf Len(RECEBE) = 6 Then
                vCodLM = Mid$(RECEBE, 1, 2)
                vCodSeq = Mid$(RECEBE, 3, 4)
            End If
            
            SqlTransf = "Insert into tbMPDesSel" & vTime & "(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, X, 1)
        End If
    Next
    If RECEBE <> "" Then
        'Separa para localizar: codigo da LM e código da sequência da LM
        'Se a variável recebe tiver + de 5 caracteres significa que a sequencia da LM ultrapassou a 999 registros
        'O procedimento para esse caso é diferenciado, por isso utilizasse o IF abaixo
        If Len(RECEBE) = 5 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 3)
        ElseIf Len(RECEBE) = 6 Then
            vCodLM = Mid$(RECEBE, 1, 2)
            vCodSeq = Mid$(RECEBE, 3, 4)
        End If
        SqlTransf = "Insert into tbMPDesSel" & vTime & "(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
        rsTransf.Open SqlTransf, cnBanco
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

'Localiza a classificação na tabela baseado nos dados capturados na função separaDados
Private Sub localizaClassificacao()
On Error GoTo Err
    Dim rsLocaliza As New ADODB.Recordset
    Dim SqlLocaliza As String
    If vInterValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2 AND '" & vInterValor & "' BETWEEN inter1 and inter2"
    End If
    If vInterValor = "" And vDimValor <> "" Then
        SqlLocaliza = "select * from tbClassificacao where idprd = '" & txtformula(0) & "' and idgrupo = '" & Val(vGrupo) & "' and '" & vDimValor & "' BETWEEN dim1 and dim2"
    End If
    
    If SqlLocaliza <> "" Then
        rsLocaliza.Open SqlLocaliza, cnBanco, adOpenKeyset, adLockReadOnly
        If Not rsLocaliza.EOF Then
            vTMedio = rsLocaliza.Fields(7)
            vFFadiga = rsLocaliza.Fields(8)
            vOrganiza = rsLocaliza.Fields(9)
            vSomaTempo = vSomaTempo + (var(2) / vTMedio)
            rsLocaliza.Close
            Set rsLocaliza = Nothing
        End If
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

Private Sub AlteraTreeview()
On Error GoTo Err
    Dim rsAlteraTreeview As New ADODB.Recordset
    Dim SqlAlteraTreeview As String
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView3.Nodes.Count
        If TreeView3.Nodes(llng_Contador).Selected = True Then
            vNmNo = TreeView3.Nodes(llng_Contador).FullPath
        End If
    Next
    vNmNo = Replace(vNmNo, "\", ";")
    SqlAlteraTreeview = "Select idform,nmform,formula,parametros from tbFormula where nmform = '" & vNmNo & "'"
    rsAlteraTreeview.Open SqlAlteraTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsAlteraTreeview.EOF Then txtformula(4) = rsAlteraTreeview.Fields(0)
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

Private Sub EditaTreeview()
On Error GoTo Err
    Dim rsEditaTreeview As New ADODB.Recordset
    Dim SqlEditaTreeview As String
    vNmNo = Label8
    vNmNo = Replace(vNmNo, "/", ";")
    SqlEditaTreeview = "Select idform,nmform,formula,parametros from tbFormula where nmform = '" & vNmNo & "'"
    rsEditaTreeview.Open SqlEditaTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsEditaTreeview.EOF Then txtformula(4) = rsEditaTreeview.Fields(0)
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

Private Sub TreeView3_Click()
    'Em TESTE
    'ListView4.ListItems.Clear
    AlteraTreeview
    LimpaVariaveis
    LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    compoeDadosLVs
    CompoeControles
    compoeAutomatico
    If vPAutomatico = "S" Then
        txtformula(5).Enabled = False
        If Text1.Text <> "" Then
            vAcumula = ""
            vAcumulaTempo = 0
            Status = "editar"
            'Text1 = ""
            mostraDesenhos "tbMPDesSel" & vTime, TreeView2
            txtResultado = Format(vAcumulaTempo, "#,##0.00;(#,##0.00)")
        Else
            Msgbox "Nenhum DESENHO selecionado na guia de Desenhos"
        End If
        'txtformula(5).Text = ""
        txtformula(5).Text = "AUTOMÁTICO"
    Else
        txtformula(5).Enabled = True
    End If
End Sub

Private Sub somaFadiga()
On Error GoTo Err
    Dim rsFadiga As New ADODB.Recordset
    Dim SqlFadiga As String
    Dim vFadiga As Integer, vSetup As Integer
    SqlFadiga = "select * from tbConstantesCC where idprd = '" & txtformula(0).Text & "'"
    rsFadiga.Open SqlFadiga, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsFadiga.EOF Then
        vFadiga = rsFadiga.Fields(2)
        rsFadiga.MoveNext
        vSetup = rsFadiga.Fields(2)
        '1º soma 20% de fadiga
        txtResultado = txtResultado + (txtResultado * vFadiga / 100) + vSetup
        'vAcumulaTempo = vAcumulaTempo + vSetup
    End If
    rsFadiga.Close
    Set rsFadiga = Nothing
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

Private Sub txtformula_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtformula(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtformula.Count - 1
        txtformula(X).SelStart = 0
        txtformula(X).SelLength = Len(txtformula(X).Text)
    Next
End Sub

Private Sub txtformula_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            If txtformula(0).Text = "" Then
                Msgbox "Selecione primeiro um CC - Centro de Custo"
                Exit Sub
            End If
            CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtformula(0), txtformula(1), 7, 2, txtformula(0), "S", txtformula(1), "1"
            montaEstrutTreeview
            LimpaVariaveis
            LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
            compoeDadosLVs
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            preparaDados
            txtResultado = ""
            calculaValores 1
            
            'RECURSO IMPLEMENTADO PARA CALCULO AUTOMATICO P/ ADICIONAR 20% DE FADIGA E 30 MINUTOS DE SETUP
            'PARA NÃO UTILIZAR O CALCULO DESABILITE AS 3 LINHA ABAIXO
            If vPAutomatico = "S" Then
                somaFadiga
            End If
            '--------------------------------------------------------------------------------------------
            
            
            IncluiHistorico
            SomaLV ListView4, 2, Text7
            pesoTempo
            Timer1.Enabled = True
            
        End If
    Case 12
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            CarregaFCE
        End If
    Case 13
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtformula(12) <> "" Then
                CarregaProjeto
                mostraDesenhos "tbitemlm", TreeView1
                txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
            Else
                mobjMsg.Abrir "FCE não informada", Ok, critico, "Atenção"
                txtformula(13) = ""
            End If
        End If
    Case 17
    Case 21
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTxt "tbservterc", "idservterc", "I", "", "", txtformula(21), txtformula(21), 0, 1, txtformula(21), "I", txtformula(22), "1"
        End If
    End Select
End Sub

Private Sub IncluiHistorico()
    vPonte3.Text = Label8.Caption
    If txtformula(28) = "" Then
        txtformula(28) = Format(GeraCodigoLV(ListView4), "00")
    End If
    IncluirLV ListView4, txtformula(28), txtformula(5), txtResultado, vPonte3, txtformula(11), txtformula(15), txtformula(0), txtformula(4), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28)
    txtformula(28) = ""
End Sub

Private Sub CompoeHist()
On Error GoTo Err
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    sqlCompoe = "Select a.parametros,a.formula,a.observacao,a.imagem from tbFormula as a where a.codreduzido = '" & txtformula(0) & "' and a.idform = '" & Val(txtformula(4)) & "'"
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        txtformula(2).Text = rsCompoe.Fields(0) 'Parâmetros
        txtformula(3).Text = rsCompoe.Fields(1) 'Formula
        If Not IsNull(rsCompoe.Fields(2)) Then txtformula(6).Text = rsCompoe.Fields(2) 'Observação
        If Not IsNull(rsCompoe.Fields(3)) Then Label53 = rsCompoe.Fields(3) Else Label53 = "-" 'Imagem
    Else
        txtformula(2).Text = "" 'Parâmetros
        txtformula(3).Text = "" 'Formula
        txtformula(6).Text = "" 'Observação
        Label53 = "-" 'Imagem
    End If
    If Mid$(txtformula(2).Text, 1, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 9, 1), 1
    End If
    If Mid$(txtformula(2).Text, 12, 7) = "formula" Then
        localizaFormula Mid$(txtformula(2).Text, 20, 1), 2
    End If
    
    'separaDadosTree vNmNo
    'If vNomeC <> "" Then
    '    Label8 = vNomeA & "/" & vNomeB & "/" & vNomeC
    'ElseIf vNomeC = "" And vNomeB <> "" Then
    '    Label8 = vNomeA & "/" & vNomeB
    'ElseIf vNomeB = "" Then
    '    Label8 = vNomeA
    'End If
    
    aicAlphaImage1.ClearImage
    If Label53 <> "" Or Label53 <> "-" Then
        aicAlphaImage1.LoadImage_FromFile (Label53.Text)
    End If
    'compoeAutomatico
    rsCompoe.Close
    Set rsCompoe = Nothing
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

'vPosicao indica a posicao da formula
Private Sub localizaFormula(vNForm As Integer, vPosicao As Integer)
    Dim rsFormula As New ADODB.Recordset
    Dim SqlFormula As String
    SqlFormula = "select * from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & vNForm & "'"
    rsFormula.Open SqlFormula, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsFormula.EOF Then
        If vPosicao = 1 Then
            txtformula(7).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(8).Text = rsFormula.Fields(3) 'Parametros 2
        ElseIf vPosicao = 2 Then
            txtformula(10).Text = rsFormula.Fields(4) 'Formula 2
            txtformula(9).Text = rsFormula.Fields(3) 'Parametros 2
        End If
    End If
    rsFormula.Close
    Set rsFormula = Nothing
End Sub

Private Sub substituiValores(vFormula As TextBox)
    Dim X As Integer
    Dim vPreserva As String
    vPreserva = ""
    vPreserva = vFormula
    For X = 1 To 50
        vFormula = Replace(vFormula, "cons(" & (X) & ")", cons(X))
        vFormula = Replace(vFormula, "var(" & (X) & ")", var(X))
        vFormula = Replace(vFormula, "vTMedio", vTMedio)
        vFormula = Replace(vFormula, "vFFadiga", vFFadiga)
        vFormula = Replace(vFormula, "vOrganiza", vOrganiza)
    Next
    vFormula = Replace(vFormula, ",", ".")
    txtDecoder = vFormula
    vFormula = vPreserva
End Sub

Private Sub calculaValores(vQual As Integer)
On Error Resume Next
    'O ScriptControl é um componente. Ele interpreta e executa a formula/expressão numérica de um textbox
    If vQual = 1 Then
        txtResultado = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
        'SOMENTE PARA O CENTRO DE CUSTO SOLDA
        'CONVERTE O RESULTADO EM HORAS PARA MINUTOS
        
        If Mid$(txtformula(0).Text, 1, 12) = "3000.3104.SC" Then
            txtResultado = Format(txtResultado * 60, "#,##0.00;(#,##0.00)")
        End If
    Else
        vGrupo = "1"
        vDimValor = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
        vDimValor = Replace(vDimValor, ",", ".")
        vDimValor = Replace(vDimValor, "(", "")
        vDimValor = Replace(vDimValor, ")", "")
        'MsgBox vResultFormula
    End If
End Sub

Private Sub preparaDados()
    LimpaVariaveis
    If txtformula(5) = "" Then
        Msgbox "Favor informar o campo: " & txtformula(5).Tag, vbInformation, "Atenção"
        txtformula(5).SetFocus
        Exit Sub
    End If
    'Calcula as formulas carregadas a partir das funções abaixo carregadas
    'a partir dos dados informados no campo de variáveis
    If Mid$(txtformula(2).Text, 1, 7) <> "formula" Then
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
        localizaClassificacao
        substituiValores txtformula(3)
    Else
        If txtformula(7) <> "" Then
            'Acha o resultado referente a formula1
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(7)
            calculaValores 2
            localizaClassificacao
        End If
        
        If txtformula(10) <> "" Then
            'Acha o resultado referente a formula3
            separaDadosVar txtformula(5)
            separaDadosCons
            substituiValores txtformula(10)
            calculaValores 2
            localizaClassificacao
        End If
        
        vTMedio = Format(vSomaTempo, "#,##0.00;(#,##0.00)")
        'Pega o resultado das formulas 1 e 2 e aplica na formula3
        separaDadosVar txtformula(5)
        separaDadosPar txtformula(2)
        separaDadosCons
'        localizaClassificacao
        substituiValores txtformula(3)
    End If
End Sub

Private Sub transfDesenhosSel(llng_Contador As Integer, vTV As TreeView)
On Error GoTo Err
    Dim vNomeNo As String
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    
    
    If vTV.Nodes(llng_Contador).Checked = True Then
        vNomeNo = vTV.Nodes(llng_Contador).FullPath
    End If
    vNomeNo = Replace(vNomeNo, "\", ";")
    vJuntaNome = vNomeNo
    
    separaDadosTree vJuntaNome
    
    'vNomeC = right(vNomeC, 5)
    'vCodLM = Mid$(vNomeC, 1, 2)
    'vCodSeq = Mid$(vNomeC, 3, 3)
    
    If Mid$(Right(vNomeC, 6), 1, 1) = " " Then
        vNomeC = Right(vNomeC, 5)
        vCodLM = Mid$(vNomeC, 1, 2)
        vCodSeq = Mid$(vNomeC, 3, 3)
    Else
        vNomeC = Right(vNomeC, 6)
        vCodLM = Mid$(vNomeC, 1, 2)
        vCodSeq = Mid$(vNomeC, 3, 4)
    End If
10  cnBanco.BeginTrans
    
    If vAcumula = vNomeC And Label6 <> "-" Then
        cnBanco.CommitTrans
        Exit Sub
    Else
        vAcumula = vNomeC
    End If
    
    If vTV.Name = "TreeView1" Then
        If vCodLM <> "" And vCodSeq <> "" Then
            SqlTransf = "Insert into tbMPDesSel" & vTime & "(fce,codlm,codseq) Values('" & Val(txtformula(12)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
        End If
    ElseIf vTV.Name = "TreeView2" Then
        SqlTransf = "Delete from tbMPDesSel" & vTime & " where fce = '" & Val(txtformula(12)) & "' and codlm = '" & Val(vCodLM) & "' and codseq = '" & Val(vCodSeq) & "'"
        rsTransf.Open SqlTransf, cnBanco
    End If
    cnBanco.CommitTrans
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Resume Next
    End If
End Sub

Private Sub SomaLV(LV As Listview, vColunaLV As Integer, vTxtRetorno As TextBox)
    On Error Resume Next
    Dim X As Integer, Y As Integer, F As Integer
    Y = LV.ListItems.Count
    Dim somaTempo As Double
    somaTempo = 0
    For X = 1 To Y
        If LV.ListItems.Item(X).Selected = True Then F = X
    Next
    For X = 1 To Y
        LV.ListItems.Item(X).Selected = True
        'If Trim$(LV.SelectedItem.ListSubItems.Item(6)) <> " " Then
            somaTempo = somaTempo + LV.SelectedItem.ListSubItems.Item(vColunaLV)
        'End If
    Next
    If somaTempo <> 0 Then
        vTxtRetorno.Text = Format(somaTempo, "#,##00.00;(#,##0.00)")
        LV.ListItems.Item(F).Selected = True
    End If
End Sub

Private Sub txtformula_KeyPress(Index As Integer, KeyAscii As Integer)
'Substitui aspas simples por aspas duplas
    If KeyAscii = 39 Then
        KeyAscii = 34
    End If
End Sub

Private Sub txtformula_LostFocus(Index As Integer)
    voltaCorText txtformula(Index)
    Select Case Index
    Case 5
        preparaDados
        txtResultado = ""
        calculaValores 1
    Case 13
        If txtformula(12) <> "" Then
            CarregaProjeto
            mostraDesenhos "tbitemlm", TreeView1
            txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        Else
            mobjMsg.Abrir "FCE não informada", Ok, critico, "Atenção"
            txtformula(13) = ""
        End If
    End Select
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlProg = "Select a.idprogramacao,a.dataprogramacao,a.codprojeto,a.responsavel,a.ativo,b.fce,b.projeto,a.desenho from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto where a.idprogramacao = '" & Val(varGlobal) & "'"
    rsProg.Open SqlProg, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProg.RecordCount > 0 Then
        compoeControlesForm
        mostraDesenhos "tbitemlm", TreeView1
        compoeDadosLV
        SomaLV ListView1, 6, Text2
    Else
        'mobjMsg.Abrir "Programação não encontrada", Ok, critico, "Atenção"
    End If
    rsProg.Close
    Set rsProg = Nothing
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

Private Sub compoeControlesForm()
    txtformula(11).Text = Format(rsProg.Fields(0), "000000")
    DTPicker1.Value = rsProg.Fields(1)
    txtformula(16).Text = rsProg.Fields(2)
    txtformula(14).Text = rsProg.Fields(3)
    txtformula(12).Text = rsProg.Fields(5)
    txtformula(13).Text = rsProg.Fields(6)
    If Not IsNull(rsProg.Fields(7)) Then Text10.Text = rsProg.Fields(7)
End Sub

Private Sub compoeControlesOS()
On Error GoTo Err
    Dim rsCompoeOS As New ADODB.Recordset
    Dim SqlCompoeOS As String
    
    SqlCompoeOS = "SELECT IDOS,RASTREABILIDADE,OBSERVACAO,DATAOS,REVISAO,STATUS,TIPOOS FROM TBOS where idos = '" & Val(vPonte1) & "'"
    rsCompoeOS.Open SqlCompoeOS, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoeOS.EOF Then
        txtformula(17).Text = Format(rsCompoeOS.Fields(0), "000000000")
        If rsCompoeOS.Fields(4) = "" Then
            txtformula(18).Text = 0 'Revisão
        Else
            txtformula(18).Text = rsCompoeOS.Fields(4) 'Revisão
        End If
        txtformula(19).Text = rsCompoeOS.Fields(1) 'Rastreabilidade
        txtformula(20).Text = rsCompoeOS.Fields(2) 'Observação
        DTPicker3.Value = rsCompoeOS.Fields(3)   'Data da OS
        If IsNull(rsCompoeOS.Fields(6)) Then
            Combo2.Text = "Fabricação"
        Else
            If rsCompoeOS.Fields(6) = 0 Then Combo2.Text = "Fabricação"
            If rsCompoeOS.Fields(6) = 1 Then Combo2.Text = "Manutenção"
            If rsCompoeOS.Fields(6) = 2 Then Combo2.Text = "Usinagem"
            'Combo2.Text = rsCompoeOS.Fields(6) 'Tipo da OS (Fabricação, Manutenção, Usinagem)
        End If
    End If
    rsCompoeOS.Close
    Set rsCompoeOS = Nothing
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

Private Sub compoeDadosLV()
    LimpaLV ListView1
    chamaSQL "select a.idsequencia,RIGHT('000000000'+ CONVERT(VARCHAR,a.idos),9) + '/' + a.revisaoos,a.idcc,a.nomecc,a.desenhos,a.dataprevista,a.tempocalc,a.grupo,a.idprogramacao,a.variaveis,a.observacao,RIGHT('000'+ CONVERT(VARCHAR,a.idoperacao),3),a.codigobarra,a.status,Replicate ('0',9 - Len(Cast(a.idos as varchar))) + Cast(a.idos as varchar) + Replicate ('0',3 - Len(Cast(a.idoperacao as varchar))) + Cast(a.idoperacao as varchar)  as ordenação  from tbMPItens as a where a.idprogramacao = '" & Val(varGlobal) & "'"
    Compoe_Listview ListView1, Sqlp, "000"
    txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
    vPonte1.Text = Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9))
    
    'Converte a data prevista da OS em Semana do Ano
    'Em teste
    
    If ListView1.ListItems.Count <> 0 Then ListView1.ListItems.Item(1).Selected = True
    
    'DTPicker2.Value = ListView1.SelectedItem.ListSubItems.Item(5)
    If Not IsNull(DTPicker2.Value) Then
        Text9.Text = DatePart("ww", (DTPicker2.Value), vbMonday, vbFirstFourDays)
    End If
    'Em teste
    
    ListView1.Sorted = True
    ListView1.SortKey = 11
    ListView1.SortOrder = lvwAscending
    MudaCorLV1
    If vStatus > 1 Then
        'bloqueiaEdicao
    End If
End Sub

Private Sub clonarOS()
On Error GoTo Err
    Dim rsAchaProg As New ADODB.Recordset
    Dim SqlAchaProg As String
    
    Dim X As Integer, Y As Integer
    SSTab1.Tab = 1
    SqlAchaProg = "select a.idos,a.idprogramacao from tbMPItens as a where a.idos = '" & Val(Text8) & "'"
    rsAchaProg.Open SqlAchaProg, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAchaProg.RecordCount <= 0 Then
        mobjMsg.Abrir "Nº de OS não encontrado", Ok, critico, "Atenção"
        rsAchaProg.Close
        Set rsAchaProg = Nothing
        Exit Sub
    Else
        vClone = rsAchaProg.Fields(1)
    End If
    
    LimpaLV ListView1
    chamaSQL "select a.idsequencia,RIGHT('000000000'+ CONVERT(VARCHAR,a.idos),9) + '/' + a.revisaoos,a.idcc,a.nomecc,a.desenhos,'',a.tempocalc,a.grupo,a.idprogramacao,a.variaveis,a.observacao,a.idoperacao,a.codigobarra,1,Replicate ('0',9 - Len(Cast(a.idos as varchar))) + Cast(a.idos as varchar) + Replicate ('0',3 - Len(Cast(a.idoperacao as varchar))) + Cast(a.idoperacao as varchar)  as ordenação  from tbMPItens as a where a.idprogramacao = '" & Val(vClone) & "'"
    Compoe_Listview ListView1, Sqlp, "000"
    txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
    
    ListView1.Sorted = True
    ListView1.SortKey = 11
    ListView1.SortOrder = lvwAscending
    
    Y = ListView1.ListItems.Count
    vStatus = 1
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        ListView1.SelectedItem.ListSubItems.Item(1) = "0"
        ListView1.SelectedItem.ListSubItems.Item(4) = "009999"
        ListView1.SelectedItem.ListSubItems.Item(12) = ""
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

Private Sub ClonaHist()
On Error GoTo Err
    Dim rsEncontraPROG As New ADODB.Recordset
    Dim SqlEncontraPROG As String
    
    Dim rsDuplicaHist As New ADODB.Recordset
    Dim SqlDuplicaHist As String
    
    SqlEncontraPROG = "select max(idprogramacao) from tbMPItens where idos = '" & Val(Text8) & "'"
    rsEncontraPROG.Open SqlEncontraPROG, cnBanco, adOpenKeyset, adLockReadOnly
    
    
    SqlDuplicaHist = "Insert Into tbMPHist(idseqhist,variaveis,tempo,grupo,programacao,seqprog,codreduzido,idformula) " & _
                     "Select idseqhist,variaveis,tempo,grupo,'" & Val(txtformula(11)) & "',seqprog,codreduzido,idformula " & _
                     "From tbMPHist where programacao = '" & rsEncontraPROG.Fields(0) & "'"
    rsDuplicaHist.Open SqlDuplicaHist, cnBanco
    
    rsDuplicaHist.Close
    Set rsDuplicaHist = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        'Msgbox Err.Number & " - " & Err.Description
        Resume Next
    End If
End Sub

Private Sub MudaCorLV1()
    'On Error Resume Next
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    vStatus = 1
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        'verifica se a OS ja esta sendo apropriada. Se estiver o sistema não deixa editar
        '1 - Não houve apropriacao
        '2 - houve apropriação
        '3 - OS fechada
        If ListView1.SelectedItem.ListSubItems.Item(13) = "" Then
            vStatus = Val(ListView1.SelectedItem.ListSubItems.Item(13))
        Else
            If ListView1.SelectedItem.ListSubItems.Item(13) > vStatus Then
                vStatus = Val(ListView1.SelectedItem.ListSubItems.Item(13))
            End If
            If ListView1.SelectedItem.ListSubItems.Item(13) = 2 Then
                ListView1.ListItems.Item(X).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(1).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(2).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(3).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(5).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(6).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(9).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(10).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(11).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(12).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(13).ForeColor = &H8000&
                ListView1.SelectedItem.ListSubItems.Item(14).ForeColor = &H8000&
            ElseIf ListView1.SelectedItem.ListSubItems.Item(13) = 3 Then
                ListView1.ListItems.Item(X).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(1).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(2).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(3).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(5).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(6).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(8).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(9).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(10).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(11).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(12).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(13).ForeColor = &H808080
                ListView1.SelectedItem.ListSubItems.Item(14).ForeColor = &H808080
            End If
        End If
    Next
End Sub

Private Sub bloqueiaEdicao()
    Dim X As Integer
    TreeView1.Enabled = False
    TreeView2.Enabled = False
    TreeView3.Enabled = False
    For X = 0 To cmdCadastro.Count - 1
        cmdCadastro(X).Enabled = False
    Next
    cmdCadastro(13).Enabled = True
    txtformula(0).Enabled = False
    txtformula(5).Enabled = False
    txtformula(12).Enabled = False
    txtformula(13).Enabled = False
    txtformula(26).Enabled = False
    Combo1.Enabled = False
    SSTab1.TabEnabled(2) = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    SkinLabel20.Visible = True
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView1, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As Listview, Column As ColumnHeader)
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

Private Sub criaTabela()
On Error GoTo Err
    cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMPDesSel" & vTime & "(" & _
    "fce NUMERIC NOT NULL," & _
    "codlm NUMERIC NOT NULL," & _
    "codseq NUMERIC NOT NULL)"
    'cnBanco.Close
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

Private Sub excluiTabela()
On Error GoTo Err
    Dim rsExcluirTb As New ADODB.Recordset
    Dim SqlExcluirTb As String
    SqlExcluirTb = "Drop table tbMPDesSel" & vTime
    rsExcluirTb.Open SqlExcluirTb, cnBanco
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        'Msgbox Err.Number & " - " & Err.Description
        Resume Next
    End If
End Sub

Private Sub calculaTempoApropriado(vCBarra As String, vTempoOrcado As String)
On Error GoTo Err
    Dim rsHAprop As New ADODB.Recordset
    Dim sqlHAprop As String
    Dim vHorasApropriadas As String
    Dim vTempoOrcadoConvertido As String
    
    sqlHAprop = "select CONVERT (VARCHAR, a.horasai-a.horaent, 108) as horaent,dbo.FN_CONVMIN(cast(replace(replace('" & vTempoOrcado & "','.',''),',','.') as money)) as Tempo_Convertido from tbOsMov  as a where a.codigobarra = '" & vCBarra & "'"
    rsHAprop.Open sqlHAprop, cnBanco, adOpenKeyset, adLockReadOnly
    vHorasApropriadas = "0000:00"
    If rsHAprop.RecordCount > 0 Then vTempoOrcadoConvertido = rsHAprop.Fields(1) Else vTempoOrcadoConvertido = "0000:00"
    Do While Not rsHAprop.EOF
        If Not IsNull(rsHAprop.Fields(0)) Then somaTempoPPSAtraso rsHAprop.Fields(0), vHorasApropriadas
        rsHAprop.MoveNext
    Loop
    rsHAprop.Close
    Set rsHAprop = Nothing
    SkinLabel26 = vHorasApropriadas
    SkinLabel27 = vTempoOrcadoConvertido
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

Private Function somaTempoPPSAtraso(vTempo, vOndeAcumula As String)
    Dim seg As Long, min As Long, hora As Long
    Dim tempo As Long
    Dim matriz2

    matriz2 = Split(vTempo, ":")
    tempo = tempo + (CLng(matriz2(0)) * 3600)
    tempo = tempo + (CLng(matriz2(1)) * 60)
    
    If vOndeAcumula <> "" Then
        matriz2 = Split(vOndeAcumula, ":")
        tempo = tempo + (CLng(matriz2(0)) * 3600)
        tempo = tempo + (CLng(matriz2(1)) * 60)
    End If
    
    hora = Int(tempo / 3600) ' aki são calculadas qtas horas
    tempo = tempo - (hora * 3600) 'aki subtraimos do tempo a qtde de segundos referentes as horas inteiras
    min = Int(tempo / 60) ' aki calculamos os minutos
    
    vOndeAcumula = Format(hora, "0000") & ":" & Format(min, "00")
    somaTempoPPSAtraso = vOndeAcumula
End Function

' A ROTINA ABAIXO NAO ESTA CONCLUIDA
' A IDEIA É SUBSTITUIR AS ROTINAS DE GRAVAÇÃO GENERICA DO LISTVIEW1, CASO VENHA APRESENTAR PROBLEMAS
Private Function GravaLV1()
On Error GoTo Err
    GravaLV1 = True
    Dim rsGuardaDataProg As New ADODB.Recordset
    Dim sqlGuardaDataProg As String
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsGravaLV As New ADODB.Recordset
    Dim SqlGravaLV As String
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer
    Dim vGuardaData(99, 3) As String
    
10  cnBanco.BeginTrans

    'Guarda dados referente a data programada da operação em uma matriz antes de limpar a tabela tbMPItens
    sqlGuardaDataProg = "select a.idprogramacao,a.idoperacao,a.dataprogramacao,a.databaixa from tbMPItens as a where a.idprogramacao = '" & Val(txtformula(11)) & "' order by a.idprogramacao,a.idoperacao"
    rsGuardaDataProg.Open sqlGuardaDataProg, cnBanco, adOpenKeyset, adLockReadOnly
    X = 0
    Do While Not rsGuardaDataProg.EOF
        vGuardaData(X, 0) = rsGuardaDataProg.Fields(0)
        vGuardaData(X, 1) = rsGuardaDataProg.Fields(1)
        If Not IsNull(rsGuardaDataProg.Fields(2)) Then vGuardaData(X, 2) = rsGuardaDataProg.Fields(2)
        If Not IsNull(rsGuardaDataProg.Fields(3)) Then vGuardaData(X, 3) = rsGuardaDataProg.Fields(3)
        rsGuardaDataProg.MoveNext
        X = X + 1
    Loop
    rsGuardaDataProg.Close
    Set rsGuardaDataProg = Nothing
    '------------------------------------------------------------------
    
    sqlDeletar = "Delete from tbMPItens where idprogramacao = '" & Val(txtformula(11)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(5) <> "-" And ListView1.SelectedItem.ListSubItems.Item(5) <> "" Then
            SqlGravaLV = "Insert into tbMPItens(" & _
                                "idprogramacao,idsequencia,idcc,nomecc,desenhos,dataprevista,tempocalc,grupo,variaveis,idos,observacao,idoperacao,codigobarra,status,revisaoos) " & _
                                "values(" & _
                                "'" & Val(txtformula(11)) & "', " & _
                                "'" & Val(ListView1.ListItems.Item(X)) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(2) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(3) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(4) & "', " & _
                                "'" & Format(ListView1.SelectedItem.ListSubItems.Item(5), "YYYY-MM-DD") & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(6) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(7) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(9) & "', " & _
                                "'" & Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(10) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(11) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(12) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(13) & "', " & _
                                "'" & Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 11, 1) & "')"
        Else
            SqlGravaLV = "Insert into tbMPItens(" & _
                                "idprogramacao,idsequencia,idcc,nomecc,desenhos,tempocalc,grupo,variaveis,idos,observacao,idoperacao,codigobarra,status,revisaoos) " & _
                                "values(" & _
                                "'" & Val(txtformula(11)) & "', " & _
                                "'" & Val(ListView1.ListItems.Item(X)) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(2) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(3) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(4) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(6) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(7) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(9) & "', " & _
                                "'" & Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(10) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(11) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(12) & "', " & _
                                "'" & ListView1.SelectedItem.ListSubItems.Item(13) & "', " & _
                                "'" & Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 11, 1) & "')"
        End If
        rsGravaLV.Open SqlGravaLV, cnBanco
    Next
    X = 0
    While vGuardaData(X, 0) <> ""
        If vGuardaData(X, 2) <> "" Then
            sqlGuardaDataProg = "update tbMPItens set dataprogramacao = '" & Format(vGuardaData(X, 2), "yyyy-mm-dd") & "' where idprogramacao = '" & Val(vGuardaData(X, 0)) & "' and idoperacao = '" & Val(vGuardaData(X, 1)) & "'"
            rsGuardaDataProg.Open sqlGuardaDataProg, cnBanco
        End If
        If vGuardaData(X, 3) <> "" Then
            sqlGuardaDataProg = "update tbMPItens set databaixa = '" & Format(vGuardaData(X, 3), "yyyy-mm-dd") & "' where idprogramacao = '" & Val(vGuardaData(X, 0)) & "' and idoperacao = '" & Val(vGuardaData(X, 1)) & "'"
            rsGuardaDataProg.Open sqlGuardaDataProg, cnBanco
        End If
        
        X = X + 1
    Wend
    For X = 0 To 49
        For Y = 0 To 3
            vGuardaData(X, Y) = ""
        Next
    Next X
    cnBanco.CommitTrans
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        Msgbox Err.Number & " - " & Err.Description
        Resume
    End If
End Function


Private Function limpaQualquerDado()
    Dim X As Integer, Y As Integer
'    For X = LBound(vQualquerDado) To UBound(vQualquerDado)
End Function

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(1).Enabled = False
        cmdCadastro(7).Enabled = False
        cmdCadastro(11).Enabled = False
    End If
    If vEdi = "N" Then
        'cmdCadastro(2).UseGreyscale = True
        'cmdCadastro(2).DragMode = 1
        'cmdCadastro(2).SpecialEffect = cbEngraved
    
        'cmdCadastro(9).UseGreyscale = True
        'cmdCadastro(9).DragMode = 1
        'cmdCadastro(9).SpecialEffect = cbEngraved
    
        'cmdCadastro(20).UseGreyscale = True
        'cmdCadastro(20).DragMode = 1
        'cmdCadastro(20).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(12).Enabled = False
    End If
    If vExc = "N" Then
        cmdCadastro(18).Enabled = False
        cmdCadastro(3).Enabled = False
        cmdCadastro(15).Enabled = False
    End If
    If vAva = "N" Then
'        chameleonButton1.UseGreyscale = True
'        chameleonButton1.DragMode = 1
'        chameleonButton1.SpecialEffect = cbEngraved
    End If
    'If vIntegra = "S" Then SSTab1.TabEnabled(6) = True Else SSTab1.TabEnabled(6) = False
End Sub

Private Sub pesoTempo()
    On Error Resume Next
    Dim rsConstante As New ADODB.Recordset
    Dim sqlConstante As String
    Dim vPesoConvertido As Double
    vPesoConvertido = Text7 / 60
    
    If InStr(UCase(Label8), UCase("AUTOMÁTICA")) > 0 Then
        sqlConstante = "select valconst from tbConstantesCC where idseq =" & 14
    ElseIf InStr(UCase(Label8), UCase("MIG-MAG")) > 0 Then
        sqlConstante = "select valconst from tbConstantesCC where idseq =" & 13
    ElseIf InStr(UCase(Label8), UCase("ELETRODO")) > 0 Then
        sqlConstante = "select valconst from tbConstantesCC where idseq =" & 12
    ElseIf InStr(UCase(Label8), UCase("TUBULAR")) > 0 Then
        sqlConstante = "select valconst from tbConstantesCC where idseq =" & 15
    End If
    If sqlConstante <> "" Then
        rsConstante.Open sqlConstante, cnBanco, adOpenKeyset, adLockReadOnly
        vPesoConvertido = vPesoConvertido * rsConstante.Fields(0)
        rsConstante.Close
        Set rsConstante = Nothing
        Text11 = Format(vPesoConvertido, "#,##0.00;(#,##0.00)")
    Else
        Text11 = ""
    End If
End Sub
