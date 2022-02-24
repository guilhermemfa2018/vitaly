VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFCE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FCE - Ficha de Controle de Encomenda"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18435
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFCE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   18435
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmFCE.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame9"
      Tab(0).Control(4)=   "Frame20"
      Tab(0).Control(5)=   "Frame21"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Pedidos"
      TabPicture(1)   =   "frmFCE.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ListView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCadastro(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdCadastro(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdCadastro(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame16"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame17"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame18"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame19"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Faturamento"
      TabPicture(2)   =   "frmFCE.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Lista de verificação"
      TabPicture(3)   =   "frmFCE.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame13"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Obs. Administrativas"
      TabPicture(4)   =   "frmFCE.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame11"
      Tab(4).Control(1)=   "Frame12"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame21 
         Caption         =   "Tipo FCE"
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
         Left            =   -69120
         TabIndex        =   147
         Top             =   4680
         Width           =   3375
         Begin ACTIVESKINLibCtl.SkinLabel sknCadastro 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":0D56
            TabIndex        =   148
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Data Book (%)"
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
         Left            =   -69120
         TabIndex        =   143
         Top             =   5640
         Width           =   3375
         Begin VB.TextBox Text3 
            Height          =   330
            Left            =   120
            TabIndex        =   144
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Adiantamento "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   9960
         TabIndex        =   135
         Top             =   1560
         Width           =   2295
         Begin VB.ComboBox Combo4 
            Height          =   345
            ItemData        =   "frmFCE.frx":0DB0
            Left            =   120
            List            =   "frmFCE.frx":0DBD
            TabIndex        =   19
            Text            =   "-"
            Top             =   1320
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":0DD6
            TabIndex        =   137
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txtcadastro 
            Height          =   330
            Index           =   22
            Left            =   120
            TabIndex        =   18
            Text            =   "0"
            Top             =   480
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":0E62
            TabIndex        =   136
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Condições de pagamento"
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
         Left            =   3120
         TabIndex        =   134
         Top             =   2640
         Width           =   3615
         Begin VB.ComboBox Combo3 
            Height          =   345
            ItemData        =   "frmFCE.frx":0EDA
            Left            =   120
            List            =   "frmFCE.frx":0EF3
            TabIndex        =   23
            Text            =   "30 dias após faturamento da NF"
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Somatorios "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   12360
         TabIndex        =   116
         Top             =   1560
         Width           =   5775
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":0FC9
            TabIndex        =   122
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":103F
            TabIndex        =   121
            Top             =   1200
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmFCE.frx":10B5
            TabIndex        =   132
            Top             =   480
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel57 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmFCE.frx":110F
            TabIndex        =   131
            Top             =   240
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":1169
            TabIndex        =   130
            Top             =   1440
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel55 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":11C3
            TabIndex        =   129
            Top             =   1200
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":121D
            TabIndex        =   128
            Top             =   960
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":1277
            TabIndex        =   127
            Top             =   720
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":12D1
            TabIndex        =   126
            Top             =   480
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":132B
            TabIndex        =   125
            Top             =   240
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmFCE.frx":1385
            TabIndex        =   124
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel49 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmFCE.frx":13E9
            TabIndex        =   123
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":1453
            TabIndex        =   120
            Top             =   960
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":14B1
            TabIndex        =   119
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":1513
            TabIndex        =   118
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":1579
            TabIndex        =   117
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "ID"
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
         TabIndex        =   112
         Top             =   2640
         Width           =   855
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":15D9
            TabIndex        =   113
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Valores "
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
         TabIndex        =   69
         Top             =   1560
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   345
            ItemData        =   "frmFCE.frx":1633
            Left            =   5640
            List            =   "frmFCE.frx":1646
            TabIndex        =   12
            Text            =   "KG"
            Top             =   480
            Width           =   855
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   3
            Left            =   4320
            TabIndex        =   11
            Tag             =   "Valor UN c/ impostos"
            ToolTipText     =   "Valo UN c/ impostos"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtcadastro 
            Height          =   330
            Index           =   26
            Left            =   3720
            TabIndex        =   10
            Text            =   "KG"
            Top             =   480
            Width           =   495
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   2
            Left            =   2400
            TabIndex        =   9
            Tag             =   "Peso"
            ToolTipText     =   "Peso"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtcadastro 
            Height          =   330
            Index           =   25
            Left            =   1440
            TabIndex        =   8
            Text            =   "CJ"
            Top             =   480
            Width           =   735
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Tag             =   "Quantidade"
            ToolTipText     =   "Quantidade"
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   5640
            OleObjectBlob   =   "frmFCE.frx":1660
            TabIndex        =   100
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   4320
            OleObjectBlob   =   "frmFCE.frx":16C2
            TabIndex        =   99
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   3720
            OleObjectBlob   =   "frmFCE.frx":1730
            TabIndex        =   98
            Top             =   240
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   2400
            OleObjectBlob   =   "frmFCE.frx":1790
            TabIndex        =   97
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmFCE.frx":17F2
            TabIndex        =   96
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":1854
            TabIndex        =   95
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Impostos "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6840
         TabIndex        =   70
         Top             =   1560
         Width           =   3015
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   10
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":18BA
            TabIndex        =   133
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   7
            Left            =   2280
            TabIndex        =   16
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   6
            Left            =   1560
            TabIndex        =   15
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   5
            Left            =   840
            TabIndex        =   14
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   582
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
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   2280
            OleObjectBlob   =   "frmFCE.frx":191A
            TabIndex        =   104
            Top             =   240
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   1560
            OleObjectBlob   =   "frmFCE.frx":197A
            TabIndex        =   103
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFCE.frx":19DC
            TabIndex        =   102
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":1A42
            TabIndex        =   101
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "FCE referente à FO "
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
         Left            =   15240
         TabIndex        =   67
         Top             =   480
         Width           =   2895
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   330
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Base de Cálculo de ICMS"
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
         Left            =   12360
         TabIndex        =   64
         Top             =   480
         Width           =   2775
         Begin VB.OptionButton Option1 
            Caption         =   "C/ IPI"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "S/ IPI"
            Height          =   195
            Left            =   1560
            TabIndex        =   65
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Itens "
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
         Left            =   -74880
         TabIndex        =   61
         Top             =   480
         Width           =   18015
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5415
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   9551
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
      Begin VB.Frame Frame12 
         Caption         =   "Observações financeiras (Ctrl+Enter - Próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -65520
         TabIndex        =   60
         Top             =   480
         Width           =   8655
         Begin VB.TextBox txtcadastro 
            Height          =   5295
            Index           =   36
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   360
            Width           =   8415
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Observações comerciais (Ctrl+Enter - Próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -74880
         TabIndex        =   59
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtcadastro 
            Height          =   5295
            Index           =   35
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   360
            Width           =   8895
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Relatório de entregas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   18015
         Begin VB.Frame Frame8 
            Caption         =   "Dados da Nota"
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
            TabIndex        =   58
            Top             =   240
            Width           =   17775
            Begin MSMask.MaskEdBox MaskEdBox 
               Height          =   330
               Index           =   9
               Left            =   5280
               TabIndex        =   28
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   582
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.ComboBox Combo2 
               Height          =   345
               ItemData        =   "frmFCE.frx":1AA2
               Left            =   4440
               List            =   "frmFCE.frx":1ABE
               TabIndex        =   27
               Text            =   "KG"
               Top             =   480
               Width           =   735
            End
            Begin MSMask.MaskEdBox MaskEdBox 
               Height          =   330
               Index           =   8
               Left            =   3120
               TabIndex        =   26
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   582
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
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   330
               Left            =   1440
               TabIndex        =   25
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
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
               Format          =   288489473
               CurrentDate     =   40449
            End
            Begin VB.TextBox txtcadastro 
               Height          =   330
               Index           =   31
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
               Height          =   255
               Left            =   5280
               OleObjectBlob   =   "frmFCE.frx":1AE3
               TabIndex        =   109
               Top             =   240
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
               Height          =   255
               Left            =   4440
               OleObjectBlob   =   "frmFCE.frx":1B47
               TabIndex        =   108
               Top             =   240
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
               Height          =   255
               Left            =   3120
               OleObjectBlob   =   "frmFCE.frx":1BAB
               TabIndex        =   107
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
               Height          =   255
               Left            =   1440
               OleObjectBlob   =   "frmFCE.frx":1C0F
               TabIndex        =   106
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFCE.frx":1C71
               TabIndex        =   105
               Top             =   240
               Width           =   1215
            End
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3615
            Left            =   120
            TabIndex        =   54
            Top             =   2040
            Width           =   17775
            _ExtentX        =   31353
            _ExtentY        =   6376
            LabelEdit       =   1
            MultiSelect     =   -1  'True
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
         Begin ZEUS.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   5
            Left            =   1320
            TabIndex        =   55
            Top             =   1320
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
            MICON           =   "frmFCE.frx":1CE7
            PICN            =   "frmFCE.frx":1D03
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   4
            Left            =   720
            TabIndex        =   56
            Top             =   1320
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
            MICON           =   "frmFCE.frx":29DD
            PICN            =   "frmFCE.frx":29F9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   1320
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
            MICON           =   "frmFCE.frx":36D3
            PICN            =   "frmFCE.frx":36EF
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
      Begin VB.Frame Frame9 
         Caption         =   "Observações Técnicas (Ctrl+Enter - Próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6015
         Left            =   -65640
         TabIndex        =   52
         Top             =   480
         Width           =   8775
         Begin VB.TextBox Text18 
            Height          =   5415
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Width           =   8535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados do Cliente "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   10
            Left            =   2880
            TabIndex        =   41
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   9
            Left            =   120
            TabIndex        =   42
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   8
            Left            =   3360
            TabIndex        =   43
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   7
            Left            =   960
            TabIndex        =   44
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   5
            Left            =   2880
            TabIndex        =   46
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   4
            Left            =   120
            TabIndex        =   47
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   3
            Left            =   4200
            TabIndex        =   48
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   1200
            TabIndex        =   50
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Tag             =   "Código do Cliente"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmFCE.frx":43C9
            TabIndex        =   82
            Top             =   2640
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":442B
            TabIndex        =   81
            Top             =   2640
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   3360
            OleObjectBlob   =   "frmFCE.frx":448F
            TabIndex        =   80
            Top             =   2040
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmFCE.frx":44EF
            TabIndex        =   79
            Top             =   2040
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4559
            TabIndex        =   78
            Top             =   2040
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmFCE.frx":45BF
            TabIndex        =   77
            Top             =   1440
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4625
            TabIndex        =   76
            Top             =   1440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   4200
            OleObjectBlob   =   "frmFCE.frx":468B
            TabIndex        =   75
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":46EB
            TabIndex        =   74
            Top             =   840
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFCE.frx":4753
            TabIndex        =   73
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":47B5
            TabIndex        =   72
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
         Height          =   2415
         Left            =   -74880
         TabIndex        =   35
         Top             =   4080
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   14
            Left            =   2640
            TabIndex        =   36
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   13
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   12
            Left            =   1200
            TabIndex        =   38
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   330
            Index           =   11
            Left            =   120
            TabIndex        =   39
            Tag             =   "Código do Contato"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   2640
            OleObjectBlob   =   "frmFCE.frx":481B
            TabIndex        =   86
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":487F
            TabIndex        =   85
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFCE.frx":48E9
            TabIndex        =   84
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":494B
            TabIndex        =   83
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados do Pedido "
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
         TabIndex        =   34
         Top             =   480
         Width           =   12135
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   0
            ItemData        =   "frmFCE.frx":49B1
            Left            =   9720
            List            =   "frmFCE.frx":49B3
            TabIndex        =   146
            Tag             =   "Tipo FCE"
            ToolTipText     =   "Tipo FCE"
            Top             =   480
            Width           =   2295
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
            Height          =   255
            Left            =   9720
            OleObjectBlob   =   "frmFCE.frx":49B5
            TabIndex        =   145
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
            Height          =   330
            Index           =   20
            Left            =   120
            TabIndex        =   5
            Tag             =   "OC nº"
            ToolTipText     =   "OC nº"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtcadastro 
            Height          =   330
            Index           =   21
            Left            =   1440
            TabIndex        =   6
            Top             =   480
            Width           =   8175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmFCE.frx":4A1F
            TabIndex        =   94
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4A8B
            TabIndex        =   93
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
         Height          =   3975
         Left            =   -69120
         TabIndex        =   3
         Top             =   480
         Width           =   3375
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   19
            ItemData        =   "frmFCE.frx":4AEF
            Left            =   120
            List            =   "frmFCE.frx":4AF9
            TabIndex        =   142
            Top             =   3480
            Width           =   3135
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   18
            ItemData        =   "frmFCE.frx":4B0C
            Left            =   120
            List            =   "frmFCE.frx":4B16
            TabIndex        =   141
            Top             =   2880
            Width           =   3135
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   17
            ItemData        =   "frmFCE.frx":4B29
            Left            =   120
            List            =   "frmFCE.frx":4B33
            TabIndex        =   140
            Top             =   2280
            Width           =   3135
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   16
            ItemData        =   "frmFCE.frx":4B46
            Left            =   120
            List            =   "frmFCE.frx":4B50
            TabIndex        =   139
            Top             =   1680
            Width           =   3135
         End
         Begin VB.ComboBox cboCadastro 
            Height          =   345
            Index           =   15
            ItemData        =   "frmFCE.frx":4B63
            Left            =   120
            List            =   "frmFCE.frx":4B6D
            TabIndex        =   138
            Top             =   1080
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
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
            Format          =   162267137
            CurrentDate     =   40449
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4B80
            TabIndex        =   92
            Top             =   3240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4BE8
            TabIndex        =   91
            Top             =   2640
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4C56
            TabIndex        =   90
            Top             =   2040
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4CCA
            TabIndex        =   89
            Top             =   1440
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4D30
            TabIndex        =   88
            Top             =   840
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFCE.frx":4D9E
            TabIndex        =   87
            Top             =   240
            Width           =   1935
         End
      End
      Begin ZEUS.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   1320
         TabIndex        =   22
         Top             =   2760
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
         MICON           =   "frmFCE.frx":4E16
         PICN            =   "frmFCE.frx":4E32
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   720
         TabIndex        =   21
         Top             =   2760
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
         MICON           =   "frmFCE.frx":5B0C
         PICN            =   "frmFCE.frx":5B28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2760
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
         MICON           =   "frmFCE.frx":6802
         PICN            =   "frmFCE.frx":681E
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
         Height          =   3015
         Left            =   120
         TabIndex        =   71
         Top             =   3480
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   5318
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
      Width           =   18255
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1320
         TabIndex        =   62
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
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
         Format          =   162267137
         CurrentDate     =   40449
      End
      Begin VB.TextBox Text1 
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
         Left            =   3000
         TabIndex        =   1
         Top             =   480
         Width           =   15135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "frmFCE.frx":74F8
         TabIndex        =   115
         Top             =   240
         Width           =   3375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmFCE.frx":756E
         TabIndex        =   114
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFCE.frx":75D0
         TabIndex        =   110
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFCE.frx":7634
         TabIndex        =   111
         Top             =   240
         Width           =   975
      End
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   9
      Left            =   720
      TabIndex        =   33
      Top             =   7920
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
      MICON           =   "frmFCE.frx":769A
      PICN            =   "frmFCE.frx":76B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   32
      Top             =   7920
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
      MICON           =   "frmFCE.frx":8390
      PICN            =   "frmFCE.frx":83AC
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
Attribute VB_Name = "frmFCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTreeview As New ADODB.Recordset
Private vContaChecados As Integer

Private Sub cboCadastro_Click(Index As Integer)
    'Msgbox cboCadastro(0).ItemData(cboCadastro(0).ListIndex)
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If ValidaCampo = False Then Exit Sub
        ListView1.Enabled = True
        IncluirItemPed
        LimpaContPed
        SkinLabel40 = Format(GeraCodigoLV(ListView1), "00")
        SomaTotais
        txtcadastro(20).SetFocus
    Case 1
        AlteraItemPed
        SomaTotais
        txtcadastro(20).SetFocus
    Case 2
        ExcluirItemPed
        SomaTotais
    Case 3
        ListView2.Enabled = True
        IncluirItemFat
        LimpaContFat
        txtcadastro(31).SetFocus
    Case 4
        AlteraItemFat
        txtcadastro(31).SetFocus
    Case 5
        ExcluirItemFat
    Case 8
        If vContaChecados > 0 Then
            GravarDados
            Unload Me
        Else
            mobjMsg.Abrir "Favor informar os itens da lista de verificação", Ok, informacao, "ZEUS"
            SSTab1.Tab = 3
        End If
    Case 9
        Unload Me
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    'aceitar somente números e "Back Space", "Enter", "virgula"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCadastro_LostFocus(Index As Integer)
    voltaCorText txtcadastro(Index)
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtcadastro(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtcadastro.Count - 1
        txtcadastro(X).SelStart = 0
        txtcadastro(X).SelLength = Len(txtcadastro(X).Text)
    Next
End Sub


Private Sub MaskEdBox_LostFocus(Index As Integer)
    voltaCorMask MaskEdBox(Index)
End Sub

Private Sub MaskEdBox_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorMask MaskEdBox(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
'    Dim X As Integer
'    For X = 1 To txtcadastro.Count - 1
'        txtcadastro(X).SelStart = 0
'        txtcadastro(X).SelLength = Len(txtcadastro(X).Text)
'    Next
End Sub

Private Sub Form_Load()
    If varGlobal = "-" Then
        mobjMsg.Abrir "Nenhuma FCE selecionada", Ok, critico, "ZEUS"
        'Msgbox "Nenhuma FCE selecionada", vbCritical, "Zeus"
        Unload Me
        Exit Sub
    End If
    SSTab1.Tab = 0
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker4 = Date
'    Label2 = frmRecFO.txtcadastro
    Label2 = varGlobal2
    listview_cabecalho
    CompoeTreeview
    CompoeControles
    
    chamaSQL "SELECT A.ID,A.NUMOC,A.DESCRICAO,A.QUANTIDADE,A.UNQTD,A.PESO,A.UNPESO,A.VALORSIMP,A.PISPERC,A.PISVALOR,A.COFINSPERC,A.COFINSVALOR,A.ICMSPERC,A.ICMSVALOR,A.VALORCIMP,A.UND,A.SUBTOTAL,A.IPIPERC,A.IPIVALOR,A.TOTAL,A.BCALCICMS,A.FOREFERENTE,A.CONDICAOPAG,A.ADIANTAMENTO,A.ADIANTAMENTOCP, CASE WHEN A.TIPOFCEDESC IS NULL THEN '' ELSE A.TIPOFCEDESC END TIPOFCEDESC,CASE WHEN A.TIPOFCEID IS NULL THEN 0 ELSE A.TIPOFCEID END TIPOFCEID FROM TBPEDIDOS AS A WHERE FCE ='" & Val(varGlobal2) & "'"
    
    Compoe_Listview ListView1, Sqlp, "00"
    
    'COMPOE COBO DO TIPO DE FCE (SIMPLES OU COMPLEXA)
    'cboCadastro(0).ItemData(cboCadastro(0).ListIndex) 'PEGA O CODIGO DO TIPO DE FCE SELECIONADA
    CompoeCombo2 cboCadastro(0), "tbTipoFCE", "id", "nome"
    
    SkinLabel40 = Format(GeraCodigoLV(ListView1), "00")
    SomaTotais
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "OC nº", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Quant", ListView1.Width / 20
    ListView1.ColumnHeaders.Add , , "Und.", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "Peso", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Und.", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "Valor UN c/ impostos", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "PIS % ", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "PIS Valor", ListView1.Width / 16
    
    ListView1.ColumnHeaders.Add , , "Cofins %", ListView1.Width / 18
    ListView1.ColumnHeaders.Add , , "Cofins Valor", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "ICMS %", ListView1.Width / 19
    ListView1.ColumnHeaders.Add , , "ICMS Valor", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Valor UN s/ impostos", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Und.", ListView1.Width / 22
    
    ListView1.ColumnHeaders.Add , , "Subtotal", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "IPI %", ListView1.Width / 22
    ListView1.ColumnHeaders.Add , , "IPI Valor", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Total", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Base Cálculo ICMS", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Referente FO", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Condições de pagamento", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Adiantamento%", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "AdiantamentoCP", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "TipoFCE_Desc ", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "TipoFCE_ID", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Add , , "Nota Fiscal", ListView1.Width / 10
    ListView2.ColumnHeaders.Add , , "Data", ListView1.Width / 10
    ListView2.ColumnHeaders.Add , , "Quant.", ListView1.Width / 10
    ListView2.ColumnHeaders.Add , , "Und.", ListView1.Width / 14
    ListView2.ColumnHeaders.Add , , "Valor", ListView1.Width / 4
    
    ListView1.View = lvwReport
    ListView2.View = lvwReport
    
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(10).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(13).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(14).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(15).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(16).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(17).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(18).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(19).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(20).Alignment = lvwColumnRight
    
End Sub

Private Sub CompoeTreeview()
On Error GoTo Err
    Dim rsTree As New ADODB.Recordset
    Dim SqlTree
    Dim no As Node
    Dim X As Integer, Y As Integer
    SqlTree = "Select tbVerifGrupo.codgrupo, tbVerifGrupo.descricao, tbVerifItem.coditem, tbVerifItem.descricao from tbVerifGrupo,tbVerifItem where tbVerifItem.codgrupo=tbVerifGrupo.codgrupo Order by tbVerifItem.codgrupo,tbVerifItem.coditem"
    rsTree.Open SqlTree, cnBanco, adOpenKeyset, adLockOptimistic
    
    TreeView1.Nodes.Clear
    For X = 1 To rsTree.RecordCount
        Set no = TreeView1.Nodes.Add(, , "no" & X, Format(rsTree.Fields(0), "000") & "-" & rsTree.Fields(1))
        Y = rsTree.Fields(0)
        While Y = rsTree.Fields(0)
            TreeView1.Nodes.Add "no" & X, tvwChild, , Format(rsTree.Fields(2), "000") & "-" & rsTree.Fields(3)
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

Private Sub ListView1_DblClick()
    AlteraItemPed
End Sub

Private Sub ListView1_Click()
'    AlteraItemPed
End Sub

Private Sub ListView2_DblClick()
    AlteraItemFat
    SomaTotais
    txtcadastro(20).SetFocus
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    vContaChecados = 0
    With TreeView1
        For i = 1 To .Nodes.Count
            If Not .Nodes(i).Parent Is Nothing Then
                If .Nodes(i).Parent.Key = Node.Key Then
                    .Nodes(i).Checked = Node.Checked
                End If
            End If
            If TreeView1.Nodes(i).Checked = True Then
                vContaChecados = vContaChecados + 1
            End If
        Next i
    End With
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim llng_Contador As Long
    Dim SqlTreeview As String
    Dim Y As Integer, X As Integer, i As Integer
    
    Dim rsFO As New ADODB.Recordset
    Dim rsFCECtrl As New ADODB.Recordset
    Dim rsClientes As New ADODB.Recordset
    Dim rsContatos As New ADODB.Recordset
    Dim sqlFO As String
    Dim sqlFCECtrl As String
    Dim sqlClientes As String
    Dim sqlContatos As String
    vContaChecados = 0
    
    'sqlFCECtrl = "Select * from tbFCE where fce = '" & Val(Label2.Caption) & "' order by fce"
    
    sqlFCECtrl = ""
    sqlFCECtrl = sqlFCECtrl & "SELECT " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.FCE, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.DATAABERTURA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.CARTAPROPOSTA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.OBSERVACAO, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.OBSCOMERCIAL, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.OBSFINANCEIRA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.DATAENTREGA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.FABRICACAO, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.REPARO, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.MATERIAPRIMA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.TRANSPORTE, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.PINTURA, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.DATABOOK, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FCE.STATUS, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FILTRO.FCE AS FCE, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " CASE WHEN LEN([TIPO FCE]) > 0 THEN SUBSTRING([TIPO FCE],1,LEN([TIPO FCE])-1) END AS [TIPO FCE] " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "FROM " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " ( " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " SELECT  FCE, " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "     COALESCE( " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "          FROM TBPEDIDOS AS O " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "          WHERE O.FCE  = C.FCE " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "          GROUP BY TIPOFCEDESC " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO FCE] " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FROM TBPEDIDOS AS C " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " where FCE = '" & Val(Label2.Caption) & "' " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " GROUP BY FCE " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " ) AS FILTRO " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & "LEFT JOIN TBFCE AS FCE ON " & vbCrLf
    sqlFCECtrl = sqlFCECtrl & " FILTRO.FCE = FCE.FCE"
    
    rsFCECtrl.Open sqlFCECtrl, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsFCECtrl.RecordCount > 0 Then
        DTPicker1 = rsFCECtrl.Fields(1)
        Text1.Text = rsFCECtrl.Fields(2)
        txtcadastro(0) = rsFCECtrl.Fields(3)
        txtcadastro(11) = rsFCECtrl.Fields(4)
        Text18 = rsFCECtrl.Fields(5)
        DTPicker2 = rsFCECtrl.Fields(6)
        cboCadastro(15) = rsFCECtrl.Fields(7)
        cboCadastro(16) = rsFCECtrl.Fields(8)
        cboCadastro(17) = rsFCECtrl.Fields(9)
        cboCadastro(18) = rsFCECtrl.Fields(10)
        cboCadastro(19) = rsFCECtrl.Fields(11)
        If Not IsNull(rsFCECtrl.Fields(12)) Then Text3 = rsFCECtrl.Fields(12)
        If Not IsNull(rsFCECtrl.Fields(15)) Then sknCadastro.Caption = rsFCECtrl.Fields(15) ' Tipo FCE
    End If
    rsFCECtrl.Close
    Set rsFCECtrl = Nothing
    
    sqlFO = "select * from tbfo where tbfo.codfo = '" & Val(varGlobal) & "'"
    rsFO.Open sqlFO, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFO.RecordCount > 0 Then txtcadastro(0) = rsFO.Fields(5)
    
    CarregaCli
    CarregaContato
    
    Text2 = varGlobal
    'ContFOSel
    
    SqlTreeview = "Select * from tbListaVerif where tbListaVerif.fce = '" & Val(Me.Label2) & "'"
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockOptimistic
    If rsTreeview.RecordCount > 0 Then
        While Not rsTreeview.EOF
            For llng_Contador = 1 To TreeView1.Nodes.Count
                TreeView1.Nodes(llng_Contador).Expanded = True
                If rsTreeview.Fields(1) = Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 3)) And rsTreeview.Fields(2) = Val(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 3)) Then
                    TreeView1.Nodes(llng_Contador).Checked = True
                    vContaChecados = vContaChecados + 1
                End If
            Next
            rsTreeview.MoveNext
        Wend
    End If
    MaskEdBox(4).PromptInclude = False
    MaskEdBox(4).PromptInclude = False
    MaskEdBox(4).PromptInclude = False
    MaskEdBox(4).PromptInclude = False

    MaskEdBox(4) = "1,65"
    MaskEdBox(5) = "7,60"
    MaskEdBox(6) = "0"
    MaskEdBox(7) = "0"
    MaskEdBox(10) = "0"
    MaskEdBox(4).PromptInclude = True
    MaskEdBox(4).PromptInclude = True
    MaskEdBox(4).PromptInclude = True
    MaskEdBox(4).PromptInclude = True
    Option1.Value = True
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
        Msgbox "Cliente não cadastrado", vbInformation, "Zeus"
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
    SqlContato = "Select * from tbcontatos where tbcontatos.codclifor= '" & Val(txtcadastro(0)) & "'order by nome"
    rsContato.Open SqlContato, cnBanco, adOpenKeyset, adLockOptimistic
    If rsContato.EOF Then
        'MsgBox "Contato não cadastrado", vbInformation, "Zeus"
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

Private Sub IncluirItemPed()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, P As Integer
    Dim SUBTOTAL As Double, PIS As Double, COFINS As Double, ICMS As Double, IPI As Double, VALORUNIT As Double, TOTAL As Double

    'CALCULOS
    If Option1.Value = False Then
    'ICMS s/ IPI
        If Combo1 = "PÇ" Or Combo1 = "CJ" Or Combo1 = "MT²" Then
            SUBTOTAL = Format(MaskEdBox(1) * MaskEdBox(3), "#,##0.000;($#,##0.000)")
        ElseIf Combo1 = "KG" Or Combo1 = "TON" Then
            SUBTOTAL = Format(MaskEdBox(2) * MaskEdBox(3), "#,##0.000;($#,##0.000)")
        Else
            Msgbox "Selecione uma das formas de cálculo disponíveis no combo"
            Exit Sub
        End If
        
        PIS = (Format(SUBTOTAL, "#,##0.000;($#,##0.000)") * Format(MaskEdBox(4), "#,##0.000;($#,##0.000)")) / 100
        COFINS = (SUBTOTAL * MaskEdBox(5)) / 100
        ICMS = (SUBTOTAL * MaskEdBox(6)) / 100
        If MaskEdBox(7) = 0 Then
            IPI = 0
        Else
            IPI = (SUBTOTAL * MaskEdBox(7)) / 100
        End If
        
        If Combo1 = "PÇ" Or Combo1 = "CJ" Or Combo1 = "MT²" Then
            VALORUNIT = (SUBTOTAL - (ICMS + COFINS + PIS)) / Format(MaskEdBox(1), "#,##0.000;($#,##0.000")
        End If
        If Combo1 = "KG" Or Combo1 = "TON" Then
            VALORUNIT = (SUBTOTAL - (ICMS + COFINS + PIS)) / Format(MaskEdBox(2), "#,##0.000;($#,##0.000")
        End If
        
        TOTAL = SUBTOTAL + IPI
    ElseIf Option1.Value = True Then
    'ICMS c/ IPI
        
        If Combo1 = "PÇ" Or Combo1 = "CJ" Or Combo1 = "MT²" Then
            SUBTOTAL = Format(MaskEdBox(1) * MaskEdBox(3), "#,##0.000;($#,##0.000)")
        ElseIf Combo1 = "KG" Or Combo1 = "TON" Then
            SUBTOTAL = Format(MaskEdBox(2) * MaskEdBox(3), "#,##0.000;($#,##0.000)")
        Else
            Msgbox "Selecione uma das formas de cálculo disponíveis no combo"
            Exit Sub
        End If
        PIS = (Format(SUBTOTAL, "#,##0.000;($#,##0.000)") * Format(MaskEdBox(4), "#,##0.000;($#,##0.000)")) / 100
        COFINS = (SUBTOTAL * MaskEdBox(5)) / 100
        If MaskEdBox(7) = 0 Then
            IPI = 0
        Else
            IPI = (SUBTOTAL * MaskEdBox(7)) / 100
        End If
        TOTAL = SUBTOTAL + IPI
        ICMS = (TOTAL * MaskEdBox(6)) / 100
        If Combo1 = "PÇ" Or Combo1 = "CJ" Or Combo1 = "MT²" Then
            VALORUNIT = (SUBTOTAL - (ICMS + COFINS + PIS)) / Format(MaskEdBox(1), "#,##0.000;($#,##0.000")
        End If
        If Combo1 = "KG" Or Combo1 = "TON" Then
            VALORUNIT = (SUBTOTAL - (ICMS + COFINS + PIS)) / Format(MaskEdBox(2), "#,##0.000;($#,##0.000")
        End If
    End If
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView1.ListItems(X).Selected = True
            ListView1.ListItems(X).EnsureVisible
            If ListView1.ListItems.Item(X) = Me.SkinLabel40 Then
                Me.SkinLabel40 = ListView1.ListItems.Item(X) 'Identificador
                ListView1.SelectedItem.ListSubItems.Item(1) = Me.txtcadastro(20) 'Nº da Ordem de compra
                ListView1.SelectedItem.ListSubItems.Item(2) = Me.txtcadastro(21) ''Descrição
                ListView1.SelectedItem.ListSubItems.Item(3) = Format(MaskEdBox(1), "#,##0.00;(#,##0.00)") 'Quantidade
                ListView1.SelectedItem.ListSubItems.Item(4) = Me.txtcadastro(25) 'unidade de medida da quantidade
                ListView1.SelectedItem.ListSubItems.Item(5) = Format(MaskEdBox(2), "#,##0.00;(#,##0.00)") 'Peso
                ListView1.SelectedItem.ListSubItems.Item(6) = Format$(txtcadastro(26), "#,##0.00;(#,##0.00)") 'unidade de medida do peso
'                ListView1.SelectedItem.ListSubItems.Item(7) = Format$(VALORUNIT, "#,##0.00;(#,##0.00)") 'Valor unitário sem impostos
                ListView1.SelectedItem.ListSubItems.Item(7) = Format$(MaskEdBox(3), "#,##0.00;(#,##0.00)") 'Valor unitário sem impostos
                ListView1.SelectedItem.ListSubItems.Item(8) = MaskEdBox(4) '% PIS
                ListView1.SelectedItem.ListSubItems.Item(9) = Format$(PIS, "#,##0.00;(#,##0.00)") 'Valor PIS
                ListView1.SelectedItem.ListSubItems.Item(10) = MaskEdBox(5) '% COFINS
                ListView1.SelectedItem.ListSubItems.Item(11) = Format$(COFINS, "#,##0.00;(#,##0.00)") ' Valor COFINS
                ListView1.SelectedItem.ListSubItems.Item(12) = MaskEdBox(6) '% ICMS
                ListView1.SelectedItem.ListSubItems.Item(13) = Format$(ICMS, "#,##0.00;(#,##0.00)") 'Valor ICMS
                ListView1.SelectedItem.ListSubItems.Item(14) = Format$(VALORUNIT, "#,##0.00;(#,##0.00)") 'Valor unitário com impostos
                ListView1.SelectedItem.ListSubItems.Item(15) = Combo1 'unidade de medida
                ListView1.SelectedItem.ListSubItems.Item(16) = Format$(SUBTOTAL, "#,##0.00;(#,##0.00)") 'subtotal
                ListView1.SelectedItem.ListSubItems.Item(17) = MaskEdBox(7) '% IPI
                ListView1.SelectedItem.ListSubItems.Item(18) = Format$(IPI, "#,##0.00;(#,##0.00)") 'Valor IPI
                ListView1.SelectedItem.ListSubItems.Item(19) = Format$(TOTAL, "#,##0.00;(#,##0.00)") 'Total
                If Option1.Value = True Then ListView1.SelectedItem.ListSubItems.Item(20) = "C/ IPI" Else ListView1.SelectedItem.ListSubItems.Item(20) = "S/ IPI" 'Base de cálculo do ICMS
                ListView1.SelectedItem.ListSubItems.Item(21) = Text2.Text 'Nºs das OC's as quais se referem a FCE
                ListView1.SelectedItem.ListSubItems.Item(22) = Combo3.Text 'Condições de pagamento
                ListView1.SelectedItem.ListSubItems.Item(23) = Me.txtcadastro(22) 'Adiantamento %
                ListView1.SelectedItem.ListSubItems.Item(24) = Combo4.Text 'Adiantamento Condições de pagamento
                If cboCadastro(0).ListIndex <> -1 Then
                    ListView1.SelectedItem.ListSubItems.Item(25) = cboCadastro(0).Text 'Descricao Tipo FCE
                    ListView1.SelectedItem.ListSubItems.Item(26) = cboCadastro(0).ItemData(cboCadastro(0).ListIndex) 'ID Tipo FCE
                End If
                
                cmdCadastro(1).Enabled = True
                cmdCadastro(2).Enabled = True
                ListView1.Enabled = True
                
                Me.ListView1.Sorted = True
                Me.ListView1.SortKey = 0
                Me.ListView1.SortOrder = lvwAscending
                Y = ListView1.ListItems.Count
                Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(10).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(11).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(12).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(13).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(14).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(15).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(16).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(17).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(18).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(19).Alignment = lvwColumnRight
                Me.ListView1.ColumnHeaders(20).Alignment = lvwColumnRight
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , SkinLabel40)
        Me.ListView1.Sorted = True
        Me.ListView1.SortKey = 0
        Me.ListView1.SortOrder = lvwAscending
        Y = ListView1.ListItems.Count
    Else
        Y = ListView1.ListItems.Count
        Set ItemLst = ListView1.ListItems.Add(, , SkinLabel40) 'ListView1.ListItems.Add(Format(Y + 1, "0000"))
    End If
    
    ItemLst.SubItems(1) = Me.txtcadastro(20)
    ItemLst.SubItems(2) = Me.txtcadastro(21)
    ItemLst.SubItems(3) = Format(MaskEdBox(1), "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(4) = Me.txtcadastro(25)
    ItemLst.SubItems(5) = Format(MaskEdBox(2), "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(6) = Format$(txtcadastro(26), "#,##0.00;(#,##0.00)")
'    ItemLst.SubItems(7) = Format$(VALORUNIT, "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(7) = Format$(MaskEdBox(3), "#,##0.00;(#,##0.00)") 'Valor unitário sem impostos
    ItemLst.SubItems(8) = MaskEdBox(4)
    ItemLst.SubItems(9) = Format$(PIS, "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(10) = MaskEdBox(5)
    ItemLst.SubItems(11) = Format$(COFINS, "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(12) = MaskEdBox(6)
    ItemLst.SubItems(13) = Format$(ICMS, "#,##0.00;(#,##0.00)")
'    ItemLst.SubItems(14) = Format$(MaskEdBox(3), "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(14) = Format$(VALORUNIT, "#,##0.00;(#,##0.00)") 'Valor unitário com impostos
    ItemLst.SubItems(15) = Combo1
    ItemLst.SubItems(16) = Format$(SUBTOTAL, "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(17) = MaskEdBox(7)
    ItemLst.SubItems(18) = Format$(IPI, "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(19) = Format$(TOTAL, "#,##0.00;(#,##0.00)")
    If Option1.Value = True Then ItemLst.SubItems(20) = "C/ IPI" Else ItemLst.SubItems(20) = "S/ IPI"
    ItemLst.SubItems(21) = Text2.Text 'Nºs das OC's as quais se referem a FCE
    ItemLst.SubItems(22) = Combo3.Text 'Condições de pagamento
    ItemLst.SubItems(23) = Me.txtcadastro(22) 'Adiantamento %
    ItemLst.SubItems(24) = Combo4.Text 'Adiantamento Condições de pagamento
    ItemLst.SubItems(25) = cboCadastro(0).Text 'Descricao Tipo FCE
    ItemLst.SubItems(26) = cboCadastro(0).ItemData(cboCadastro(0).ListIndex) 'ID Tipo FCE
    
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(10).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(13).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(14).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(15).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(16).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(17).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(18).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(19).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(20).Alignment = lvwColumnRight
    cmdCadastro(1).Enabled = True
    cmdCadastro(2).Enabled = True
    ListView1.Enabled = True
End Sub

Private Sub AlteraItemPed()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.SkinLabel40.Caption = ListView1.ListItems.Item(X)
    Me.txtcadastro(20).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.txtcadastro(21).Text = ListView1.SelectedItem.ListSubItems.Item(2)
    Me.MaskEdBox(1) = ListView1.SelectedItem.ListSubItems.Item(3)
    Me.txtcadastro(25) = ListView1.SelectedItem.ListSubItems.Item(4)
    Me.MaskEdBox(2) = ListView1.SelectedItem.ListSubItems.Item(5)
    Me.txtcadastro(26).Text = ListView1.SelectedItem.ListSubItems.Item(6)
    Me.MaskEdBox(4) = ListView1.SelectedItem.ListSubItems.Item(8)
    Me.MaskEdBox(5) = ListView1.SelectedItem.ListSubItems.Item(10)
    Me.MaskEdBox(6) = ListView1.SelectedItem.ListSubItems.Item(12)
    
    Me.MaskEdBox(3) = ListView1.SelectedItem.ListSubItems.Item(7)
    Me.Combo1 = ListView1.SelectedItem.ListSubItems.Item(15)
'    Me.MaskEdBox(3) = ListView1.SelectedItem.ListSubItems.Item(14)
'    Me.Combo1 = ListView1.SelectedItem.ListSubItems.Item(15)
    Me.MaskEdBox(7) = ListView1.SelectedItem.ListSubItems.Item(17)
    If ListView1.SelectedItem.ListSubItems.Item(20) = "C/ IPI" Then Option1.Value = True Else Option2.Value = True
    Me.Text2.Text = ListView1.SelectedItem.ListSubItems.Item(21)
    Me.Combo3.Text = ListView1.SelectedItem.ListSubItems.Item(22)
    Me.txtcadastro(22).Text = ListView1.SelectedItem.ListSubItems.Item(23)
    Me.Combo4.Text = ListView1.SelectedItem.ListSubItems.Item(24)
    Me.cboCadastro(0).Text = ListView1.SelectedItem.ListSubItems.Item(25)

End Sub

Private Sub ExcluirItemPed()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    ListView1.ListItems.Remove (X)
End Sub

Private Sub IncluirItemFat()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, P As Integer
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView2.ListItems(X).Selected = True
            ListView2.ListItems(X).EnsureVisible
            If ListView2.ListItems.Item(X) = Me.txtcadastro(31) Then
                Me.txtcadastro(31) = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = Me.DTPicker4
                ListView2.SelectedItem.ListSubItems.Item(2) = Format(MaskEdBox(8), "#,##0.00;(#,##0.00)")
                ListView2.SelectedItem.ListSubItems.Item(3) = Me.Combo2
                ListView2.SelectedItem.ListSubItems.Item(4) = Format(MaskEdBox(9), "#,##0.00;(#,##0.00)")
                
                cmdCadastro(4).Enabled = True
                cmdCadastro(5).Enabled = True
                ListView2.Enabled = True
                
                Me.ListView2.Sorted = True
                Me.ListView2.SortKey = 0
                Me.ListView2.SortOrder = lvwAscending
                Y = ListView2.ListItems.Count
                Me.ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
                Me.ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , txtcadastro(31))
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 0
        Me.ListView2.SortOrder = lvwAscending
        Y = ListView2.ListItems.Count
    Else
        Y = ListView2.ListItems.Count
        Set ItemLst = ListView2.ListItems.Add(, , txtcadastro(31)) 'ListView2.ListItems.Add(Format(Y + 1, "0000"))
    End If
    ItemLst.SubItems(1) = Me.DTPicker4
    ItemLst.SubItems(2) = Format(MaskEdBox(8), "#,##0.00;(#,##0.00)")
    ItemLst.SubItems(3) = Me.Combo2
    ItemLst.SubItems(4) = Format(MaskEdBox(9), "#,##0.00;(#,##0.00)")
    
    Me.ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    cmdCadastro(4).Enabled = True
    cmdCadastro(5).Enabled = True
    ListView2.Enabled = True
End Sub

Private Sub AlteraItemFat()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtcadastro(31).Text = ListView2.ListItems.Item(X)
    Me.DTPicker4 = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.MaskEdBox(8) = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.Combo2 = ListView2.SelectedItem.ListSubItems.Item(3)
    Me.MaskEdBox(9) = ListView2.SelectedItem.ListSubItems.Item(4)
End Sub

Private Sub ExcluirItemFat()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    ListView2.ListItems.Remove (X)
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    For X = 1 To 3
        If MaskEdBox(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o valor do campo " & MaskEdBox(X).Tag, Ok, informacao, "Atenção"
            MaskEdBox(X).SetFocus
            Exit Function
        End If
    Next
    
    If cboCadastro(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o valor do campo " & cboCadastro(0).Tag, Ok, informacao, "Atenção"
        cboCadastro(0).SetFocus
        Exit Function
    End If
    
    If txtcadastro(20).Text = "" Then
        mobjMsg.Abrir "Favor informar o valor do campo " & txtcadastro(20).Tag, Ok, informacao, "Atenção"
        txtcadastro(20).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub GravarDados()
On Error GoTo Err
    'If ValidaCampo = False Then Exit Sub
    Dim rsDeleta As New ADODB.Recordset
    Dim rsGravaFCE As New ADODB.Recordset
    Dim rsGravaPedidos As New ADODB.Recordset
    Dim rsGravaFaturamento As New ADODB.Recordset
    Dim rsGravaListaVer As New ADODB.Recordset
    Dim rsGravaFO As New ADODB.Recordset
    vTransacaoAtiva = 0
    Dim sqlExc As String
    Dim sql As String
    Dim Y As Integer, X As Integer
10  cnBanco.BeginTrans
    sql = "Select * from tbfo"
    rsGravaFO.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        MeuLV.ListView1.ListItems(X).Selected = True
        MeuLV.ListView1.ListItems(X).EnsureVisible
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            While Not rsGravaFO.EOF
                If Val(MeuLV.ListView1.ListItems.Item(X)) = rsGravaFO.Fields(0) Then
                    rsGravaFO.Fields(2) = 2
                End If
                rsGravaFO.MoveNext
            Wend
        End If
    Next
'    rsGravaFO.Fields(2) = 2
    If Not rsGravaFO.EOF Then rsGravaFO.Update
    rsGravaFO.Close
    
    sql = "Select * from tbFCE order by fce"
    rsGravaFCE.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
    sqlExc = "Delete from tbfce where tbfce.fce = '" & Val(Label2.Caption) & "'"
    rsDeleta.Open sqlExc, cnBanco
    
    rsGravaFCE.AddNew
    rsGravaFCE.Fields(0) = Label2.Caption
    rsGravaFCE.Fields(1) = DTPicker1
    rsGravaFCE.Fields(2) = Text1.Text
    rsGravaFCE.Fields(3) = txtcadastro(0)
    rsGravaFCE.Fields(4) = txtcadastro(11)
    rsGravaFCE.Fields(5) = Text18
    rsGravaFCE.Fields(6) = DTPicker2
    rsGravaFCE.Fields(7) = cboCadastro(15)
    rsGravaFCE.Fields(8) = cboCadastro(16)
    rsGravaFCE.Fields(9) = cboCadastro(17)
    rsGravaFCE.Fields(10) = cboCadastro(18)
    rsGravaFCE.Fields(11) = cboCadastro(19)
    rsGravaFCE.Fields(12) = Text3.Text 'Databook
    rsGravaFCE.Fields(13) = 0 'Status
    
    If Not rsGravaFCE.EOF Then rsGravaFCE.Update
    rsGravaFCE.Close
    
    sql = "Select * from tbpedidos where tbpedidos.fce = '" & Val(Label2) & "'"
    rsGravaPedidos.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
    sqlExc = "Delete from tbpedidos where tbpedidos.fce = '" & Val(Label2) & "'"
    rsDeleta.Open sqlExc, cnBanco
    
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaPedidos.AddNew
        rsGravaPedidos(0) = Val(ListView1.ListItems.Item(X))
        rsGravaPedidos(1) = Label2
        rsGravaPedidos(2) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsGravaPedidos(3) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsGravaPedidos(4) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsGravaPedidos(5) = ListView1.SelectedItem.ListSubItems.Item(4)
        rsGravaPedidos(6) = ListView1.SelectedItem.ListSubItems.Item(5)
        rsGravaPedidos(7) = ListView1.SelectedItem.ListSubItems.Item(6)
        rsGravaPedidos(8) = ListView1.SelectedItem.ListSubItems.Item(7)
        rsGravaPedidos(9) = ListView1.SelectedItem.ListSubItems.Item(8)
        rsGravaPedidos(10) = ListView1.SelectedItem.ListSubItems.Item(9)
        rsGravaPedidos(11) = ListView1.SelectedItem.ListSubItems.Item(10)
        rsGravaPedidos(12) = ListView1.SelectedItem.ListSubItems.Item(11)
        rsGravaPedidos(13) = ListView1.SelectedItem.ListSubItems.Item(12)
        rsGravaPedidos(14) = ListView1.SelectedItem.ListSubItems.Item(13)
        rsGravaPedidos(15) = ListView1.SelectedItem.ListSubItems.Item(14)
        rsGravaPedidos(16) = ListView1.SelectedItem.ListSubItems.Item(15)
        rsGravaPedidos(17) = ListView1.SelectedItem.ListSubItems.Item(16)
        rsGravaPedidos(18) = ListView1.SelectedItem.ListSubItems.Item(17)
        rsGravaPedidos(19) = ListView1.SelectedItem.ListSubItems.Item(18)
        rsGravaPedidos(20) = ListView1.SelectedItem.ListSubItems.Item(19)
        rsGravaPedidos(21) = ListView1.SelectedItem.ListSubItems.Item(20)
        rsGravaPedidos(22) = Text2.Text
        rsGravaPedidos(23) = ListView1.SelectedItem.ListSubItems.Item(22) 'Combo3.Text
        If ListView1.SelectedItem.ListSubItems.Item(23) <> "" Then rsGravaPedidos(24) = ListView1.SelectedItem.ListSubItems.Item(23)
        If ListView1.SelectedItem.ListSubItems.Item(24) <> "" Then rsGravaPedidos(25) = ListView1.SelectedItem.ListSubItems.Item(24)
        If ListView1.SelectedItem.ListSubItems.Item(25) <> "" Then rsGravaPedidos(26) = ListView1.SelectedItem.ListSubItems.Item(25)
        If ListView1.SelectedItem.ListSubItems.Item(26) <> "" Then rsGravaPedidos(27) = ListView1.SelectedItem.ListSubItems.Item(26)
    Next
    If Not rsGravaPedidos.EOF Then rsGravaPedidos.Update
    rsGravaPedidos.Close
    
    sql = "Select * from tbfaturamento where tbfaturamento.fce = '" & Val(Label2) & "'"
    rsGravaFaturamento.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
    sqlExc = "Delete from tbfaturamento where tbfaturamento.fce = '" & Val(Label2) & "'"
    rsDeleta.Open sqlExc, cnBanco
    
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaFaturamento.AddNew
        rsGravaFaturamento(0) = Label2
        rsGravaFaturamento(1) = ListView2.ListItems.Item(X)
        rsGravaFaturamento(2) = ListView2.SelectedItem.ListSubItems.Item(1)
        rsGravaFaturamento(3) = Val(ListView2.SelectedItem.ListSubItems.Item(2))
        rsGravaFaturamento(4) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsGravaFaturamento(5) = ListView2.SelectedItem.ListSubItems.Item(4)
    Next
    If Not rsGravaFaturamento.EOF Then rsGravaFaturamento.Update
    rsGravaFaturamento.Close
    
    sql = "Select * from tblistaverif where tblistaverif.fce = '" & Val(Label2) & "'"
    rsGravaListaVer.Open sql, cnBanco, adOpenKeyset, adLockOptimistic

    sqlExc = "Delete from tblistaverif where tblistaverif.fce = '" & Val(Label2) & "'"
    rsDeleta.Open sqlExc, cnBanco, adOpenKeyset, adLockOptimistic

    With TreeView1
        For i = 1 To .Nodes.Count
            If InStr(TreeView1.Nodes(i).FullPath, "\") <> 0 Then
                If TreeView1.Nodes(i).Checked = True Then
                    rsGravaListaVer.AddNew
                    rsGravaListaVer.Fields(0) = Label2.Caption
                    rsGravaListaVer.Fields(1) = Val(Mid$(TreeView1.Nodes(i).FullPath, 1, 3))
                    rsGravaListaVer.Fields(2) = Val(Mid$(TreeView1.Nodes(i).FullPath, InStr(TreeView1.Nodes(i).FullPath, "\") + 1, 3))
                End If
            End If
        Next
    End With
    If Not rsGravaListaVer.EOF Then rsGravaListaVer.Update
    rsGravaListaVer.Close

'----Inicio da Rotina p gravar numero de FCE na FO---------
    Y = MeuLV.ListView1.ListItems.Count
    sql = "Select * from tbFO"
    rsGravaFO.Open sql, cnBanco, adOpenKeyset, adLockOptimistic

    For X = 1 To Y
        MeuLV.ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(13) = Label2.Caption
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(14) = "Serviço"
            rsGravaFO.MoveFirst
            rsGravaFO.Find "codfo=" & "'" & Val(MeuLV.ListView1.ListItems.Item(X)) & "'"
            If Not rsGravaFO.EOF Then
                rsGravaFO.Fields(3) = Label2.Caption
                rsGravaFO.Fields(2) = 2
            End If
        End If
    Next
'----Fim da Rotina p gravar numero de FCE na FO---------
    rsGravaFO.Update
    rsGravaFO.Close

    cnBanco.CommitTrans
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

Private Sub LimpaContPed()
    Dim X As Integer
    MaskEdBox(1).PromptInclude = False
    MaskEdBox(2).PromptInclude = False
    MaskEdBox(3).PromptInclude = False
    MaskEdBox(1) = ""
    MaskEdBox(2) = ""
    MaskEdBox(3) = ""
    MaskEdBox(1).PromptInclude = True
    MaskEdBox(2).PromptInclude = True
    MaskEdBox(3).PromptInclude = True
    For X = 20 To 21
        txtcadastro(X) = ""
    Next
End Sub

Private Sub LimpaContFat()
    Dim X As Integer
    MaskEdBox(8).PromptInclude = False
    MaskEdBox(9).PromptInclude = False
    MaskEdBox(8) = ""
    MaskEdBox(9) = ""
    MaskEdBox(8).PromptInclude = True
    MaskEdBox(9).PromptInclude = True
    txtcadastro(31) = ""
    Combo2 = "KG"
End Sub

Private Sub ContFOSel()
    Dim Y As Integer, codfornec As Integer
    Dim numFO As String
    Contador = 0
    codfornec = 0
    Mensagem = ""
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        MeuLV.ListView1.ListItems(X).Selected = True
        MeuLV.ListView1.ListItems(X).EnsureVisible
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            If numFO = "" Then
                numFO = MeuLV.ListView1.ListItems.Item(X)
            Else
                numFO = numFO & "," & MeuLV.ListView1.ListItems.Item(X)
            End If
            Contador = Contador + 1
        End If
    Next
    Text2 = numFO
End Sub

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
End Sub

Private Function SomaTotais()
On Error GoTo TrataErro
    SkinLabel12.Caption = ""
    SomaTotais = True
    Dim Y As Integer, vPIS As Double, vCofins As Double, vICMS As Double, vIPI As Double, vVSImp As Double, vCCImp As Double, vSubTotal As Double, vTotal As Double
    Y = ListView1.ListItems.Count
    vPIS = 0
    vCofins = 0
    vICMS = 0
    vIPI = 0
    vVSImp = 0
    vCCImp = 0
    vSubTotal = 0
    vTotal = 0
    For W = 1 To Y
        ListView1.ListItems(W).Selected = True
        ListView1.SelectedItem.ListSubItems.Item(9) = Format(ListView1.SelectedItem.ListSubItems.Item(9), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(11) = Format(ListView1.SelectedItem.ListSubItems.Item(11), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(13) = Format(ListView1.SelectedItem.ListSubItems.Item(13), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(18) = Format(ListView1.SelectedItem.ListSubItems.Item(18), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(14) = Format(ListView1.SelectedItem.ListSubItems.Item(14), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(7) = Format(ListView1.SelectedItem.ListSubItems.Item(7), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(16) = Format(ListView1.SelectedItem.ListSubItems.Item(16), "#,##0.00;(#,##0.00)")
        ListView1.SelectedItem.ListSubItems.Item(19) = Format(ListView1.SelectedItem.ListSubItems.Item(19), "#,##0.00;(#,##0.00)")
        
        vPIS = vPIS + ListView1.SelectedItem.ListSubItems.Item(9)
        vCofins = vCofins + ListView1.SelectedItem.ListSubItems.Item(11)
        vICMS = vICMS + ListView1.SelectedItem.ListSubItems.Item(13)
        vIPI = vIPI + ListView1.SelectedItem.ListSubItems.Item(18)
        vVSImp = vVSImp + ListView1.SelectedItem.ListSubItems.Item(14)
        vCCImp = vCCImp + ListView1.SelectedItem.ListSubItems.Item(7)
        vSubTotal = vSubTotal + ListView1.SelectedItem.ListSubItems.Item(16)
        vTotal = vTotal + ListView1.SelectedItem.ListSubItems.Item(19)
    Next
    SkinLabel51 = Format(vPIS, "#,##0.00;(#,##0.00)")
    SkinLabel52 = Format(vCofins, "#,##0.00;(#,##0.00)")
    SkinLabel53 = Format(vICMS, "#,##0.00;(#,##0.00)")
    SkinLabel54 = Format(vIPI, "#,##0.00;(#,##0.00)")
    SkinLabel55 = Format(vVSImp, "#,##0.00;(#,##0.00)")
    SkinLabel56 = Format(vCCImp, "#,##0.00;(#,##0.00)")
    SkinLabel57 = Format(vSubTotal, "#,##0.00;(#,##0.00)")
    SkinLabel58 = Format(vTotal, "#,##0.00;(#,##0.00)")
    Exit Function
TrataErro:
    SomaTotais = False
    'mobjMsg.Abrir "Existem itens marcados no Relatorio que não possuem Qtd. liberada", Ok, informacao, "Atenção"
    'SkinLabel12.Caption = "Existem itens marcados no Relatorio que não possuem Qtd. liberada"
End Function
