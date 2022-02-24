VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRNCF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RNC - Registro de Não Conformidade"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16590
   Icon            =   "frmRNCF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   16590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame14 
      Caption         =   "Legenda do Status"
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
      Left            =   1440
      TabIndex        =   73
      Top             =   10200
      Width           =   6615
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   255
         Left            =   6120
         OleObjectBlob   =   "frmRNCF.frx":0CCA
         TabIndex        =   86
         Top             =   240
         Width           =   375
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   85
         Top             =   240
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmRNCF.frx":0D28
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4800
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   83
         Top             =   240
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "frmRNCF.frx":0D86
         TabIndex        =   82
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3600
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   80
         Top             =   240
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmRNCF.frx":0DE8
         TabIndex        =   79
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   78
         Top             =   240
         Width           =   255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "frmRNCF.frx":0E48
         TabIndex        =   77
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "frmRNCF.frx":0EAA
         TabIndex        =   76
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   75
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   74
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   120
      Picture         =   "frmRNCF.frx":0F0A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10200
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   13
      Left            =   720
      Picture         =   "frmRNCF.frx":1BD4
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10200
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   40
      Top             =   3720
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Identificação"
      TabPicture(0)   =   "frmRNCF.frx":289E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Identificação Itens"
      TabPicture(1)   =   "frmRNCF.frx":28BA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame31"
      Tab(1).Control(1)=   "cmdCadastro(7)"
      Tab(1).Control(2)=   "cmdCadastro(8)"
      Tab(1).Control(3)=   "Frame11"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Plano de ação"
      TabPicture(2)   =   "frmRNCF.frx":28D6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame8"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame12"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Combo2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRNCF.frx":28F2
         Left            =   8400
         List            =   "frmRNCF.frx":28FF
         TabIndex        =   62
         Tag             =   "Tipos de Ações"
         ToolTipText     =   "Tipos de Ações"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Frame Frame12 
         Caption         =   " Ações Corretivas e/ou Preventivas (Crtl+ Enter - próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   8280
         TabIndex        =   61
         Top             =   3240
         Width           =   7935
         Begin VB.TextBox txtRNC 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Index           =   16
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   64
            Tag             =   "Ações Corretivas e/ou Preventivas"
            ToolTipText     =   "Ações Corretivas e/ou Preventivas"
            Top             =   600
            Width           =   7695
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Causa Raiz Determinada (Crtl+ Enter - próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   8280
         TabIndex        =   60
         Top             =   360
         Width           =   7935
         Begin VB.TextBox txtRNC 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Index           =   15
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   63
            Tag             =   "Causa Raiz Determinada"
            ToolTipText     =   "Causa Raiz Determinada"
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Identifique o CC - Centro de Custo responsável pela não conformidade "
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
         Left            =   -74880
         TabIndex        =   58
         Top             =   5400
         Width           =   8175
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   59
            Tag             =   "Centro de Custo responsável pela não conformidade"
            ToolTipText     =   "Centro de Custo responsável pela não conformidade"
            Top             =   360
            Width           =   7935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Correção Adotada (Crtl+ Enter - próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   51
         Top             =   3240
         Width           =   7935
         Begin VB.TextBox txtRNC 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Index           =   12
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   24
            Tag             =   "Correção"
            ToolTipText     =   "Correção"
            Top             =   240
            Width           =   7695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Não Conformidade Ocorrida (Crtl+ Enter - próxima linha)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   7935
         Begin VB.TextBox txtRNC 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Index           =   11
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   23
            Tag             =   "Incidente ocorrido"
            ToolTipText     =   "Incidente ocorrido"
            Top             =   240
            Width           =   7695
         End
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
         Height          =   5895
         Left            =   -66240
         TabIndex        =   44
         Top             =   360
         Width           =   7455
         Begin VB.Frame Frame6 
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
            TabIndex        =   46
            Top             =   4800
            Width           =   7215
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   285
               Left            =   5160
               TabIndex        =   22
               Top             =   480
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CheckBox        =   -1  'True
               Format          =   244187137
               CurrentDate     =   41780
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   5160
               OleObjectBlob   =   "frmRNCF.frx":2930
               TabIndex        =   49
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox txtRNC 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   10
               Left            =   2640
               TabIndex        =   21
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox txtRNC 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   120
               TabIndex        =   20
               Top             =   480
               Width           =   2415
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   2640
               OleObjectBlob   =   "frmRNCF.frx":29AA
               TabIndex        =   48
               Top             =   240
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmRNCF.frx":2A34
               TabIndex        =   47
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   4440
            Visible         =   0   'False
            Width           =   7215
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4575
            Left            =   120
            TabIndex        =   19
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   8070
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
      Begin VB.CommandButton cmdCadastro 
         Caption         =   ">"
         Height          =   615
         Index           =   7
         Left            =   -67200
         TabIndex        =   17
         Top             =   1800
         Width           =   735
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "<"
         Height          =   615
         Index           =   8
         Left            =   -67200
         TabIndex        =   18
         Top             =   2520
         Width           =   735
      End
      Begin VB.Frame Frame11 
         Caption         =   "Itens da Operação"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74880
         TabIndex        =   43
         Top             =   360
         Width           =   7455
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5535
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   9763
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
      Begin VB.Frame Frame5 
         Caption         =   "Causais"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -66480
         TabIndex        =   42
         Top             =   360
         Width           =   7575
         Begin MSComctlLib.ListView ListView2 
            Height          =   5535
            Left            =   120
            TabIndex        =   15
            Tag             =   "Causais"
            ToolTipText     =   "Causais"
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   9763
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483635
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
      Begin VB.Frame Frame4 
         Caption         =   "Centros de Custos Orçados (Identifique onde ocorreu o problema)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   8175
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   54
            Top             =   4440
            Visible         =   0   'False
            Width           =   2415
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4575
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   8070
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483635
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
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   14160
      TabIndex        =   38
      Top             =   10080
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "Gerou Retrabalho"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados da CD - Comunicação de Desvio "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   8175
      Begin VB.TextBox txtRNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   4320
         TabIndex        =   69
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "frmRNCF.frx":2AA0
         TabIndex        =   68
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   5160
         TabIndex        =   6
         Top             =   1080
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmRNCF.frx":2B02
         TabIndex        =   53
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":2B6A
         TabIndex        =   52
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3480
         OleObjectBlob   =   "frmRNCF.frx":2BD2
         TabIndex        =   39
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtRNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   1575
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   7935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":2C32
         TabIndex        =   37
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5160
         TabIndex        =   3
         Top             =   480
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "frmRNCF.frx":2CA0
         TabIndex        =   36
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   162660353
         CurrentDate     =   41773
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmRNCF.frx":2D2C
         TabIndex        =   35
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtRNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRNCF.frx":2DA0
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":2E04
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do RNC - Registro de Não Conformidade "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   8400
      TabIndex        =   28
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame13 
         Caption         =   "Status"
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
         Left            =   6960
         TabIndex        =   70
         Top             =   240
         Width           =   975
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   71
            Top             =   360
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRNCF.frx":2E68
            TabIndex        =   72
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   285
         Left            =   5040
         TabIndex        =   66
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
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
         CheckBox        =   -1  'True
         Format          =   162660353
         CurrentDate     =   41809
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   5040
         OleObjectBlob   =   "frmRNCF.frx":2EC2
         TabIndex        =   65
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmRNCF.frx":2F3A
         TabIndex        =   56
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtRNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   10
         Tag             =   "Registro do colaborador responsável pelo registro da RNC"
         ToolTipText     =   "Registro do colaborador responsável pelo registro da RNC"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtRNC 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   8
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   13
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   1680
         Width           =   7815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":2FB6
         TabIndex        =   31
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   7440
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   1800
         TabIndex        =   11
         Tag             =   "Nome do colaborador responsável pelo registro da RNC"
         ToolTipText     =   "Nome do colaborador responsável pelo registro da RNC"
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtRNC 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":3062
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Tag             =   "Data início"
         ToolTipText     =   "Data início"
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Format          =   162660353
         CurrentDate     =   41366
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   1440
         OleObjectBlob   =   "frmRNCF.frx":30C8
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRNCF.frx":3142
         TabIndex        =   55
         Top             =   840
         Width           =   4335
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   3120
         TabIndex        =   57
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
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
         CheckBox        =   -1  'True
         Format          =   162660353
         CurrentDate     =   41787
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   375
      Left            =   8160
      OleObjectBlob   =   "frmRNCF.frx":31FC
      TabIndex        =   67
      Top             =   10320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   7680
      TabIndex        =   81
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "frmRNCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vNomeA As String
Private vNomeB As String
Private vNomeC As String
Private vJuntaNome As String
Private vNmNo As String
Private rsRNC As New ADODB.Recordset
Private SqlRNC As String
Private rsDeletar As New ADODB.Recordset
Private sqlDeletar As String
Private rsLocal As New ADODB.Recordset
Private vStatus As Integer
Private vCausais As String, vCCNConforme As String, vItens As String

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 7
        Text1.Text = ""
        sqlDeletar = "Delete from tbMPDesSel"
        rsDeletar.Open sqlDeletar, cnBanco
        buscaChecado2 TreeView1
        mostraDesenhos "tbMPDesSel", TreeView2
    Case 8
        Text1.Text = ""
        buscaChecado2 TreeView2
        mostraDesenhos "tbMPDesSel", TreeView2
    Case 9
        ChamaGridColab
        chamaChapa
    Case 12
        
        If DTPicker4.Value <> "" And IsNull(DTPicker5.Value) Then
            If verificaDados(8) = False Then Exit Sub
            If vStatus <> 10 And vStatus <> 20 Then
                vStatus = 20
            End If
        ElseIf DTPicker5.Value <> "" And DTPicker4.Value <> "" Then
            If verificaDados(20) = False Then Exit Sub
            If vStatus <> 10 And vStatus <> 20 Then
                vStatus = 8
            End If
        Else
            If vStatus <> 10 And vStatus <> 20 Then
                vStatus = 7
            End If
        End If
        
        If salvar_Dados = True Then
            AtualizaListview
            If vStatus = 7 Then
                mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
            ElseIf vStatus = 20 Then
                mobjMsg.Abrir "Conclusão realizada com sucesso!", Ok, informacao, "ZEUS"
                If Check1.Value = 1 Then
                    If checaEnvioEmail = False Then
                        enviaEmail
                    Else
                        mobjMsg.Abrir "Deseja enviar novamente o email", YesNo, pergunta, "Zeus"
                        If Tp = 1 Then
                            enviaEmail
                        End If
                    End If
                End If
                Unload Me
            ElseIf vStatus = 8 Or vStatus = 10 Then
                mobjMsg.Abrir "Fechamento realizado com sucesso!", Ok, informacao, "ZEUS"
                Unload Me
            End If
        Else
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ZEUS"
        End If
    Case 13
        Unload Me
    End Select
End Sub

Private Function checaEnvioEmail()
On Error GoTo Err
    checaEnvioEmail = False
    Dim rschecaEnvioEmail As New ADODB.Recordset
    Dim SqlchecaEnvioEmail As String
    Dim rsAtualizaEnvioEmail As New ADODB.Recordset
    Dim SqlAtualizaEnvioEmail As String
    
    SqlchecaEnvioEmail = "Select a.email from tbrnc as a where a.idrnc = '" & Val(txtRNC(7).Text) & "'"
    rschecaEnvioEmail.Open SqlchecaEnvioEmail, cnBanco, adOpenKeyset, adLockReadOnly
    If rschecaEnvioEmail.Fields(0) = 1 Then
        checaEnvioEmail = True
    Else
        SqlAtualizaEnvioEmail = "Update tbrnc set email = 1 where idrnc = '" & Val(txtRNC(7).Text) & "'"
        rsAtualizaEnvioEmail.Open SqlAtualizaEnvioEmail, cnBanco
    End If
    rschecaEnvioEmail.Close
    Set rschecaEnvioEmail = Nothing
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

Private Sub Form_Load()
    DTPicker2 = Date
    DTPicker3 = Date
    DTPicker4 = Date
    DTPicker5 = Date
    DTPicker5.Value = ""
    DTPicker4.Value = ""
    DTPicker3.Value = ""
    SSTab1.Tab = 0
    listview_cabecalho
    
    Status = Pesquisa
    ResultPesq
    'SE STATUS FOR = A CONCLUIDO, BLOQUEIA EDIÇÃO
    If MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "CONCLUIDO1" Or MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "PRETO" Or MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "FABRICANDO" Then bloqueiaEdicao
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
End Sub

Private Sub txtRNC_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtRNC(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim x As Integer
    For x = 1 To txtRNC.Count - 1
        txtRNC(x).SelStart = 0
        txtRNC(x).SelLength = Len(txtRNC(x).Text)
    Next
End Sub

Private Sub txtRNC_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1 Or 17
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtRNC(17) = "" Then
                mobjMsg.Abrir "Digite o nº da revisao da OS", Ok, critico, "Atenção"
                Exit Sub
            End If
            valida_OS
            'Abaixo: Limpa listview e em seguida preenche novamente
            ListView1.ListItems.Clear
            chamaSQL "select a.idcc,a.nomecc,a.desenhos from tbMPItens as a where idos = '" & Val(txtRNC(1)) & "' order by a.idoperacao"
            Compoe_Listview ListView1, Sqlp, ""
        End If
    Case 13
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If chamaChapa = False Then Exit Sub
        End If
    End Select
End Sub

Private Function chamaChapa()
On Error GoTo Err
    chamaChapa = False
    Dim rschamaChapa As New ADODB.Recordset
    Dim SqlchamaChapa As String
    
    SqlchamaChapa = "select a.chapa,a.nome from " & vBancoTotvs & ".dbo.PFUNC as a where a.CODCOLIGADA = 1 and a.CODSITUACAO in('A','F','P','Z') and a.chapa = '" & Format(txtRNC(13).Text, "00000") & "' UNION select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI as chapa,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI as nome from tbTerceirizados as a where a.chapa = '" & txtRNC(13).Text & "' and a.ativo = 'S'"
    rschamaChapa.Open SqlchamaChapa, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rschamaChapa.EOF Then
        If Mid$(txtRNC(13).Text, 1, 5) <> "CONTR" Then
            txtRNC(13).Text = Format(txtRNC(13).Text, "00000")
        End If
        txtRNC(14).Text = rschamaChapa.Fields(1)  'Nome
        CompoeControles = True
    Else
        mobjMsg.Abrir "Registro de colaborador não identificado no sistema", Ok, critico, "Atenção"
        txtRNC(13).Text = ""
        txtRNC(14).Text = "-"
        txtRNC(13).SetFocus
    End If
    rschamaChapa.Close
    Set rschamaChapa = Nothing
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

Private Sub ChamaGridColab()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "select a.chapa,a.nome from " & vBancoTotvs & ".dbo.PFUNC as a where a.CODCOLIGADA = 1 and a.CODSITUACAO in('A','F','P','Z') UNION select a.chapa COLLATE SQL_Latin1_General_CP1_CI_AI as chapa,a.nome COLLATE SQL_Latin1_General_CP1_CI_AI as nome from tbTerceirizados as a where a.ativo = 'S'"
    procnom = "chamaChapa"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de colaboradores"
    'Pesquisa = frmRNCF.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "chapa=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtRNC(13) = Pesquisa
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

Private Sub txtRNC_LostFocus(Index As Integer)
    voltaCorText txtRNC(Index)
    Select Case Index
    Case 1 Or 17
        If txtRNC(17) = "" Then
            mobjMsg.Abrir "Digite o nº da revisao da OS", Ok, critico, "Atenção"
            Exit Sub
        End If
        valida_OS
        'Abaixo: Limpa Listview e em seguida preenche novamente
        ListView1.ListItems.Clear
        chamaSQL "select a.idcc,a.nomecc,a.desenhos from tbMPItens as a where a.idos = '" & Val(txtRNC(1)) & "' and a.revisaoos = '" & Val(txtRNC(17)) & "' order by a.idoperacao"
        Compoe_Listview ListView1, Sqlp, ""
    End Select
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID CC", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nome Centro de Custo", ListView1.Width / 1.5
    ListView1.ColumnHeaders.Add , , "Desenhos", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID Causal", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Causal", ListView2.Width / 2
    ListView2.ColumnHeaders.Add , , "Grupo", ListView2.Width / 4
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlRNC = "select a.idcd,a.dataabertura as cd_abertura,a.idos,a.responsavel as cd_responsavel,f.nome,d.fce,substring(a.observacao,1,500) as cd_observacao,g.idrnc,g.dataabertura,g.responsavel,substring(g.observacao,1,500) as rnc_observacao,g.idcc,g.esbocos,g.qtdpecas,g.datareinsp,substring(g.incidente,1,500) as incidente,substring(g.correcao,1,500) as correcao,g.gerouretrabalho,substring(g.itensrnc,1,500) as itensrnc,g.dataconclusao,g.idccresponsavel,substring(g.causaraiz,1,500),substring(g.obsacao,1,500),g.tipoacao,a.revisao,a.status,g.datafechamento " & _
              "from tbComunicacaoDesvio as a left join tbMPItens as b on a.idos = b.idos left join tbMP as c on b.idprogramacao = c.idprogramacao left join tbProjetos as d on c.codprojeto = d.codprojeto left join tbFo as e on d.fce = e.fce left join tbclifor as f on e.codclifor = f.codclifor left join tbrnc as g on a.idcd = g.idcd where a.idcd = '" & Val(varGlobal) & "' " & _
              "group by a.idcd,a.dataabertura,a.idos,a.responsavel,f.nome,d.fce,substring(a.observacao,1,500),g.idrnc,g.dataabertura,g.responsavel,substring(g.observacao,1,500),g.idcc,g.esbocos,g.qtdpecas,g.datareinsp,substring(g.incidente,1,500),substring(g.correcao,1,500),g.gerouretrabalho,substring(g.itensrnc,1,500),g.dataconclusao,g.idccresponsavel,substring(g.causaraiz,1,500),substring(g.obsacao,1,500),g.tipoacao,a.revisao,a.status,g.datafechamento"
    rsRNC.Open SqlRNC, cnBanco, adOpenKeyset, adLockReadOnly
    If rsRNC.RecordCount > 0 Then
        compoeControlesForm
        
        'COMPOE TREEVIEW 2
        separaDadosText1 Text1
        mostraDesenhos "tbMPDesSel", TreeView2
        
        'COMPOE TREEVIEW 1
        Compoe_CC_RNC
        If Text2.Text <> "" Then
            EditaLVMP
        End If
        CompoeCausais
        If Not IsNull(rsRNC.Fields(18)) Then Text1 = rsRNC.Fields(18)
    End If
    rsRNC.Close
    Set rsRNC = Nothing
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
    CompoeComboCC Combo1
    txtRNC(0) = Format(rsRNC.Fields(0), "000000") 'CD Nº
    txtRNC(1) = Format(rsRNC.Fields(2), "000000") 'OS Nº
    txtRNC(2) = rsRNC.Fields(3) 'RESPONSAVEL PELA ABERTURA DA CD
    If Not IsNull(rsRNC.Fields(4)) Then txtRNC(3) = rsRNC.Fields(4)  'NOME DO CLIENTE
    If Not IsNull(rsRNC.Fields(5)) Then txtRNC(4) = rsRNC.Fields(5) 'FCE
    If Not IsNull(rsRNC.Fields(24)) Then txtRNC(17) = rsRNC.Fields(24) ' nº da revisão da OS
    
    txtRNC(6) = rsRNC.Fields(6) 'OBSERVAÇÃO DA CD
    
    If Not IsNull(rsRNC.Fields(7)) Then
        txtRNC(7) = Format(rsRNC.Fields(7), "000000") 'RNC Nº
    Else
        txtRNC(7).Text = Format(GeraCodigoTB("tbrnc", "idrnc", "", ""), "000000")
    End If
    
    
    If Not IsNull(rsRNC.Fields(9)) Then
        txtRNC(13) = Mid$(rsRNC.Fields(9), 1, 5)
        txtRNC(14) = Mid$(rsRNC.Fields(9), 9, 100)
    End If
    
    If Not IsNull(rsRNC.Fields(11)) Then Text2.Text = rsRNC.Fields(11) 'CÓDIGO DO CENTRO DE CUSTO
    
    If Not IsNull(rsRNC.Fields(10)) Then txtRNC(8) = rsRNC.Fields(10) 'OBSERVAÇÃO DA RNC
    If Not IsNull(rsRNC.Fields(12)) Then txtRNC(9) = rsRNC.Fields(12) 'ESBOÇO Nº
    If Not IsNull(rsRNC.Fields(13)) Then txtRNC(10) = rsRNC.Fields(13) 'QUANTIDADE DE PEÇAS
    If Not IsNull(rsRNC.Fields(15)) Then txtRNC(11) = rsRNC.Fields(15) 'INCIDENTE DADOS
    If Not IsNull(rsRNC.Fields(16)) Then txtRNC(12) = rsRNC.Fields(16) 'CORREÇÃO DADOS
    If Not IsNull(rsRNC.Fields(1)) Then DTPicker1 = rsRNC.Fields(1) 'DATA DA ABERTURA DA CD
    If Not IsNull(rsRNC.Fields(8)) Then DTPicker2 = rsRNC.Fields(8) 'DATA DA ABERTURA DA RNC
    If Not IsNull(rsRNC.Fields(14)) Then DTPicker3 = rsRNC.Fields(14) 'DATA DA RE-INSPEÇÃO
    If rsRNC.Fields(17) = "N" Or IsNull(rsRNC.Fields(17)) Then '
        Check1.Value = 0
    Else
        Check1.Value = 1
    End If
    If Not IsNull(rsRNC.Fields(18)) Then Text1 = rsRNC.Fields(18)
    If Not IsNull(rsRNC.Fields(19)) Then DTPicker4 = rsRNC.Fields(19) 'DATA DE CONCLUSAO DA RNC
    If Not IsNull(rsRNC.Fields(20)) Then Combo1.Text = rsRNC.Fields(20) 'CENTRO DE CUSTO RESPONSÁVEL PELA RNC
    
    If Not IsNull(rsRNC.Fields(21)) Then txtRNC(15) = rsRNC.Fields(21) 'CAUSA RAIZ DETERMINADA
    If Not IsNull(rsRNC.Fields(22)) Then txtRNC(16) = rsRNC.Fields(22) 'DESCRITIVO DAS AÇÕES CORRETIVAS E/OU PREVENTIVAS
    If Not IsNull(rsRNC.Fields(23)) Then Combo2.Text = rsRNC.Fields(23) 'CLASSIFICAÇÃO DAS AÇÕES CORRETIVAS E/OU PREVENTIVAS
    If Not IsNull(rsRNC.Fields(25)) Then SkinLabel20.Caption = rsRNC.Fields(25) 'STATUS
    If Not IsNull(rsRNC.Fields(26)) Then DTPicker5 = rsRNC.Fields(26) 'DATA DE FECHAMENTO DA RNC

    vStatus = SkinLabel20.Caption
    If SkinLabel20.Caption = 7 Then 'AZUL
        Picture1.BackColor = &HC00000
    ElseIf SkinLabel20.Caption = 20 Then 'AMARELO
        Picture1.BackColor = &HC0C0&
    ElseIf SkinLabel20.Caption = 8 Then 'VERDE
        Picture1.BackColor = &H8000&
    ElseIf SkinLabel20.Caption = 9 Then 'LARANJA
        Picture1.BackColor = &H80FF&
    ElseIf SkinLabel20.Caption = 10 Then 'PRETO
        Picture1.BackColor = &H0&
    Else ' VERMELHO
        Picture1.BackColor = &HFF&
    End If
    
    chamaSQL "select a.idcc,a.nomecc,a.desenhos from tbMPItens as a where idos = '" & Val(txtRNC(1)) & "' and revisaoos = '" & Val(txtRNC(17)) & "' order by a.idoperacao"
    Compoe_Listview ListView1, Sqlp, ""
    
    chamaSQL "Select a.idcausal,a.nomecausal,b.nomegrupocausal from tbCausais as a inner join tbCausaisGrupos as b on a.idgrupocausal = b.idgrupocausal order by a.idgrupocausal,a.idcausal"
    Compoe_Listview ListView2, Sqlp, "00"
    
End Sub

Private Sub ListView1_Click()
    Text1.Text = ""
    desmarcaCC
    EditaLVMP
End Sub

Private Sub EditaLVMP()
    AlteraLV ListView1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1, Text1
    separaDadosText1 Text1
    mostraDesenhos "tbMPDesSel", TreeView1
End Sub

'A função abaixo separa os valores do texbox TEXT1 e grava na tabela tbMPDesSel
Private Sub separaDadosText1(vTxtForm As TextBox)
On Error GoTo Err
    Dim rsTransf As New ADODB.Recordset
    Dim SqlTransf As String
    Dim vCodLM As String, vCodSeq As String
    
    SqlTransf = "Delete from tbMPDesSel where fce = '" & Val(txtRNC(4)) & "'"
    rsTransf.Open SqlTransf, cnBanco
    
    Dim RECEBE As String
    Dim Contador As Integer, x As Integer
    Contador = 0
    For x = 1 To Len(vTxtForm)
        If Mid(vTxtForm, x, 1) = ";" Then
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
            
            SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtRNC(4)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, x, 1)
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
        SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtRNC(4)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
        rsTransf.Open SqlTransf, cnBanco
    End If
    achaProjeto vCodLM, vCodSeq
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        criaTabela
    End If
End Sub

Private Sub criaTabela()
On Error GoTo Err
    cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbMPDesSel(" & _
    "fce NUMERIC NOT NULL," & _
    "codlm NUMERIC NOT NULL," & _
    "codseq NUMERIC NOT NULL)"
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

Private Sub achaProjeto(vIDLM As String, vIDSeq As String)
On Error GoTo Err
    Dim rsProjeto As New ADODB.Recordset
    Dim SqlProjeto As String
    SqlProjeto = "select c.projeto from tbItemLM as a inner join tbdesenhos as b on a.codigodes = b.iddesenho inner join tbProjetos as c on b.codprojeto = c.codprojeto where c.fce = '" & txtRNC(4).Text & "' and a.codlm = '" & Val(vIDLM) & "' and a.codseq = '" & Val(vIDSeq) & "'"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsProjeto.EOF Then txtRNC(5) = rsProjeto.Fields(0)
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

    If vTabela = "tbMPDesSel" Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(h.idos) as OS " & _
        "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq inner join tbProjetos as g on g.codprojeto = c.codprojeto left join tbositens as h on a.fce = h.fce and a.codlm = h.codlm and a.codseq = h.codseq Where a.fce = '" & Val(txtRNC(4)) & "'" & _
        "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    
    End If
    
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
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
        If vItens = "" Then
            If TV.Name = "TreeView2" Then vItens = vNome1
        End If
        'TESTE DE COR --------------
        If Not IsNull(rsTreeview.Fields(13)) Then
            nd.ForeColor = &H8000&
        End If
        '----------------------------
        
        Do While Mid$(vNome1, 1, Len(vNome1) - 1) = Mid$(vNomeA, 1, Len(vNome1) - 1) And Not rsTreeview.EOF
            If vNomeB <> "" Then
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
                'SEGUNDO NO
                Set nd = TV.Nodes.Add(vNome1, tvwChild, vNomeNo, vNome2)
                If TV.Name = "TreeView2" Then vItens = vItens & "/" & vNome2
                'TESTE DE COR --------------
                If Not IsNull(rsTreeview.Fields(13)) Then
                    nd.ForeColor = &H8000&
                End If
                '----------------------------
                
                'TEORICAMENTE INICIALIZAÇÃO DA VARIAVEL QUE RECEBERA O SOMATORIO DO VALOR DA POSIÇÃO DO DESENHO IRÁ
                'FICAR NESSE LOCAL
                vPesoPosicao = 0
                
                'Teste
                If Mid$(right(vNome2, 14), 1, 3) = "OS:" Then
                    Dim vTamanho1 As Integer
                    vTamanho1 = Len(vNome2) - 11
                    vNome2 = Mid$(vNome2, 1, vTamanho1) & ")"
                End If
                
                Do While Mid$(vNome1, 1, Len(vNome1) - 1) = Mid$(vNomeA, 1, Len(vNome1) - 1) And Mid(vNome2, 1, Len(vNome2) - 1) = Mid$(vNomeB, 1, Len(vNome2) - 1) And vNomeC <> "" And Not rsTreeview.EOF
                    'TERCEIRO NO
                    'OBS: OS VALORES DOS NOs NÃO PODEM SE REPETIR
                    'FOI ADICIONADO UM CONTADOR AO IDENTIFICADOR DO NO PARA QUE ELE NÃO SE REPITA
                    If TV.Name = "TreeView2" Then
                        
                        'Abaixo é calculado o peso de cada posicao de cada desenho e realizado a classificação
                        'dentro da formula
                        vPesoPosicao = vPesoPosicao + (rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        
                        vPesoTotal2 = vPesoTotal2 + (rsTreeview.Fields(6) * rsTreeview.Fields(7) * rsTreeview.Fields(9))
                        If Text1.Text = "" Then
                            Text1 = Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                        Else
                            Text1 = Text1.Text & ";" & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                        End If
                    End If
                    Set nd = TV.Nodes.Add(vNomeNo, tvwChild, vNomeC & vNo, vNomeC)
                    'vItens = vItens & "/" & vNomeC

                    If Not IsNull(rsTreeview.Fields(13)) Then
                        nd.ForeColor = &H8000&
                    End If
                    '----------------------------
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & " - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                    separaDadosTree vJuntaNome
                    vPula = 1
                Loop
                
                vNo = vNo + 1
                vNomeNo = vNome2 & vNo
            End If
            If vPula = 0 Then
                If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
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

Private Sub separaDadosTree(vTxtForm As String)
    Dim RECEBE As String
    Dim Contador As Integer, x As Integer
    Contador = 0
    vNomeA = ""
    vNomeB = ""
    vNomeC = ""
    For x = 1 To Len(vTxtForm)
        If Mid(vTxtForm, x, 1) = ";" Then
            If Contador = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            If Contador = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
            Contador = Contador + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(vTxtForm, x, 1)
        End If
    Next
    If RECEBE <> "" Then
        If Contador = 0 Then vNomeA = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 1 Then vNomeB = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
        If Contador = 2 Then vNomeC = RECEBE 'Variavel vGrupo recebe o valor do primeiro parâmetro
    End If
End Sub

Private Sub desmarcaCC()
    Dim x As Integer, y As Integer, j As Integer
    y = ListView1.ListItems.Count
    If y = 0 Then Exit Sub
    j = ListView1.SelectedItem.Index
    For x = 1 To y
        If ListView1.ListItems.Item(x).Checked = True Then
            ListView1.ListItems.Item(x).Checked = False
        End If
    Next
    ListView1.ListItems.Item(j).Checked = True
    Text2 = ListView1.ListItems.Item(j)
End Sub

Private Sub buscaChecado2(vLV As TreeView)
    Dim x As Integer, vContador As Integer, vQtdNos As Integer
    vContador = 0
    x = 0
    vQtdNos = vLV.Nodes.Count
    For x = 1 To vQtdNos
        If vLV.Nodes.Item(x).Checked = True Then
            transfDesenhosSel x, vLV
        End If
    Next
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
    
    If Mid$(right(vNomeC, 6), 1, 1) = " " Then
        vNomeC = right(vNomeC, 5)
        vCodLM = Mid$(vNomeC, 1, 2)
        vCodSeq = Mid$(vNomeC, 3, 3)
    Else
        vNomeC = right(vNomeC, 6)
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
            SqlTransf = "Insert into tbMPDesSel(fce,codlm,codseq) Values('" & Val(txtRNC(4)) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "')"
            rsTransf.Open SqlTransf, cnBanco
        End If
    ElseIf vTV.Name = "TreeView2" Then
        SqlTransf = "Delete from tbMPDesSel where fce = '" & Val(txtRNC(4)) & "' and codlm = '" & Val(vCodLM) & "' and codseq = '" & Val(vCodSeq) & "'"
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

Private Function salvar_Dados()
On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
        
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
        
        
    salvar_Dados = True
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(1, 1) = txtRNC(7).Text 'IDRNC
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = txtRNC(0).Text 'IDCD
    vQualquerDado(2, 2) = "I"
    vQualquerDado(3, 1) = DTPicker2.Value 'Data
    vQualquerDado(3, 2) = "D"
    vQualquerDado(4, 1) = txtRNC(13) & " - " & txtRNC(14).Text 'Responsável
    vQualquerDado(4, 2) = "S"
    vQualquerDado(5, 1) = txtRNC(8).Text 'Observação
    vQualquerDado(5, 2) = "S"
    
    If MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon <> "FABRICANDO" Or MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon <> "PRETO" Then
        If vStatus <> 20 And vStatus <> 10 Then
            vQualquerDado(6, 1) = vStatus 'Status
        Else
'            If vStatus <> 10 Then
'                vQualquerDado(6, 1) = vStatus 'Status
'            End If
        End If
    End If
    vQualquerDado(6, 2) = "I"
    vQualquerDado(7, 1) = Text2.Text 'Centro de Custo selecionado
    vQualquerDado(7, 2) = "S"
    vQualquerDado(8, 1) = txtRNC(9).Text 'Esboços
    vQualquerDado(8, 2) = "S"
    vQualquerDado(9, 1) = txtRNC(10).Text 'Quantidade de peças
    vQualquerDado(9, 2) = "I"
    If DTPicker3.Value <> "" Then
        vQualquerDado(10, 1) = DTPicker3 'Data da re-inspeção
        vQualquerDado(10, 2) = "D"
    End If
    vQualquerDado(11, 1) = txtRNC(11).Text 'Incidente
    vQualquerDado(11, 2) = "S"
    vQualquerDado(12, 1) = txtRNC(12).Text 'Correção
    vQualquerDado(12, 2) = "S"
    
    If Check1.Value = 1 Then 'Gerou retrabalho? (1)Gerou / (0) Não gerou
        vQualquerDado(13, 1) = "S"
    Else
        vQualquerDado(13, 1) = "N"
    End If
    vQualquerDado(13, 2) = "S"
    vQualquerDado(14, 1) = Text1.Text 'Itens RNC
    vQualquerDado(14, 2) = "S"
    If DTPicker4.Value <> "" Then
        vQualquerDado(15, 1) = DTPicker2.Value 'Data Conclusão
        vQualquerDado(15, 2) = "D"
    End If
    
    vQualquerDado(16, 1) = Combo1.Text 'Centro de Custo responsável pela RNC
    vQualquerDado(16, 2) = "S"
    
    vQualquerDado(17, 1) = txtRNC(15).Text 'Causa Raiz Determinada
    vQualquerDado(17, 2) = "S"
    vQualquerDado(18, 1) = txtRNC(16).Text 'Tipo de Ações - Observação
    vQualquerDado(18, 2) = "S"
    
    vQualquerDado(19, 1) = Combo2.Text 'Tipo de Ações - Seleção
    vQualquerDado(19, 2) = "S"
    
    If DTPicker5.Value <> "" Then
        vQualquerDado(20, 1) = DTPicker5.Value 'Data Fechamento
        vQualquerDado(20, 2) = "D"
    End If
    
    GravaDados "tbRNC", "idrnc", "I", txtRNC(7), 20, "", "", txtRNC(7)
        
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(4, 1) = txtRNC(1).Text 'OS nº - Correção
    vQualquerDado(4, 2) = "I"
    vQualquerDado(6, 1) = vStatus 'status IDCD
    vQualquerDado(6, 2) = "I"
    vQualquerDado(7, 1) = txtRNC(17).Text 'Revisão OS - Correção
    vQualquerDado(7, 2) = "I"
    GravaDados "tbComunicacaoDesvio", "idcd", "I", txtRNC(0), 6, "", "", txtRNC(0)
    
    'Grava dados ListView2
    sqlDeletar = "Delete from tbRNCCausais where idrnc = '" & Val(txtRNC(7)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbRNCCausais where idrnc = '" & Val(txtRNC(7)) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    If ListView2.ListItems.Count > 0 Then
        For x = 1 To ListView2.ListItems.Count
            ListView2.ListItems.Item(x).Selected = True
            If ListView2.ListItems.Item(x).Checked = True Then
                rsSalvar.AddNew
                rsSalvar.Fields(0) = Val(txtRNC(7))
                rsSalvar.Fields(1) = Val(ListView2.ListItems.Item(x))
                rsSalvar.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(1)
            End If
        Next
    End If
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        salvar_Dados = False
        Exit Function
    End If
End Function

Private Sub Compoe_CC_RNC()
    Dim ItemLst As ListItem
    Dim x As Integer, y As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        ListView1.ListItems(x).Selected = True
        If ListView1.ListItems.Item(x) = Text2.Text Then
            ListView1.ListItems.Item(x).Checked = True
            ListView1.ListItems.Item(x).Selected = True
        End If
    Next
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub CompoeCausais()
On Error GoTo Err
    Dim rsCausais As New ADODB.Recordset
    Dim SqlCausais As String
    SqlCausais = "Select * from tbRNCCausais where idrnc = '" & Val(txtRNC(7)) & "'"
    rsCausais.Open SqlCausais, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim x As Integer, y As Integer
    If rsCausais.RecordCount = 0 Then Exit Sub
    y = ListView2.ListItems.Count
    While Not rsCausais.EOF
        For x = 1 To y
            ListView2.ListItems(x).Selected = True
            If Val(ListView2.ListItems.Item(x)) = rsCausais.Fields(1) Then
                ListView2.ListItems.Item(x).Checked = True
            End If
        Next
        rsCausais.MoveNext
    Wend
    rsCausais.Close
    Set rsCausais = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
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

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Dim y As Integer, x As Integer
    y = MeuLV.ListView1.ListItems.Count
    For x = 1 To y
        If MeuLV.ListView1.ListItems.Item(x).Selected = True Then
            Exit For
        End If
    Next
    If vStatus = 7 Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "AVALIANDO1"
    ElseIf vStatus = 8 Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "CONCLUIDO1"
    ElseIf vStatus = 20 Then
        If MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon <> "FABRICANDO" Or MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon <> "PRETO" Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ReportIcon = "FECHADO"
        End If
    End If
    If Check1.Value = 1 Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(10).ReportIcon = "OK"
    ElseIf Check1.Value = 0 Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(10) = ""
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(10).ReportIcon = "EXC"
    End If
    
    If txtRNC(7).Text <> "" Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = txtRNC(7).Text
        'MeuLV.ListView1.SelectedItem.ListSubItems.Item(8).ReportIcon = "EXC"
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "ZEUS"
    Exit Sub
End Sub

Private Function verificaDados(vCF As Integer) 'C - Concluido / F - Fechado
    verificaDados = False
    '1º passo - Verificar se o campo "Responsável" esta preenchido
    If txtRNC(13).Text = "" Then
        mobjMsg.Abrir "O campo " & txtRNC(13).Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
        txtRNC(13).SetFocus
        Exit Function
    End If
    
    '2º passo - Verificar se o campo "Incidente" esta preenchido
    If txtRNC(11).Text = "" Then
        mobjMsg.Abrir "O campo " & txtRNC(11).Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
        SSTab1.Tab = 2
        txtRNC(11).SetFocus
        Exit Function
    End If
    
    '3º passo - Verificar se o campo "Correção" esta preenchido
    If txtRNC(12).Text = "" Then
        mobjMsg.Abrir "O campo " & txtRNC(12).Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
        SSTab1.Tab = 2
        txtRNC(12).SetFocus
        Exit Function
    End If
    
    '4º passo - Verificar se o ComboBox do CC - Centro de Custo responsável pela não conformidade foi informado
    If Combo1.Text = "" Then
        mobjMsg.Abrir "O campo " & Combo1.Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
        SSTab1.Tab = 0
        Combo1.SetFocus
        Exit Function
    End If
    
    '5º passo - Verificar se os campos "Causa Raiz" e "Ações Corretivas e/ou Preventivas" estão preenchidos
    'Esse passo somente será executado se a "Data de Fechamento" da RNC estiver marcada. vCF = 20
    If vCF = 20 Then
        If Combo2.Text = "" Then
            mobjMsg.Abrir "O campo " & Combo2.Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
            SSTab1.Tab = 2
            Combo2.SetFocus
            Exit Function
        End If
        If txtRNC(15).Text = "" Then
            mobjMsg.Abrir "O campo " & txtRNC(15).Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
            SSTab1.Tab = 2
            txtRNC(15).SetFocus
            Exit Function
        End If
        If txtRNC(16).Text = "" Then
            mobjMsg.Abrir "O campo " & txtRNC(16).Tag & " não foi preenchido", Ok, exclamacao, "ZEUS"
            SSTab1.Tab = 2
            txtRNC(16).SetFocus
            Exit Function
        End If
    End If
    
    '6º passo - Verificar se pelo menos 1 causal foi selecionado
    Dim x As Integer, y As Integer
    If ListView2.ListItems.Count > 0 Then
        y = ListView2.ListItems.Count
        For x = 1 To y
            ListView2.ListItems.Item(x).Selected = True
            If ListView2.ListItems.Item(x).Checked = True Then
                verificaDados = True
                Exit Function
            End If
        Next
        verificaDados = False
        mobjMsg.Abrir "Nenhuma causal foi selecionada", Ok, exclamacao, "ZEUS"
        SSTab1.Tab = 0
        ListView2.SetFocus
        Exit Function
    End If
    verificaDados = True
End Function

Private Sub bloqueiaEdicao()
    Dim x As Integer
    ListView1.Enabled = False
    ListView2.Enabled = False
    TreeView1.Enabled = False
    cmdCadastro(7).Enabled = False
    cmdCadastro(8).Enabled = False
    cmdCadastro(9).Enabled = False
    cmdCadastro(12).Enabled = False
    cmdCadastro(13).Enabled = True
    
    For x = 0 To txtRNC.Count - 1
        txtRNC(x).Enabled = False
    Next
    
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DTPicker3.Enabled = False
    DTPicker4.Enabled = False
    Check1.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    If DTPicker5.Value <> "" Then
        DTPicker5.Enabled = False
        txtRNC(15).Enabled = False
        txtRNC(16).Enabled = False
        Combo2.Enabled = False
        cmdCadastro(12).Enabled = False
    Else
        DTPicker5.Enabled = True
        txtRNC(15).Enabled = True
        txtRNC(16).Enabled = True
        Combo2.Enabled = True
        cmdCadastro(12).Enabled = True
    End If
End Sub

Private Sub enviaEmail()
'PRECISA INCLUIR NO PROJETO A DLL MICROSOFT CDO FOR WINDOWS 2000 LIBRARY
'DICA: CRIA O DOCUMENTO NO WORD, COPIA, ABRE O OUTLOOK, CRIE UM NOVO EMAIL, COLE E SALVE COMO HTML
'VC IRÁ TRABALHAR COM O ARQUIVO HTML GERADO.
'SE O ARQUIVO FOR MUITO GRANDE, CRIE VÁRIAS VARIÁVEIS DO TIPO STRING PARA ARMAZENAR O ARQUIVO PICADO
'QUANDO FOR ENVIAR CONCATENE TODAS AS VARIAVEIS
On Error GoTo errMail
    Dim vCorDecisao As String
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    
    Dim vHtml1 As String, vHtml2 As String, vHtml3 As String, vHtml4 As String, vHtml5 As String, vHtml6 As String, vHtml7 As String, vHtml8 As String, vHtml9 As String, vHtml10 As String, vHtml11 As String, vHtml12 As String
    montaAgrupados
vHtml1 = "<META HTTP-EQUIV='Content-Type' CONTENT='text/html;charset=iso-8859-1'>" & _
        "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & _
        "<HTML><HEAD>" & _
        "<META content='text/html; charset=iso-8859-1' http-equiv=Content-Type>" & _
        "<META name=GENERATOR content='MSHTML 8.00.6001.18702'>" & _
        "<STYLE></STYLE>" & _
        "</HEAD>" & _
        "<BODY bgColor=#ffffff>" & _
        "<BLOCKQUOTE style='MARGIN-RIGHT: 0px' dir=ltr>" & _
        "  <DIV>" & _
        "  <TABLE " & _
        "  style='BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -8.8pt; WIDTH: 538.7pt; BORDER-COLLAPSE: collapse; BORDER-TOP: medium none; BORDER-RIGHT: medium none; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt; mso-border-insideh: .5pt solid windowtext; mso-border-insidev: .5pt solid windowtext' " & _
        "  class=MsoNormalTable border=1 cellSpacing=0 cellPadding=0 width=718>" & _
        "    <TBODY>" & _
        "    <TR style='mso-yfti-irow: 0; mso-yfti-firstrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-bottom-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718 colSpan=2>" & _
        "        <P style='TEXT-ALIGN: center; MARGIN: 0cm 0cm 0pt' class=MsoNormal " & _
        "        align=center><B style='mso-bidi-font-weight: normal'><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: black; FONT-SIZE: 14pt; mso-bidi-font-size: 12.0pt'>COMUNICAÇÃO " & _
        "        DE RETRABALHO<?xml:namespace prefix = o ns = " & _
        "        'urn:schemas-microsoft-com:office:office' " & _
        "    /><o:p></o:p></SPAN></B></P></TD></TR>"

vHtml2 = "    <TR style='mso-yfti-irow: 1'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718 colSpan=2>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><B " & _
        "        style='mso-bidi-font-weight: normal'><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'>Dados da CD – Comunicação de " & _
        "        Desvio<SPAN " & _
        "    style='COLOR: #548dd4'><o:p></o:p></SPAN></SPAN></B></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 2'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 227.05pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt dotted; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=303>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>CD nº: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(0).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: #d4d0c8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 311.65pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: dotted windowtext .5pt; mso-border-right-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=416>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Data "
vHtml3 = "        abertura: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & DTPicker1.Value & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 3'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 227.05pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt dotted; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=303>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>OS nº: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(1).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: #d4d0c8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 311.65pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: dotted windowtext .5pt; mso-border-right-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=416>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Responsável: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & Mid$(txtRNC(2).Text, 9, 23) & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 4'>"
vHtml4 = "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 227.05pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt dotted; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=303>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Cliente: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(3).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: #d4d0c8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 311.65pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: dotted windowtext .5pt; mso-border-right-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=416>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>FCE: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(4).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 5'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718 colSpan=2>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Projeto: </SPAN><SPAN "
vHtml5 = "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(5).Text & "</SPAN><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='HEIGHT: 2cm; mso-yfti-irow: 6; mso-yfti-lastrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; HEIGHT: 2cm; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718 colSpan=2>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Observação: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(6).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR></TBODY></TABLE>" & _
        "  <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "  style='FONT-FAMILY: Calibri,sans-serif'><o:p>&nbsp;</o:p></SPAN></P>" & _
        "  <TABLE " & _
        "  style='BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -8.8pt; WIDTH: 538.7pt; BORDER-COLLAPSE: collapse; BORDER-TOP: medium none; BORDER-RIGHT: medium none; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt; mso-border-insideh: .5pt solid windowtext; mso-border-insidev: .5pt solid windowtext' " & _
        "  class=MsoNormalTable border=1 cellSpacing=0 cellPadding=0 width=718>" & _
        "    <TBODY>" & _
        "    <TR style='mso-yfti-irow: 0; mso-yfti-firstrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-bottom-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718 colSpan=3>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><B "
vHtml6 = "        style='mso-bidi-font-weight: normal'><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'>Dados do RNC – Registro de " & _
        "        Não Conformidade<SPAN " & _
        "        style='COLOR: #548dd4'><o:p></o:p></SPAN></SPAN></B></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 1'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 163.05pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt dotted; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=217>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>RNC nº: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(7).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: #d4d0c8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 177.2pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt dotted; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=236>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Abertura: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & DTPicker2.Value & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD>" & _
        "      <TD "
vHtml7 = "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: #d4d0c8; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 7cm; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: dotted windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt; mso-border-left-alt: dotted windowtext .5pt; mso-border-right-alt: solid windowtext .5pt' " & _
        "      vAlign=top width=265>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Fechamento: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & DTPicker4.Value & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 2'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718 colSpan=3>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Responsável: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(14).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='HEIGHT: 2cm; mso-yfti-irow: 3; mso-yfti-lastrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; HEIGHT: 2cm; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718 colSpan=3>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Observação: "
vHtml8 = "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(8).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR></TBODY></TABLE>" & _
        "  <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "  style='FONT-FAMILY: Calibri,sans-serif'><o:p>&nbsp;</o:p></SPAN></P>" & _
        "  <TABLE " & _
        "  style='BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -8.8pt; WIDTH: 538.7pt; BORDER-COLLAPSE: collapse; BORDER-TOP: medium none; BORDER-RIGHT: medium none; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt; mso-border-insideh: .5pt dotted windowtext; mso-border-insidev: .5pt dotted windowtext' " & _
        "  class=MsoNormalTable border=1 cellSpacing=0 cellPadding=0 width=718>" & _
        "    <TBODY>" & _
        "    <TR style='mso-yfti-irow: 0; mso-yfti-firstrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-bottom-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><B " & _
        "        style='mso-bidi-font-weight: normal'><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'>Dados da Identificação da " & _
        "        Não Conformidade<SPAN " & _
        "        style='COLOR: #548dd4'><o:p></o:p></SPAN></SPAN></B></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 1'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN "
vHtml9 = "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Centro de " & _
        "        Custo onde ocorreu a Não Conformidade: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & vCCNConforme & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 2'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Centro de " & _
        "        Custo responsável pela Não Conformidade: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & Combo1.Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 3'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Causais: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & vCausais & "</SPAN><SPAN "
vHtml10 = "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 4'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Itens: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & vItens & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 5'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Esboço nº: " & _
        "        </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(9).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 6; mso-yfti-lastrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt' "
vHtml11 = "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Qtd. Peças a " & _
        "        serem retrabalhadas: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(10).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR></TBODY></TABLE>" & _
        "  <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "  style='FONT-FAMILY: Calibri,sans-serif'><o:p>&nbsp;</o:p></SPAN></P>" & _
        "  <TABLE " & _
        "  style='BORDER-BOTTOM: medium none; BORDER-LEFT: medium none; MARGIN: auto auto auto -8.8pt; WIDTH: 538.7pt; BORDER-COLLAPSE: collapse; BORDER-TOP: medium none; BORDER-RIGHT: medium none; mso-border-alt: solid windowtext .5pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 5.4pt 0cm 5.4pt; mso-border-insideh: .5pt dotted windowtext; mso-border-insidev: .5pt dotted windowtext' " & _
        "  class=MsoNormalTable border=1 cellSpacing=0 cellPadding=0 width=718>" & _
        "    <TBODY>" & _
        "    <TR style='mso-yfti-irow: 0; mso-yfti-firstrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-bottom-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><B " & _
        "        style='mso-bidi-font-weight: normal'><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'>Plano de ação<SPAN " & _
        "        style='COLOR: #548dd4'><o:p></o:p></SPAN></SPAN></B></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 1'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt dotted; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-bottom-alt: dotted; mso-border-top-alt: dotted; mso-border-left-alt: solid; mso-border-right-alt: solid; mso-border-color-alt: windowtext; mso-border-width-alt: .5pt' "
vHtml12 = "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Incidente " & _
        "        ocorrido: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #c00000'>" & txtRNC(11).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR>" & _
        "    <TR style='mso-yfti-irow: 2; mso-yfti-lastrow: yes'>" & _
        "      <TD " & _
        "      style='BORDER-BOTTOM: windowtext 1pt solid; BORDER-LEFT: windowtext 1pt solid; PADDING-BOTTOM: 0cm; BACKGROUND-COLOR: transparent; PADDING-LEFT: 5.4pt; WIDTH: 538.7pt; PADDING-RIGHT: 5.4pt; BORDER-TOP: #d4d0c8; BORDER-RIGHT: windowtext 1pt solid; PADDING-TOP: 0cm; mso-border-alt: solid windowtext .5pt; mso-border-top-alt: dotted windowtext .5pt' " & _
        "      vAlign=top width=718>" & _
        "        <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif; COLOR: #548dd4'>Correção " & _
        "        apresentada: </SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri; COLOR: #c00000'>" & txtRNC(12).Text & "</SPAN><SPAN " & _
        "        style='FONT-FAMILY: Calibri,sans-serif'><o:p></o:p></SPAN></P></TD></TR></TBODY></TABLE>" & _
        "  <P style='MARGIN: 0cm 0cm 0pt' class=MsoNormal><SPAN " & _
        "  style='FONT-FAMILY: Calibri,sans-serif'><o:p>&nbsp;</o:p></SPAN></P></DIV></BLOCKQUOTE></BODY></HTML>"

    
    vDecisao = "Aprovado"
    vCorDecisao = "#228B22"

    'vSMTP = "smtp.viga.ind.br"
    'vUsuEmail = "taos@viga.ind.br"
    'vSenhaEmail = "taos2017@"

    With Camp
        .Item(cdoSendUsingMethod) = 2   ' cdoSendUsingPort
        .Item(cdoSMTPServer) = vSMTP  '"smtp.mail.yahoo.com.br"   ‘informe o servidor smtp aqui
        .Item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
        .Item(cdoSMTPAuthenticate) = 1
        .Item(cdoSendUserName) = vUsuEmail ' informe o usuario de autenticação
        .Item(cdoSendPassword) = vSenhaEmail  'Informe a Senha aqui
        .Update
    End With

    With Msg
        Set .Configuration = Cof
        .To = sEmailRNC
'        .To = "planejamento3@viga.ind.br;planejamento4@viga.ind.br;viga@viga.ind.br;planejamento5@viga.ind.br;planejamento7@viga.ind.br;planejamento8@viga.ind.br" 'destinatarios separados por ;
        .From = "viga@viga.ind.br"  '"contatos@flowsys.com.br"   'remetente@email.com.br ‘ remetente"
        .Subject = "RNC - Registro de Não Conformidade: " & txtRNC(0)
        .HTMLBody = vHtml1 & vHtml2 & vHtml3 & vHtml4 & vHtml5 & vHtml6 & vHtml7 & vHtml8 & vHtml9 & vHtml10 & vHtml11 & vHtml12
        .Send
    End With
    mobjMsg.Abrir "Email enviado com sucesso!", Ok, informacao, "ZEUS"
    Exit Sub
errMail:
    Msgbox "Email não enviado para o usuário solicitante" & vbCrLf & vbCrLf & _
    "ERRO de autenticação! Favor verificar se as configurações de SMTP e email estão corretas." & vbCrLf & _
    "Reporte o ERRO ao administrador do sistema.", vbCritical, "Zeus"
    Exit Sub
End Sub

Private Sub montaAgrupados()
On Error GoTo Err
    Dim rsCausal As New ADODB.Recordset
    Dim SqlCausal As String
    SqlCausal = "select a.idcausal,a.nomecausal from tbRNCCausais as a where a.idrnc ='" & Val(txtRNC(7).Text) & "'"
    rsCausal.Open SqlCausal, cnBanco, adOpenKeyset, adLockReadOnly
    vCausais = ""
    If rsCausal.RecordCount > 0 Then
        While Not rsCausal.EOF
            vCausais = vCausais & rsCausal.Fields(0) & "-" & rsCausal.Fields(1) & "/"
            rsCausal.MoveNext
        Wend
    End If
    rsCausal.Close
    Set rsCausal = Nothing
    
    vCCNConforme = Text2.Text
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

Private Sub valida_OS()
On Error GoTo Err
    Dim rsValidaOS As New ADODB.Recordset
    Dim sqlValidaOS As String
    If IsNumeric(txtRNC(1)) Then
        sqlValidaOS = "SELECT IDOS,RASTREABILIDADE,OBSERVACAO,DATAOS,REVISAO,STATUS FROM TBOS where idos = '" & Val(txtRNC(1)) & "' and revisao = '" & Val(txtRNC(17)) & "' and status < 3"
        rsValidaOS.Open sqlValidaOS, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    
    If rsValidaOS.RecordCount = 0 Then
        SkinLabel18.Visible = True
        SkinLabel18.Caption = "A OS informada não é válida ou já esta fechada"
    Else
        SkinLabel18.Visible = False
        txtRNC(1).Text = Format(txtRNC(1).Text, "000000000")
    End If
    rsValidaOS.Close
    Set rsValidaOS = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        SkinLabel18.Visible = True
        SkinLabel18.Caption = "A OS informada não é válida ou já esta fechada"
        rsValidaOS.Close
        Set rsValidaOS = Nothing
    End If
End Sub
