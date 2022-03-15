VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFO 
   BackColor       =   &H00B7B7B7&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15525
   Icon            =   "frmFO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   15525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton chamCad 
      Height          =   615
      Index           =   6
      Left            =   1320
      Picture         =   "frmFO.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   313
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8880
      Width           =   615
   End
   Begin VB.CommandButton chamCad 
      Height          =   615
      Index           =   5
      Left            =   720
      Picture         =   "frmFO.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   315
      Tag             =   "Salvar"
      ToolTipText     =   "Salvar"
      Top             =   8880
      Width           =   615
   End
   Begin VB.CommandButton chamCad 
      Height          =   615
      Index           =   4
      Left            =   120
      Picture         =   "frmFO.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   314
      Tag             =   "Salvar"
      ToolTipText     =   "Salvar"
      Top             =   8880
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14160
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   307
      Top             =   8880
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   15266
      _Version        =   393216
      Tabs            =   9
      Tab             =   2
      TabsPerRow      =   9
      TabHeight       =   520
      BackColor       =   12040119
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "FO"
      TabPicture(0)   =   "frmFO.frx":3328
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Imposto / Serviços"
      TabPicture(1)   =   "frmFO.frx":3344
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Matéria Prima"
      TabPicture(2)   =   "frmFO.frx":3360
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ListView11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "SkinLabel40"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "ScriptControl1"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame9(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame14"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtLvw"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame17"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame15"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame20"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Frame16"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "chamCad(0)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "chamCad(1)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "chamCad(7)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Resumo-MP"
      TabPicture(3)   =   "frmFO.frx":337C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ListView12"
      Tab(3).Control(1)=   "SkinLabel47"
      Tab(3).Control(2)=   "SkinLabel46"
      Tab(3).Control(3)=   "txtcadastro(78)"
      Tab(3).Control(4)=   "txtcadastro(79)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Despesas/Créditos"
      TabPicture(4)   =   "frmFO.frx":3398
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3(1)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Pintura"
      TabPicture(5)   =   "frmFO.frx":33B4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "SSTab4"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Testes e Ensaios"
      TabPicture(6)   =   "frmFO.frx":33D0
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame8(1)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Transportes"
      TabPicture(7)   =   "frmFO.frx":33EC
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "SSTab5"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Tintas"
      TabPicture(8)   =   "frmFO.frx":3408
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "SSTab6"
      Tab(8).Control(1)=   "Frame13"
      Tab(8).ControlCount=   2
      Begin VB.CommandButton chamCad 
         Height          =   615
         Index           =   7
         Left            =   1320
         Picture         =   "frmFO.frx":3424
         Style           =   1  'Graphical
         TabIndex        =   311
         Tag             =   "Novo Grupo"
         ToolTipText     =   "Novo Grupo"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton chamCad 
         Height          =   615
         Index           =   1
         Left            =   720
         Picture         =   "frmFO.frx":40EE
         Style           =   1  'Graphical
         TabIndex        =   310
         Tag             =   "Excluir Grupo"
         ToolTipText     =   "Excluir Grupo"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton chamCad 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmFO.frx":4DB8
         Style           =   1  'Graphical
         TabIndex        =   312
         Tag             =   "Incluir Grupo"
         ToolTipText     =   "Incluir Grupo"
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox txtcadastro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   345
         Index           =   79
         Left            =   -61680
         TabIndex        =   306
         Text            =   "x.xxx,xx"
         Top             =   8160
         Width           =   1695
      End
      Begin VB.TextBox txtcadastro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Height          =   345
         Index           =   78
         Left            =   -61680
         TabIndex        =   305
         Text            =   "x.xxx,xx"
         Top             =   7800
         Width           =   1695
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Fornecedor "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   -74760
         TabIndex        =   302
         Top             =   480
         Width           =   8415
         Begin VB.CommandButton Command1 
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
            Height          =   345
            Left            =   7680
            TabIndex        =   304
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtcadastro 
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
            Index           =   77
            Left            =   120
            TabIndex        =   303
            Tag             =   "Pintura"
            Top             =   240
            Width           =   7455
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Parâmetro/Cálculo Área de pintura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   6000
         TabIndex        =   297
         Top             =   3360
         Width           =   3255
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   75
            Left            =   120
            TabIndex        =   299
            Tag             =   "Pintura"
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pintura:"
            Height          =   255
            Left            =   120
            TabIndex        =   298
            Top             =   240
            Value           =   1  'Checked
            Width           =   975
         End
      End
      Begin VB.Frame Frame20 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Informações do Material"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   120
         TabIndex        =   279
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   99
            Left            =   6240
            TabIndex        =   290
            Top             =   1200
            Width           =   495
         End
         Begin VB.CommandButton chameleonButton5 
            Caption         =   "..."
            Height          =   345
            Left            =   8520
            TabIndex        =   289
            Top             =   480
            Width           =   495
         End
         Begin VB.Frame Frame21 
            Height          =   975
            Left            =   120
            TabIndex        =   287
            Top             =   1560
            Width           =   8895
            Begin VB.TextBox Text1 
               BackColor       =   &H80000004&
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
               Height          =   615
               Left            =   105
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   288
               Top             =   165
               Width           =   8655
            End
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   98
            Left            =   8640
            TabIndex        =   286
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   97
            Left            =   120
            TabIndex        =   285
            Tag             =   "Código do Material"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   96
            Left            =   7920
            TabIndex        =   284
            Tag             =   "Quant. CJ"
            ToolTipText     =   "Quant. CJ"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   95
            Left            =   6855
            TabIndex        =   283
            Tag             =   "Quantidade Unitária"
            ToolTipText     =   "Quantidade Unitária"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   94
            Left            =   120
            TabIndex        =   282
            Top             =   1200
            Width           =   6015
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   345
            Index           =   93
            Left            =   2040
            TabIndex        =   280
            Top             =   480
            Width           =   6375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmFO.frx":5A82
            TabIndex        =   281
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
            Height          =   255
            Left            =   6240
            OleObjectBlob   =   "frmFO.frx":5AE4
            TabIndex        =   291
            Top             =   960
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
            Height          =   255
            Left            =   6840
            OleObjectBlob   =   "frmFO.frx":5B46
            TabIndex        =   292
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
            Height          =   255
            Left            =   7920
            OleObjectBlob   =   "frmFO.frx":5BB4
            TabIndex        =   293
            Top             =   960
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "frmFO.frx":5C26
            TabIndex        =   294
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel55 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5C92
            TabIndex        =   295
            Top             =   240
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5CF8
            TabIndex        =   296
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Total Individual "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   9360
         TabIndex        =   274
         Top             =   480
         Width           =   5895
         Begin ACTIVESKINLibCtl.SkinLabel Label39 
            Height          =   345
            Left            =   2280
            OleObjectBlob   =   "frmFO.frx":5D62
            TabIndex        =   275
            Top             =   480
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label38 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5DBC
            TabIndex        =   276
            Top             =   480
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
            Height          =   255
            Left            =   2280
            OleObjectBlob   =   "frmFO.frx":5E16
            TabIndex        =   277
            Top             =   240
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5E8A
            TabIndex        =   278
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Total geral"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   9360
         TabIndex        =   269
         Top             =   2280
         Width           =   5895
         Begin ACTIVESKINLibCtl.SkinLabel lblTotPint 
            Height          =   345
            Left            =   2280
            OleObjectBlob   =   "frmFO.frx":5EF8
            TabIndex        =   270
            Top             =   480
            Width           =   2490
         End
         Begin ACTIVESKINLibCtl.SkinLabel lblTotal 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5F52
            TabIndex        =   271
            Top             =   480
            Width           =   1935
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":5FAC
            TabIndex        =   272
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
            Height          =   255
            Left            =   2280
            OleObjectBlob   =   "frmFO.frx":601A
            TabIndex        =   273
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox txtLvw 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4560
         TabIndex        =   268
         Top             =   4800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
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
         ForeColor       =   &H80000008&
         Height          =   680
         Left            =   2040
         TabIndex        =   265
         Top             =   4380
         Width           =   1695
         Begin ACTIVESKINLibCtl.SkinLabel Label36 
            Height          =   345
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":6092
            TabIndex        =   266
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label37 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmFO.frx":60FA
            TabIndex        =   267
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Cálculo por "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   260
         Top             =   3360
         Width           =   5775
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   76
            Left            =   2970
            TabIndex        =   264
            Tag             =   "Peso"
            Top             =   495
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Height          =   345
            Index           =   74
            Left            =   120
            TabIndex        =   261
            Tag             =   "Dimensão"
            Top             =   480
            Width           =   2655
         End
         Begin VB.OptionButton optCadastro 
            Caption         =   "Peso:"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   263
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optCadastro 
            Caption         =   "Dimensão:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   262
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab SSTab6 
         Height          =   6975
         Left            =   -74760
         TabIndex        =   196
         Top             =   1440
         Width           =   14925
         _ExtentX        =   26326
         _ExtentY        =   12303
         _Version        =   393216
         TabOrientation  =   1
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
         TabCaption(0)   =   "Interna"
         TabPicture(0)   =   "frmFO.frx":6160
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "SSTab2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Externa"
         TabPicture(1)   =   "frmFO.frx":617C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab7"
         Tab(1).ControlCount=   1
         Begin TabDlg.SSTab SSTab2 
            Height          =   6255
            Left            =   120
            TabIndex        =   197
            Top             =   120
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   11033
            _Version        =   393216
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
            TabCaption(0)   =   "Levantamento de Tintas"
            TabPicture(0)   =   "frmFO.frx":6198
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Galão"
            TabPicture(1)   =   "frmFO.frx":61B4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "ListView13"
            Tab(1).Control(1)=   "Frame23"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Balde"
            TabPicture(2)   =   "frmFO.frx":61D0
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "ListView14"
            Tab(2).Control(1)=   "Frame24"
            Tab(2).ControlCount=   2
            Begin VB.Frame Frame24 
               Appearance      =   0  'Flat
               BackColor       =   &H00B7B7B7&
               Caption         =   "Balde (L)"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   -74880
               TabIndex        =   250
               Top             =   360
               Width           =   1095
               Begin VB.TextBox txtcadastro 
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
                  Index           =   109
                  Left            =   120
                  TabIndex        =   251
                  Tag             =   "Pintura"
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame23 
               Appearance      =   0  'Flat
               BackColor       =   &H00B7B7B7&
               Caption         =   "Galão (L)"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   735
               Left            =   -74880
               TabIndex        =   248
               Top             =   360
               Width           =   1095
               Begin VB.TextBox txtcadastro 
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
                  Index           =   108
                  Left            =   120
                  TabIndex        =   249
                  Tag             =   "Pintura"
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame11 
               Appearance      =   0  'Flat
               BackColor       =   &H00B7B7B7&
               Caption         =   "Levantamento de Tintas"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   5655
               Left            =   240
               TabIndex        =   198
               Top             =   360
               Width           =   14295
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   24
                  Left            =   2040
                  Picture         =   "frmFO.frx":61EC
                  Style           =   1  'Graphical
                  TabIndex        =   204
                  Tag             =   "Excluir Grupo"
                  ToolTipText     =   "Excluir Grupo"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   25
                  Left            =   1440
                  Picture         =   "frmFO.frx":6EB6
                  Style           =   1  'Graphical
                  TabIndex        =   205
                  Tag             =   "Editar Nome do Grupo"
                  ToolTipText     =   "Editar Nome do Grupo"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   26
                  Left            =   840
                  Picture         =   "frmFO.frx":7B80
                  Style           =   1  'Graphical
                  TabIndex        =   206
                  Tag             =   "Novo Grupo"
                  ToolTipText     =   "Novo Grupo"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtcadastro 
                  Height          =   345
                  Index           =   59
                  Left            =   960
                  TabIndex        =   209
                  Tag             =   "Nome do Grupo"
                  ToolTipText     =   "Nome do Grupo"
                  Top             =   480
                  Width           =   3375
               End
               Begin VB.TextBox txtcadastro 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   58
                  Left            =   120
                  TabIndex        =   208
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   735
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   27
                  Left            =   240
                  Picture         =   "frmFO.frx":884A
                  Style           =   1  'Graphical
                  TabIndex        =   207
                  Tag             =   "Incluir Grupo"
                  ToolTipText     =   "Incluir Grupo"
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtcadastro 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   57
                  Left            =   6600
                  TabIndex        =   203
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox txtcadastro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
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
                  Height          =   285
                  Index           =   56
                  Left            =   12360
                  TabIndex        =   202
                  Text            =   "x.xxx,xx"
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.ComboBox Combo3 
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
                  Left            =   4440
                  TabIndex        =   201
                  Text            =   "ACABAMENTO"
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.TextBox txtcadastro 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   55
                  Left            =   7920
                  TabIndex        =   200
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.TextBox txtcadastro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
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
                  Height          =   285
                  Index           =   54
                  Left            =   12360
                  TabIndex        =   199
                  Text            =   "x.xxx,xx"
                  Top             =   1320
                  Width           =   1695
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   30
                  Left            =   960
                  OleObjectBlob   =   "frmFO.frx":9514
                  TabIndex        =   210
                  Top             =   240
                  Width           =   2535
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmFO.frx":9596
                  TabIndex        =   211
                  Top             =   240
                  Width           =   855
               End
               Begin MSComctlLib.ListView ListView7 
                  Height          =   3735
                  Left            =   120
                  TabIndex        =   212
                  Top             =   1680
                  Width           =   14055
                  _ExtentX        =   24791
                  _ExtentY        =   6588
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
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   31
                  Left            =   4440
                  OleObjectBlob   =   "frmFO.frx":95F8
                  TabIndex        =   213
                  Top             =   240
                  Width           =   855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   32
                  Left            =   6600
                  OleObjectBlob   =   "frmFO.frx":965C
                  TabIndex        =   214
                  Top             =   240
                  Width           =   855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   33
                  Left            =   10200
                  OleObjectBlob   =   "frmFO.frx":96BC
                  TabIndex        =   215
                  Top             =   960
                  Width           =   2055
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   34
                  Left            =   7920
                  OleObjectBlob   =   "frmFO.frx":972E
                  TabIndex        =   216
                  Top             =   240
                  Width           =   1095
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   35
                  Left            =   9720
                  OleObjectBlob   =   "frmFO.frx":979C
                  TabIndex        =   217
                  Top             =   1320
                  Width           =   2535
               End
            End
            Begin MSComctlLib.ListView ListView13 
               Height          =   4815
               Left            =   -74880
               TabIndex        =   252
               Top             =   1200
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8493
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
            Begin MSComctlLib.ListView ListView14 
               Height          =   4815
               Left            =   -74880
               TabIndex        =   253
               Top             =   1200
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8493
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
         Begin TabDlg.SSTab SSTab7 
            Height          =   6255
            Left            =   -74880
            TabIndex        =   218
            Top             =   120
            Width           =   14655
            _ExtentX        =   25850
            _ExtentY        =   11033
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Lata"
            TabPicture(0)   =   "frmFO.frx":9828
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame12"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Galão"
            TabPicture(1)   =   "frmFO.frx":9844
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "ListView15"
            Tab(1).Control(1)=   "Frame25"
            Tab(1).ControlCount=   2
            TabCaption(2)   =   "Balde"
            TabPicture(2)   =   "frmFO.frx":9860
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "ListView16"
            Tab(2).Control(1)=   "Frame26"
            Tab(2).ControlCount=   2
            Begin VB.Frame Frame26 
               BackColor       =   &H00B7B7B7&
               Caption         =   "Balde (L)"
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
               Left            =   -74880
               TabIndex        =   257
               Top             =   360
               Width           =   1095
               Begin VB.TextBox txtcadastro 
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
                  Index           =   111
                  Left            =   120
                  TabIndex        =   258
                  Tag             =   "Pintura"
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame25 
               BackColor       =   &H00B7B7B7&
               Caption         =   "Galão (L)"
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
               Left            =   -74880
               TabIndex        =   254
               Top             =   360
               Width           =   1095
               Begin VB.TextBox txtcadastro 
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
                  Index           =   110
                  Left            =   120
                  TabIndex        =   255
                  Tag             =   "Pintura"
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame12 
               Appearance      =   0  'Flat
               BackColor       =   &H00B7B7B7&
               Caption         =   "Levantamento de Tintas"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   5655
               Left            =   120
               TabIndex        =   219
               Top             =   360
               Width           =   14415
               Begin VB.Frame Frame18 
                  BackColor       =   &H00B7B7B7&
                  Caption         =   "Lata"
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
                  Left            =   240
                  TabIndex        =   308
                  Top             =   240
                  Width           =   1095
                  Begin VB.TextBox txtcadastro 
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
                     Index           =   80
                     Left            =   120
                     TabIndex        =   309
                     Tag             =   "Pintura"
                     Top             =   240
                     Width           =   855
                  End
               End
               Begin VB.TextBox txtcadastro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
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
                  Height          =   285
                  Index           =   73
                  Left            =   12360
                  TabIndex        =   230
                  Text            =   "x.xxx,xx"
                  Top             =   1320
                  Width           =   1695
               End
               Begin VB.TextBox txtcadastro 
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
                  Height          =   345
                  Index           =   72
                  Left            =   9360
                  TabIndex        =   229
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.ComboBox Combo6 
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
                  Left            =   5880
                  TabIndex        =   228
                  Text            =   "ACABAMENTO"
                  Top             =   480
                  Width           =   2055
               End
               Begin VB.TextBox txtcadastro 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000F&
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
                  Height          =   285
                  Index           =   71
                  Left            =   12360
                  TabIndex        =   227
                  Text            =   "x.xxx,xx"
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.TextBox txtcadastro 
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
                  Height          =   345
                  Index           =   70
                  Left            =   8040
                  TabIndex        =   226
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   39
                  Left            =   2040
                  Picture         =   "frmFO.frx":987C
                  Style           =   1  'Graphical
                  TabIndex        =   225
                  Tag             =   "Excluir Grupo"
                  ToolTipText     =   "Excluir Grupo"
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   38
                  Left            =   1440
                  Picture         =   "frmFO.frx":A546
                  Style           =   1  'Graphical
                  TabIndex        =   224
                  Tag             =   "Editar Nome do Grupo"
                  ToolTipText     =   "Editar Nome do Grupo"
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   37
                  Left            =   840
                  Picture         =   "frmFO.frx":B210
                  Style           =   1  'Graphical
                  TabIndex        =   223
                  Tag             =   "Novo Grupo"
                  ToolTipText     =   "Novo Grupo"
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   36
                  Left            =   240
                  Picture         =   "frmFO.frx":BEDA
                  Style           =   1  'Graphical
                  TabIndex        =   222
                  Tag             =   "Incluir Grupo"
                  ToolTipText     =   "Incluir Grupo"
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.TextBox txtcadastro 
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
                  Height          =   345
                  Index           =   69
                  Left            =   1560
                  TabIndex        =   221
                  Tag             =   "Código do Grupo"
                  ToolTipText     =   "Código do Grupo"
                  Top             =   480
                  Width           =   735
               End
               Begin VB.TextBox txtcadastro 
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
                  Index           =   68
                  Left            =   2400
                  TabIndex        =   220
                  Tag             =   "Nome do Grupo"
                  ToolTipText     =   "Nome do Grupo"
                  Top             =   480
                  Width           =   3375
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   44
                  Left            =   2400
                  OleObjectBlob   =   "frmFO.frx":CBA4
                  TabIndex        =   231
                  Top             =   240
                  Width           =   2535
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
                  Height          =   255
                  Left            =   1560
                  OleObjectBlob   =   "frmFO.frx":CC26
                  TabIndex        =   232
                  Top             =   240
                  Width           =   855
               End
               Begin MSComctlLib.ListView ListView10 
                  Height          =   3615
                  Left            =   120
                  TabIndex        =   233
                  Top             =   1800
                  Width           =   14055
                  _ExtentX        =   24791
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
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   45
                  Left            =   5880
                  OleObjectBlob   =   "frmFO.frx":CC88
                  TabIndex        =   234
                  Top             =   240
                  Width           =   855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   46
                  Left            =   8040
                  OleObjectBlob   =   "frmFO.frx":CCEC
                  TabIndex        =   235
                  Top             =   240
                  Width           =   855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   47
                  Left            =   10200
                  OleObjectBlob   =   "frmFO.frx":CD4C
                  TabIndex        =   236
                  Top             =   960
                  Width           =   2055
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   48
                  Left            =   9360
                  OleObjectBlob   =   "frmFO.frx":CDBE
                  TabIndex        =   237
                  Top             =   240
                  Width           =   1095
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
                  Height          =   255
                  Index           =   49
                  Left            =   9720
                  OleObjectBlob   =   "frmFO.frx":CE2C
                  TabIndex        =   238
                  Top             =   1320
                  Width           =   2535
               End
            End
            Begin MSComctlLib.ListView ListView15 
               Height          =   4815
               Left            =   -74880
               TabIndex        =   256
               Top             =   1200
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8493
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
            Begin MSComctlLib.ListView ListView16 
               Height          =   4815
               Left            =   -74880
               TabIndex        =   259
               Top             =   1200
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8493
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
      End
      Begin TabDlg.SSTab SSTab5 
         Height          =   7935
         Left            =   -74760
         TabIndex        =   155
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "Matéria Prima"
         TabPicture(0)   =   "frmFO.frx":CEB8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame9(1)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Produto Industrializado"
         TabPicture(1)   =   "frmFO.frx":CED4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame10"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame10 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "Produto Industrializado "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7335
            Left            =   120
            TabIndex        =   176
            Top             =   360
            Width           =   14655
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   41
               Left            =   7920
               TabIndex        =   187
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1215
            End
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
               Left            =   5040
               TabIndex        =   186
               Text            =   "Normal"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   43
               Left            =   6600
               TabIndex        =   185
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   16
               Left            =   1920
               Picture         =   "frmFO.frx":CEF0
               Style           =   1  'Graphical
               TabIndex        =   184
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   17
               Left            =   1320
               Picture         =   "frmFO.frx":DBBA
               Style           =   1  'Graphical
               TabIndex        =   183
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   18
               Left            =   720
               Picture         =   "frmFO.frx":E884
               Style           =   1  'Graphical
               TabIndex        =   182
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   19
               Left            =   120
               Picture         =   "frmFO.frx":F54E
               Style           =   1  'Graphical
               TabIndex        =   181
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   44
               Left            =   120
               TabIndex        =   180
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtcadastro 
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
               Index           =   45
               Left            =   1560
               TabIndex        =   179
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   3375
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   42
               Left            =   12720
               TabIndex        =   178
               Text            =   "x.xxx,xx"
               Top             =   6480
               Width           =   1695
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   47
               Left            =   12720
               TabIndex        =   177
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   18
               Left            =   1560
               OleObjectBlob   =   "frmFO.frx":10218
               TabIndex        =   188
               Top             =   240
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":10286
               TabIndex        =   189
               Top             =   240
               Width           =   855
            End
            Begin MSComctlLib.ListView ListView5 
               Height          =   4695
               Left            =   120
               TabIndex        =   190
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8281
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   19
               Left            =   5040
               OleObjectBlob   =   "frmFO.frx":102F0
               TabIndex        =   191
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   20
               Left            =   6600
               OleObjectBlob   =   "frmFO.frx":10352
               TabIndex        =   192
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   22
               Left            =   7920
               OleObjectBlob   =   "frmFO.frx":103B6
               TabIndex        =   193
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   21
               Left            =   10560
               OleObjectBlob   =   "frmFO.frx":10422
               TabIndex        =   194
               Top             =   6600
               Width           =   2055
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   24
               Left            =   10080
               OleObjectBlob   =   "frmFO.frx":10494
               TabIndex        =   195
               Top             =   6960
               Width           =   2535
            End
         End
         Begin VB.Frame Frame9 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "Matéria Prima"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7335
            Index           =   1
            Left            =   -74880
            TabIndex        =   156
            Top             =   360
            Width           =   14655
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   15
               Left            =   1920
               Picture         =   "frmFO.frx":10520
               Style           =   1  'Graphical
               TabIndex        =   162
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   14
               Left            =   1320
               Picture         =   "frmFO.frx":111EA
               Style           =   1  'Graphical
               TabIndex        =   163
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   13
               Left            =   720
               Picture         =   "frmFO.frx":11EB4
               Style           =   1  'Graphical
               TabIndex        =   164
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
               Height          =   345
               Index           =   36
               Left            =   1560
               TabIndex        =   167
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   3375
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   37
               Left            =   120
               TabIndex        =   166
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1335
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   12
               Left            =   120
               Picture         =   "frmFO.frx":12B7E
               Style           =   1  'Graphical
               TabIndex        =   165
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   39
               Left            =   6600
               TabIndex        =   161
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   40
               Left            =   12720
               TabIndex        =   160
               Text            =   "x.xxx,xx"
               Top             =   6480
               Width           =   1695
            End
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
               Left            =   5040
               TabIndex        =   159
               Text            =   "Normal"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   38
               Left            =   7920
               TabIndex        =   158
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   46
               Left            =   12720
               TabIndex        =   157
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   13
               Left            =   1560
               OleObjectBlob   =   "frmFO.frx":13848
               TabIndex        =   168
               Top             =   240
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":138B6
               TabIndex        =   169
               Top             =   240
               Width           =   855
            End
            Begin MSComctlLib.ListView ListView4 
               Height          =   4695
               Left            =   120
               TabIndex        =   170
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
               _ExtentY        =   8281
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   14
               Left            =   5040
               OleObjectBlob   =   "frmFO.frx":13920
               TabIndex        =   171
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   15
               Left            =   6600
               OleObjectBlob   =   "frmFO.frx":13982
               TabIndex        =   172
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   16
               Left            =   10560
               OleObjectBlob   =   "frmFO.frx":139E6
               TabIndex        =   173
               Top             =   6600
               Width           =   2055
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   17
               Left            =   7920
               OleObjectBlob   =   "frmFO.frx":13A58
               TabIndex        =   174
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   23
               Left            =   10080
               OleObjectBlob   =   "frmFO.frx":13AC4
               TabIndex        =   175
               Top             =   6960
               Width           =   2535
            End
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   7935
         Left            =   -74760
         TabIndex        =   122
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
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
         TabCaption(0)   =   "Interna"
         TabPicture(0)   =   "frmFO.frx":13B50
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame7"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Externa"
         TabPicture(1)   =   "frmFO.frx":13B6C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame4"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "Esquema "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7335
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   14655
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   32
               Left            =   1920
               Picture         =   "frmFO.frx":13B88
               Style           =   1  'Graphical
               TabIndex        =   143
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   33
               Left            =   1320
               Picture         =   "frmFO.frx":14852
               Style           =   1  'Graphical
               TabIndex        =   144
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   34
               Left            =   720
               Picture         =   "frmFO.frx":1551C
               Style           =   1  'Graphical
               TabIndex        =   145
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Index           =   67
               Left            =   2760
               TabIndex        =   148
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   5535
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   35
               Left            =   120
               Picture         =   "frmFO.frx":161E6
               Style           =   1  'Graphical
               TabIndex        =   147
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   66
               Left            =   8400
               TabIndex        =   146
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   65
               Left            =   9960
               TabIndex        =   142
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   12
               Left            =   12720
               TabIndex        =   141
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin VB.ComboBox Combo5 
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
               TabIndex        =   140
               Text            =   "Intermediária"
               Top             =   480
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   40
               Left            =   2760
               OleObjectBlob   =   "frmFO.frx":16EB0
               TabIndex        =   149
               Top             =   240
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":16F1E
               TabIndex        =   150
               Top             =   240
               Width           =   855
            End
            Begin MSComctlLib.ListView ListView9 
               Height          =   5055
               Left            =   120
               TabIndex        =   151
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   41
               Left            =   8400
               OleObjectBlob   =   "frmFO.frx":16F88
               TabIndex        =   152
               Top             =   240
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   42
               Left            =   9960
               OleObjectBlob   =   "frmFO.frx":16FF6
               TabIndex        =   153
               Top             =   240
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   43
               Left            =   12000
               OleObjectBlob   =   "frmFO.frx":17060
               TabIndex        =   154
               Top             =   6960
               Width           =   615
            End
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "Esquema "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7335
            Left            =   -74880
            TabIndex        =   123
            Top             =   360
            Width           =   14655
            Begin VB.ComboBox Combo4 
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
               TabIndex        =   138
               Text            =   "Intermediária"
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   8
               Left            =   12720
               TabIndex        =   131
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   9
               Left            =   9960
               TabIndex        =   130
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1695
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   4
               Left            =   1920
               Picture         =   "frmFO.frx":170C4
               Style           =   1  'Graphical
               TabIndex        =   129
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   5
               Left            =   1320
               Picture         =   "frmFO.frx":17D8E
               Style           =   1  'Graphical
               TabIndex        =   128
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   6
               Left            =   720
               Picture         =   "frmFO.frx":18A58
               Style           =   1  'Graphical
               TabIndex        =   127
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   10
               Left            =   8400
               TabIndex        =   126
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   7
               Left            =   120
               Picture         =   "frmFO.frx":19722
               Style           =   1  'Graphical
               TabIndex        =   125
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   26
               Left            =   2760
               TabIndex        =   124
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   5535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   4
               Left            =   2760
               OleObjectBlob   =   "frmFO.frx":1A3EC
               TabIndex        =   132
               Top             =   240
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":1A45A
               TabIndex        =   133
               Top             =   240
               Width           =   855
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   5055
               Left            =   120
               TabIndex        =   134
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   5
               Left            =   8400
               OleObjectBlob   =   "frmFO.frx":1A4C4
               TabIndex        =   135
               Top             =   240
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   6
               Left            =   9960
               OleObjectBlob   =   "frmFO.frx":1A532
               TabIndex        =   136
               Top             =   240
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   7
               Left            =   12000
               OleObjectBlob   =   "frmFO.frx":1A59C
               TabIndex        =   137
               Top             =   6960
               Width           =   615
            End
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7935
         Left            =   -74760
         TabIndex        =   89
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13996
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   12040119
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Impostos"
         TabPicture(0)   =   "frmFO.frx":1A600
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Serviços"
         TabPicture(1)   =   "frmFO.frx":1A61C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3(2)"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "Serviços "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   7335
            Index           =   2
            Left            =   -74880
            TabIndex        =   106
            Top             =   360
            Width           =   14655
            Begin VB.CommandButton Command7 
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
               Height          =   345
               Left            =   11640
               TabIndex        =   245
               Top             =   480
               Width           =   495
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   64
               Left            =   12600
               TabIndex        =   115
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   63
               Left            =   9840
               TabIndex        =   114
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1695
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   31
               Left            =   1920
               Picture         =   "frmFO.frx":1A638
               Style           =   1  'Graphical
               TabIndex        =   113
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   30
               Left            =   1320
               Picture         =   "frmFO.frx":1B302
               Style           =   1  'Graphical
               TabIndex        =   112
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   29
               Left            =   720
               Picture         =   "frmFO.frx":1BFCC
               Style           =   1  'Graphical
               TabIndex        =   111
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   62
               Left            =   8640
               TabIndex        =   110
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   28
               Left            =   120
               Picture         =   "frmFO.frx":1CC96
               Style           =   1  'Graphical
               TabIndex        =   109
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
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
               Height          =   345
               Index           =   61
               Left            =   120
               TabIndex        =   108
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtcadastro 
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
               Index           =   60
               Left            =   1320
               TabIndex        =   107
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   7215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   36
               Left            =   1320
               OleObjectBlob   =   "frmFO.frx":1D960
               TabIndex        =   116
               Top             =   240
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":1D9C2
               TabIndex        =   117
               Top             =   240
               Width           =   615
            End
            Begin MSComctlLib.ListView ListView8 
               Height          =   5055
               Left            =   120
               TabIndex        =   118
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   37
               Left            =   8640
               OleObjectBlob   =   "frmFO.frx":1DA20
               TabIndex        =   119
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   38
               Left            =   9840
               OleObjectBlob   =   "frmFO.frx":1DA8A
               TabIndex        =   120
               Top             =   240
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   39
               Left            =   11880
               OleObjectBlob   =   "frmFO.frx":1DB02
               TabIndex        =   121
               Top             =   6960
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
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
            ForeColor       =   &H80000008&
            Height          =   7335
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   14655
            Begin VB.CommandButton Command6 
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
               Height          =   345
               Left            =   11640
               TabIndex        =   244
               Top             =   480
               Width           =   495
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   3
               Left            =   1920
               Picture         =   "frmFO.frx":1DB66
               Style           =   1  'Graphical
               TabIndex        =   93
               Tag             =   "Excluir Grupo"
               ToolTipText     =   "Excluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   2
               Left            =   1320
               Picture         =   "frmFO.frx":1E830
               Style           =   1  'Graphical
               TabIndex        =   94
               Tag             =   "Editar Nome do Grupo"
               ToolTipText     =   "Editar Nome do Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   1
               Left            =   720
               Picture         =   "frmFO.frx":1F4FA
               Style           =   1  'Graphical
               TabIndex        =   95
               Tag             =   "Novo Grupo"
               ToolTipText     =   "Novo Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   0
               Left            =   120
               Picture         =   "frmFO.frx":201C4
               Style           =   1  'Graphical
               TabIndex        =   97
               Tag             =   "Incluir Grupo"
               ToolTipText     =   "Incluir Grupo"
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
               Height          =   345
               Index           =   2
               Left            =   1320
               TabIndex        =   99
               Tag             =   "Nome do Grupo"
               ToolTipText     =   "Nome do Grupo"
               Top             =   480
               Width           =   7215
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   3
               Left            =   120
               TabIndex        =   98
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   5
               Left            =   8640
               TabIndex        =   96
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   4
               Left            =   9840
               TabIndex        =   92
               Tag             =   "Código do Grupo"
               ToolTipText     =   "Código do Grupo"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtcadastro 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
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
               Height          =   345
               Index           =   7
               Left            =   12600
               TabIndex        =   91
               Text            =   "x.xxx,xx"
               Top             =   6840
               Width           =   1695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   0
               Left            =   1320
               OleObjectBlob   =   "frmFO.frx":20E8E
               TabIndex        =   100
               Top             =   240
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":20EF0
               TabIndex        =   101
               Top             =   240
               Width           =   615
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   5055
               Left            =   120
               TabIndex        =   102
               Top             =   1680
               Width           =   14415
               _ExtentX        =   25426
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
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   1
               Left            =   8640
               OleObjectBlob   =   "frmFO.frx":20F4E
               TabIndex        =   103
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   2
               Left            =   9840
               OleObjectBlob   =   "frmFO.frx":20FB8
               TabIndex        =   104
               Top             =   240
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   345
               Index           =   3
               Left            =   11880
               OleObjectBlob   =   "frmFO.frx":21030
               TabIndex        =   105
               Top             =   6960
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Despesas / Créditos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7935
         Index           =   1
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   15015
         Begin VB.CommandButton Command8 
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
            Height          =   345
            Left            =   9960
            TabIndex        =   246
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   20
            Left            =   1920
            Picture         =   "frmFO.frx":21094
            Style           =   1  'Graphical
            TabIndex        =   74
            Tag             =   "Excluir Grupo"
            ToolTipText     =   "Excluir Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   21
            Left            =   1320
            Picture         =   "frmFO.frx":21D5E
            Style           =   1  'Graphical
            TabIndex        =   75
            Tag             =   "Editar Nome do Grupo"
            ToolTipText     =   "Editar Nome do Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   22
            Left            =   720
            Picture         =   "frmFO.frx":22A28
            Style           =   1  'Graphical
            TabIndex        =   76
            Tag             =   "Novo Grupo"
            ToolTipText     =   "Novo Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   53
            Left            =   6840
            TabIndex        =   87
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   52
            Left            =   1320
            TabIndex        =   80
            Tag             =   "Nome do Grupo"
            ToolTipText     =   "Nome do Grupo"
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   51
            Left            =   120
            TabIndex        =   79
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   23
            Left            =   120
            Picture         =   "frmFO.frx":236F2
            Style           =   1  'Graphical
            TabIndex        =   78
            Tag             =   "Incluir Grupo"
            ToolTipText     =   "Incluir Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   50
            Left            =   5160
            TabIndex        =   77
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   49
            Left            =   8160
            TabIndex        =   73
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   345
            Index           =   48
            Left            =   13080
            TabIndex        =   72
            Text            =   "x.xxx,xx"
            Top             =   7440
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   25
            Left            =   1320
            OleObjectBlob   =   "frmFO.frx":243BC
            TabIndex        =   81
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":2441E
            TabIndex        =   82
            Top             =   240
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView6 
            Height          =   5655
            Left            =   120
            TabIndex        =   83
            Top             =   1680
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   9975
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   26
            Left            =   5160
            OleObjectBlob   =   "frmFO.frx":2447C
            TabIndex        =   84
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   27
            Left            =   8160
            OleObjectBlob   =   "frmFO.frx":244E0
            TabIndex        =   85
            Top             =   240
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   345
            Index           =   28
            Left            =   12360
            OleObjectBlob   =   "frmFO.frx":24558
            TabIndex        =   86
            Top             =   7560
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   29
            Left            =   6840
            OleObjectBlob   =   "frmFO.frx":245BC
            TabIndex        =   88
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Testes e Ensaios:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7935
         Index           =   1
         Left            =   -74880
         TabIndex        =   53
         Top             =   480
         Width           =   15015
         Begin VB.CommandButton Command9 
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
            Height          =   345
            Left            =   1920
            TabIndex        =   247
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   35
            Left            =   11760
            TabIndex        =   69
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Height          =   345
            Index           =   34
            Left            =   13080
            TabIndex        =   62
            Text            =   "x.xxx,xx"
            Top             =   7440
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   33
            Left            =   9960
            TabIndex        =   61
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   11
            Left            =   1920
            Picture         =   "frmFO.frx":2462E
            Style           =   1  'Graphical
            TabIndex        =   60
            Tag             =   "Excluir Grupo"
            ToolTipText     =   "Excluir Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   10
            Left            =   1320
            Picture         =   "frmFO.frx":252F8
            Style           =   1  'Graphical
            TabIndex        =   59
            Tag             =   "Editar Nome do Grupo"
            ToolTipText     =   "Editar Nome do Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   9
            Left            =   720
            Picture         =   "frmFO.frx":25FC2
            Style           =   1  'Graphical
            TabIndex        =   58
            Tag             =   "Novo Grupo"
            ToolTipText     =   "Novo Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   32
            Left            =   8280
            TabIndex        =   57
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   8
            Left            =   120
            Picture         =   "frmFO.frx":26C8C
            Style           =   1  'Graphical
            TabIndex        =   56
            Tag             =   "Incluir Grupo"
            ToolTipText     =   "Incluir Grupo"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   31
            Left            =   120
            TabIndex        =   55
            Tag             =   "Código do Grupo"
            ToolTipText     =   "Código do Grupo"
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   29
            Left            =   2640
            TabIndex        =   54
            Tag             =   "Nome do Grupo"
            ToolTipText     =   "Nome do Grupo"
            Top             =   480
            Width           =   5535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   8
            Left            =   2640
            OleObjectBlob   =   "frmFO.frx":27956
            TabIndex        =   63
            Top             =   240
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":279DA
            TabIndex        =   64
            Top             =   240
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   5655
            Left            =   120
            TabIndex        =   65
            Top             =   1680
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   9975
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   9
            Left            =   8280
            OleObjectBlob   =   "frmFO.frx":27A44
            TabIndex        =   66
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   10
            Left            =   9960
            OleObjectBlob   =   "frmFO.frx":27ABC
            TabIndex        =   67
            Top             =   240
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   345
            Index           =   11
            Left            =   12360
            OleObjectBlob   =   "frmFO.frx":27B36
            TabIndex        =   68
            Top             =   7440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   12
            Left            =   11760
            OleObjectBlob   =   "frmFO.frx":27B9A
            TabIndex        =   70
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
         Caption         =   "Dados da FO "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   7935
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   6015
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00B7B7B7&
            Caption         =   "FCE - Ficha de Controle de Encomenda "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   975
            Left            =   120
            TabIndex        =   44
            Top             =   2640
            Width           =   5775
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":27C04
               TabIndex        =   45
               Top             =   480
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel Label32 
               Height          =   255
               Left            =   1920
               OleObjectBlob   =   "frmFO.frx":27C64
               TabIndex        =   46
               Top             =   480
               Width           =   2415
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFO.frx":27CBC
               TabIndex        =   47
               Top             =   240
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   1920
               OleObjectBlob   =   "frmFO.frx":27D2C
               TabIndex        =   48
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3615
            Index           =   28
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   37
            Tag             =   "Observação"
            Top             =   4080
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
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
            Index           =   11
            Left            =   120
            TabIndex        =   36
            Tag             =   "Nº da SDC"
            ToolTipText     =   "Nº da Solicitação de Cotação"
            Top             =   1920
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
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
            Index           =   30
            Left            =   120
            TabIndex        =   35
            Tag             =   "Descrição do serviço"
            Top             =   1200
            Width           =   5775
         End
         Begin VB.TextBox txtcadastro 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   390
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Tag             =   "Nº Ficha de Orçamento"
            Top             =   480
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   345
            Left            =   3240
            TabIndex        =   32
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
            Format          =   161480705
            CurrentDate     =   40449
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   345
            Left            =   1680
            TabIndex        =   33
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
            Format          =   161480705
            CurrentDate     =   40449
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":27D9C
            TabIndex        =   38
            Top             =   3840
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":27E0A
            TabIndex        =   39
            Top             =   1680
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":27E70
            TabIndex        =   40
            Top             =   960
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "frmFO.frx":27EDC
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "frmFO.frx":27F44
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":27FAE
            TabIndex        =   43
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
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
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   -68640
         TabIndex        =   24
         Top             =   4560
         Width           =   8775
         Begin VB.CommandButton Command5 
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
            Height          =   345
            Left            =   8160
            TabIndex        =   243
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   1
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   0
            Left            =   3960
            TabIndex        =   49
            Top             =   1200
            Width           =   4695
         End
         Begin VB.TextBox txtcadastro 
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
            Index           =   24
            Left            =   120
            TabIndex        =   27
            Tag             =   "Código do Contato"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   25
            Left            =   1200
            TabIndex        =   26
            Top             =   480
            Width           =   6855
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   27
            Left            =   120
            TabIndex        =   25
            Top             =   1920
            Width           =   8535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":28012
            TabIndex        =   28
            Top             =   1680
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmFO.frx":28076
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":280D8
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3960
            OleObjectBlob   =   "frmFO.frx":2813E
            TabIndex        =   51
            Top             =   960
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":281A6
            TabIndex        =   52
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B7B7B7&
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
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   -68640
         TabIndex        =   1
         Top             =   480
         Width           =   8775
         Begin VB.CommandButton Command4 
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
            Height          =   345
            Left            =   8160
            TabIndex        =   242
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtcadastro 
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
            Index           =   13
            Left            =   120
            TabIndex        =   12
            Tag             =   "Código do Cliente"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   14
            Left            =   1080
            TabIndex        =   11
            Top             =   480
            Width           =   6975
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   15
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   7335
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   16
            Left            =   7560
            TabIndex        =   9
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   17
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   3735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   18
            Left            =   3960
            TabIndex        =   7
            Top             =   1920
            Width           =   3855
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   19
            Left            =   7920
            TabIndex        =   6
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   20
            Left            =   120
            TabIndex        =   5
            Top             =   2640
            Width           =   3735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   21
            Left            =   3960
            TabIndex        =   4
            Top             =   2640
            Width           =   4695
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   22
            Left            =   120
            TabIndex        =   3
            Top             =   3360
            Width           =   3735
         End
         Begin VB.TextBox txtcadastro 
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
            Height          =   345
            Index           =   23
            Left            =   3960
            TabIndex        =   2
            Top             =   3360
            Width           =   4695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   3960
            OleObjectBlob   =   "frmFO.frx":28210
            TabIndex        =   13
            Top             =   3120
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":28272
            TabIndex        =   14
            Top             =   3120
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   3960
            OleObjectBlob   =   "frmFO.frx":282D6
            TabIndex        =   15
            Top             =   2400
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":2833E
            TabIndex        =   16
            Top             =   2400
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Index           =   1
            Left            =   7920
            OleObjectBlob   =   "frmFO.frx":283A8
            TabIndex        =   17
            Top             =   1680
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Index           =   1
            Left            =   3960
            OleObjectBlob   =   "frmFO.frx":2840E
            TabIndex        =   18
            Top             =   1680
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":28474
            TabIndex        =   19
            Top             =   1680
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   7560
            OleObjectBlob   =   "frmFO.frx":284DA
            TabIndex        =   20
            Top             =   960
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":2853A
            TabIndex        =   21
            Top             =   960
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   1080
            OleObjectBlob   =   "frmFO.frx":285A4
            TabIndex        =   22
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFO.frx":28606
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   -63960
         OleObjectBlob   =   "frmFO.frx":2866C
         TabIndex        =   239
         Top             =   8280
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   -63960
         OleObjectBlob   =   "frmFO.frx":286F6
         TabIndex        =   240
         Top             =   7920
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView12 
         Height          =   7215
         Left            =   -74760
         TabIndex        =   241
         Top             =   480
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   12726
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
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   3840
         Top             =   4440
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
         Height          =   735
         Left            =   9480
         OleObjectBlob   =   "frmFO.frx":28774
         TabIndex        =   300
         Top             =   3480
         Width           =   5655
      End
      Begin MSComctlLib.ListView ListView11 
         Height          =   3375
         Left            =   120
         TabIndex        =   301
         Tag             =   "Duplo clique para editar"
         ToolTipText     =   "Duplo clique para editar"
         Top             =   5160
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   5953
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
      Begin VB.Shape Shape1 
         BackColor       =   &H00B7B7B7&
         BorderColor     =   &H000000C0&
         Height          =   945
         Left            =   9360
         Top             =   3360
         Width           =   5880
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      FillColor       =   &H80000006&
      Height          =   1455
      Left            =   10680
      Top             =   1200
      Width           =   2160
   End
End
Attribute VB_Name = "frmFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLocal As New ADODB.Recordset
Private rsMaterial As New ADODB.Recordset
Private const0(9) As Double
Private vAr0(9) As Double
Private X As Integer, Y As Integer
Private Conta As Integer
Private Formula As String
Private ForPint As String
Private SqlM As String
Private SomaTotal As Double
Private SomaPint As Double
Private QuantCJ As Double
Private PesoTotal As Double
Private TipoCad As String

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 0
        txtCadastro_KeyDown 2, 13, 1
        LimpaControles
        SomaListview
        txtcadastro(0).SetFocus
    Case 1
        ExcluirItem
        SomaListview
        txtcadastro(0).SetFocus
    Case 2
        AlterarItem1
    Case 3
        'GerarResumo
        SSTab1.Tab = 2
        optCadastro(3).Value = True
        Check3.Value = 1
        Msgbox "Resumo gerado com sucesso"
    Case 4
        GravarDados
        txtcadastro(0).SetFocus
    Case 5
        'ExportaExcel
        Msgbox "Dados exportados com sucesso", vbInformation, "Zeus"
    Case 6
        If Msgbox("Deseja sair da tela de cadastro?", vbQuestion + vbYesNo, "Zeus") = vbYes Then
            'CancelaSN = 1
            Unload Me
        End If
    End Select
End Sub

Private Sub Command4_Click()
    txtcadastro(13) = ""
    ChamaGridCli
End Sub

Private Sub Command5_Click()
    txtcadastro(24) = ""
    ChamaGridCont
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub optCadastro_Click(Index As Integer)
    If optCadastro(0).Value = True Then
        txtcadastro(2).Enabled = True
        txtcadastro(8).Enabled = False
        txtcadastro(8).BackColor = &H80000004
        txtcadastro(2).BackColor = &H80000005
    End If
    If optCadastro(1).Value = True Then
        txtcadastro(8).Enabled = True
        txtcadastro(2).Enabled = False
        txtcadastro(2).BackColor = &H80000004
        txtcadastro(8).BackColor = &H80000005
    End If

'    If optCadastro(2).Value = True Then
'        txtcadastro(34).Enabled = True
'        txtcadastro(35).Enabled = False
'        txtcadastro(36).Enabled = False
'        txtcadastro(37).Enabled = True
'        txtcadastro(36).BackColor = &H80000004
'        txtcadastro(35).BackColor = &H80000004
'        txtcadastro(34).BackColor = &H80000005
'        txtcadastro(37).BackColor = &H80000005
'        Check3.Value = 0
'        Check3.Enabled = False
'    End If
'    If optCadastro(3).Value = True Then
'        txtcadastro(34).Enabled = False
'        txtcadastro(35).Enabled = True
'        txtcadastro(37).Enabled = False
'        txtcadastro(35).BackColor = &H80000005
'        txtcadastro(34).BackColor = &H80000004
'        txtcadastro(37).BackColor = &H80000004
'        Check3.Enabled = True
'        Check3.Value = 0
'    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value = False Then
        Frame14.Enabled = False
        Label32.Enabled = False
        Label33.Enabled = False
        SkinLabel45.Enabled = False
        txtcadastro(41).Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option3.Value = False Then
        Frame14.Enabled = False
        Label32.Enabled = False
        Label33.Enabled = False
        SkinLabel45.Enabled = False
        txtcadastro(41).Enabled = False
    End If
End Sub

Private Sub Option3_Click()
    If Option3.Value = True Then
        Frame14.Enabled = True
        Label32.Enabled = True
        Label33.Enabled = True
        SkinLabel45.Enabled = True
        txtcadastro(41).Enabled = True
    End If
'    SkinLabel45.SetFocus
End Sub

Private Sub Form_Load()
    inicializa_tabs
    lv_cab
    
    SomaTotal = 0
    SomaPint = 0
    TipoCad = Pesquisa
    If TipoCad = "novo" Then
        LimpaControles
        txtcadastro(6) = Format(GeraCodigo, "000000") & ""
        txtcadastro(6).Enabled = False
        optCadastro_Click (0)
    ElseIf TipoCad = "editar" Then
        ResultPesq
        DesbloqueiaControles
        txtCadastro_KeyDown 6, 13, 6
    End If
    optCadastro_Click 1
    carregarIconBotao
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
    
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdcadastro(0), 0, 46 'Inserir
    carregaImagemBotao cmdcadastro(2), 2, 32 'Editar
    carregaImagemBotao cmdcadastro(1), 1, 31 'Novo
    carregaImagemBotao cmdcadastro(3), 3, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(23), 23, 46 'Inserir
    carregaImagemBotao cmdcadastro(21), 21, 32 'Editar
    carregaImagemBotao cmdcadastro(22), 22, 31 'Novo
    carregaImagemBotao cmdcadastro(20), 20, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(7), 7, 46 'Inserir
    carregaImagemBotao cmdcadastro(5), 5, 32 'Editar
    carregaImagemBotao cmdcadastro(6), 6, 31 'Novo
    carregaImagemBotao cmdcadastro(4), 4, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(35), 35, 46 'Inserir
    carregaImagemBotao cmdcadastro(33), 33, 32 'Editar
    carregaImagemBotao cmdcadastro(34), 34, 31 'Novo
    carregaImagemBotao cmdcadastro(32), 32, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(8), 8, 46 'Inserir
    carregaImagemBotao cmdcadastro(10), 10, 32 'Editar
    carregaImagemBotao cmdcadastro(9), 9, 31 'Novo
    carregaImagemBotao cmdcadastro(11), 11, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(12), 12, 46 'Inserir
    carregaImagemBotao cmdcadastro(14), 14, 32 'Editar
    carregaImagemBotao cmdcadastro(13), 13, 31 'Novo
    carregaImagemBotao cmdcadastro(15), 15, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(19), 19, 46 'Inserir
    carregaImagemBotao cmdcadastro(17), 17, 32 'Editar
    carregaImagemBotao cmdcadastro(18), 18, 31 'Novo
    carregaImagemBotao cmdcadastro(16), 16, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(36), 36, 46 'Inserir
    carregaImagemBotao cmdcadastro(38), 38, 32 'Editar
    carregaImagemBotao cmdcadastro(37), 37, 31 'Novo
    carregaImagemBotao cmdcadastro(39), 39, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(27), 27, 46 'Inserir
    carregaImagemBotao cmdcadastro(25), 25, 32 'Editar
    carregaImagemBotao cmdcadastro(26), 26, 31 'Novo
    carregaImagemBotao cmdcadastro(24), 24, 33 'Excluir
    
    carregaImagemBotao cmdcadastro(28), 28, 46 'Inserir
    carregaImagemBotao cmdcadastro(30), 30, 32 'Editar
    carregaImagemBotao cmdcadastro(29), 29, 31 'Novo
    carregaImagemBotao cmdcadastro(31), 31, 33 'Excluir
    
    carregaImagemBotao chamCad(0), 0, 46 'Inserir
    carregaImagemBotao chamCad(1), 1, 33 'Excluir
    carregaImagemBotao chamCad(7), 7, 50 'Engrenagem
    carregaImagemBotao chamCad(5), 5, 51 'Excel
    carregaImagemBotao chamCad(4), 4, 45 'Salvar
    carregaImagemBotao chamCad(6), 6, 34 'Sair
    
End Sub

Private Sub LimpaControles()
    Formula = ""
    ForPint = ""
    Conta = 0
End Sub

Private Sub LimpaControles1()
    Dim X As Integer
    For X = 0 To 5
        txtcadastro(X) = ""
    Next
    For X = 7 To 38
        txtcadastro(X) = ""
    Next
    txtcadastro(8) = ""
    ListView1.ListItems.Clear
    Formula = ""
    ForPint = ""
    Conta = 0
End Sub

Private Sub ResultPesq()
    txtcadastro(6) = varGlobal
End Sub

Sub lv_cab()
    
'-- IMPOSTOS E SERVIÇOS
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Alíquota", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Total", ListView1.Width / 8
    
    ListView8.ColumnHeaders.Add , , "ID", ListView8.Width / 8
    ListView8.ColumnHeaders.Add , , "Nome", ListView8.Width / 4
    ListView8.ColumnHeaders.Add , , "Alíquota", ListView8.Width / 8
    ListView8.ColumnHeaders.Add , , "Total", ListView8.Width / 8
    
'-- TESTES E ENSAIOS
    ListView3.ColumnHeaders.Add , , "Sistema", ListView3.Width / 8
    ListView3.ColumnHeaders.Add , , "Qtd. Pessoas", ListView3.Width / 4
    ListView3.ColumnHeaders.Add , , "Nº Dias por Mês", ListView3.Width / 8
    ListView3.ColumnHeaders.Add , , "Qtd. Meses", ListView3.Width / 8
    ListView3.ColumnHeaders.Add , , "Valor MO", ListView3.Width / 8
    
'-- DESPESAS/CREDITOS
    ListView6.ColumnHeaders.Add , , "ID", ListView6.Width / 8
    ListView6.ColumnHeaders.Add , , "Nome", ListView6.Width / 4
    ListView6.ColumnHeaders.Add , , "Tipo", ListView6.Width / 8
    ListView6.ColumnHeaders.Add , , "Percentual (%)", ListView6.Width / 8
    ListView6.ColumnHeaders.Add , , "Valor Calculado", ListView6.Width / 8

'-- PINTURA (INTERNA/EXTERNA)
    ListView2.ColumnHeaders.Add , , "Sistema", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Referência", ListView2.Width / 4
    ListView2.ColumnHeaders.Add , , "Quantidade", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Valor MO", ListView2.Width / 8

    ListView9.ColumnHeaders.Add , , "Sistema", ListView9.Width / 8
    ListView9.ColumnHeaders.Add , , "Referência", ListView9.Width / 4
    ListView9.ColumnHeaders.Add , , "Quantidade", ListView9.Width / 8
    ListView9.ColumnHeaders.Add , , "Valor MO", ListView9.Width / 8

'-- TRANSPORTE (MATÉRIA PRIMA/INDUSTRIALIZADO)
    ListView4.ColumnHeaders.Add , , "Sistema", ListView4.Width / 8
    ListView4.ColumnHeaders.Add , , "Fornecedor", ListView4.Width / 4
    ListView4.ColumnHeaders.Add , , "Tipo", ListView4.Width / 8
    ListView4.ColumnHeaders.Add , , "Valor", ListView4.Width / 8
    ListView4.ColumnHeaders.Add , , "Subtotal", ListView4.Width / 8

    ListView5.ColumnHeaders.Add , , "Sistema", ListView5.Width / 8
    ListView5.ColumnHeaders.Add , , "Fornecedor", ListView5.Width / 4
    ListView5.ColumnHeaders.Add , , "Tipo", ListView5.Width / 8
    ListView5.ColumnHeaders.Add , , "Valor", ListView5.Width / 8
    ListView5.ColumnHeaders.Add , , "Subtotal", ListView5.Width / 8

'--TINTAS
    ListView7.ColumnHeaders.Add , , "Item", ListView7.Width / 8
    ListView7.ColumnHeaders.Add , , "Des. Produto", ListView7.Width / 4
    ListView7.ColumnHeaders.Add , , "Tinta", ListView7.Width / 8
    ListView7.ColumnHeaders.Add , , "Cor", ListView7.Width / 8
    ListView7.ColumnHeaders.Add , , "Diluição K", ListView7.Width / 8
    
    ListView10.ColumnHeaders.Add , , "Item", ListView10.Width / 8
    ListView10.ColumnHeaders.Add , , "Des. Produto", ListView10.Width / 4
    ListView10.ColumnHeaders.Add , , "Tinta", ListView10.Width / 8
    ListView10.ColumnHeaders.Add , , "Cor", ListView10.Width / 8
    ListView10.ColumnHeaders.Add , , "Diluição K", ListView10.Width / 8
    
'-- TINTAS | GALÃO
    ListView13.ColumnHeaders.Add , , "Arredondamento", ListView13.Width / 8
    ListView13.ColumnHeaders.Add , , "Valor (Unit)", ListView13.Width / 4
    ListView13.ColumnHeaders.Add , , "Qtd. Final", ListView13.Width / 8
    ListView13.ColumnHeaders.Add , , "Custo Total", ListView13.Width / 8
    ListView13.ColumnHeaders.Add , , "Custo m² com Solvente", ListView13.Width / 8
    
    ListView15.ColumnHeaders.Add , , "Arredondamento", ListView15.Width / 8
    ListView15.ColumnHeaders.Add , , "Valor (Unit)", ListView15.Width / 4
    ListView15.ColumnHeaders.Add , , "Qtd. Final", ListView15.Width / 8
    ListView15.ColumnHeaders.Add , , "Custo Total", ListView15.Width / 8
    ListView15.ColumnHeaders.Add , , "Custo m² com Solvente", ListView15.Width / 8
    
'-- TINTAS | BALDE
    ListView14.ColumnHeaders.Add , , "Arredondamento", ListView14.Width / 8
    ListView14.ColumnHeaders.Add , , "Valor (Unit)", ListView14.Width / 4
    ListView14.ColumnHeaders.Add , , "Qtd. Final", ListView14.Width / 8
    ListView14.ColumnHeaders.Add , , "Custo Total", ListView14.Width / 8
    ListView14.ColumnHeaders.Add , , "Custo m² com Solvente", ListView14.Width / 8
    
    ListView16.ColumnHeaders.Add , , "Arredondamento", ListView16.Width / 8
    ListView16.ColumnHeaders.Add , , "Valor (Unit)", ListView16.Width / 4
    ListView16.ColumnHeaders.Add , , "Qtd. Final", ListView16.Width / 8
    ListView16.ColumnHeaders.Add , , "Custo Total", ListView16.Width / 8
    ListView16.ColumnHeaders.Add , , "Custo m² com Solvente", ListView16.Width / 8
    
    
'-- MATÉRIA PRIMA
    ListView11.ColumnHeaders.Add , , "Código", ListView11.Width / 9 'gravado
    ListView11.ColumnHeaders.Add , , "Descrição", ListView11.Width / 4 'gravado
    ListView11.ColumnHeaders.Add , , "Material", ListView11.Width / 6 'gravado
    ListView11.ColumnHeaders.Add , , "Dimensão", ListView11.Width / 10 'gravado
    ListView11.ColumnHeaders.Add , , "Q.Unit", ListView11.Width / 16 'gravado
    ListView11.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView11.Width / 7.6 'gravado
    ListView11.ColumnHeaders.Add , , "Q.CJ", ListView11.Width / 19.5 'gravado
    ListView11.ColumnHeaders.Add , , "codigo+material", ListView11.Width / 10000 'gravado
    ListView11.ColumnHeaders.Add , , "Peso Total", ListView11.Width / 7 'calculado
    ListView11.ColumnHeaders.Add , , "sequencia", ListView11.Width / 10000 'gravado
    ListView11.ColumnHeaders.Add , , "Un", ListView11.Width / 28 'gravado
    ListView11.ColumnHeaders.Add , , "Área Pint.", ListView11.Width / 10 'calculado
    ListView11.ColumnHeaders.Add , , "Observação", ListView11.Width / 7 'gravado
    ListView11.ColumnHeaders.Add , , "Peso Especifico", ListView11.Width / 10000 'gravado
    ListView11.ColumnHeaders.Add , , "FO", ListView11.Width / 16 'gravado
    ListView11.ColumnHeaders.Add , , "Comprimento", ListView11.Width / 10000 'calculado
    ListView11.ColumnHeaders.Add , , "Calculo por", ListView11.Width / 10000 'gravado
    ListView11.ColumnHeaders.Add , , "Conjunto", ListView11.Width / 10000 'gravado
    ListView11.ColumnHeaders.Add , , "Peso Posição", ListView11.Width / 10 'gravado
    
'-- RESUMO
    ListView12.ColumnHeaders.Add , , "Item", ListView3.Width / 16
    ListView12.ColumnHeaders.Add , , "Código", ListView3.Width / 16
    ListView12.ColumnHeaders.Add , , "Descrição", ListView3.Width / 5
    ListView12.ColumnHeaders.Add , , "Material", ListView3.Width / 6
    ListView12.ColumnHeaders.Add , , "Un", ListView3.Width / 32
    ListView12.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView3.Width / 7.6
    ListView12.ColumnHeaders.Add , , "Área Pint.", ListView3.Width / 10
    ListView12.ColumnHeaders.Add , , "Comprimento", ListView3.Width / 10000
    ListView12.ColumnHeaders.Add , , "Peso Especifico", ListView3.Width / 10000
    ListView12.ColumnHeaders.Add , , "Observação", ListView3.Width / 5
    
    ListView1.View = lvwReport
    ListView2.View = lvwReport
    ListView3.View = lvwReport
    ListView4.View = lvwReport
    ListView5.View = lvwReport
    ListView6.View = lvwReport
    ListView7.View = lvwReport
    ListView8.View = lvwReport
    ListView9.View = lvwReport
    ListView10.View = lvwReport
    ListView11.View = lvwReport
    ListView12.View = lvwReport
    ListView13.View = lvwReport
    ListView14.View = lvwReport
    ListView15.View = lvwReport
    ListView16.View = lvwReport


End Sub

Private Sub inicializa_tabs()
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    SSTab4.Tab = 0
    SSTab5.Tab = 0
    SSTab6.Tab = 0
    SSTab7.Tab = 0
    
    SubClassSSTAB SSTab1, Picture1
    SubClassSSTAB SSTab2, Picture1
    SubClassSSTAB SSTab3, Picture1
    SubClassSSTAB SSTab4, Picture1
    SubClassSSTAB SSTab5, Picture1
    SubClassSSTAB SSTab6, Picture1
    SubClassSSTAB SSTab7, Picture1

    
End Sub

Private Sub DesbloqueiaControles()
    Dim X As Integer
    
    For X = 0 To 11
        txtcadastro(X).Enabled = True
    Next
    txtcadastro(1).Enabled = False
    txtcadastro(3).Enabled = False
    txtcadastro(13).Enabled = True
    
    'mskCadastro(1).Enabled = True
    'mskCadastro(2).Enabled = False
    For X = 39 To 40
        txtcadastro(X).Enabled = True
    Next
    'cboCadastro.Enabled = True
    
    txtcadastro(24).Enabled = True
    txtcadastro(28).Enabled = True
    txtcadastro(29).Enabled = True
    txtcadastro(30).Enabled = True
    txtcadastro(34).Enabled = True
    DTPicker1.Enabled = True
    Check1.Enabled = True
    optCadastro(0).Enabled = True
    optCadastro(1).Enabled = True
    chamCad(0).Enabled = True
    chamCad(1).Enabled = True
    chamCad(4).Enabled = True
    chameleonButton5.Enabled = True
    chamCad(5).Enabled = True
    'chameleonButton7.Enabled = True
    'chameleonButton8.Enabled = True
    'chameleonButton9.Enabled = True
    'chamCad(3).Enabled = True
    ListView1.Enabled = True
    ListView2.Enabled = True
End Sub

Private Sub BloqueiaControles()
    txtcadastro(6).Enabled = False
End Sub

Private Sub txtCadastro_GotFocus(Index As Integer)
    If Index = 4 Then
        txtcadastro(4).SelStart = 0
        txtcadastro(4).SelLength = Len(txtcadastro(4).Text)
    End If
    Dim X As Integer
    For X = 1 To 11
        txtcadastro(X).SelStart = 0
        txtcadastro(X).SelLength = Len(txtcadastro(X).Text)
    Next
    For X = 13 To 40
        txtcadastro(X).SelStart = 0
        txtcadastro(X).SelLength = Len(txtcadastro(X).Text)
    Next
End Sub

Private Sub txtCadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtcadastro(0) = "" Then ChamaGrid
            CarregaDados (Index)
        End If
    ElseIf Index = 1 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtcadastro(4).SetFocus
        End If
    ElseIf Index = 2 Or Index = 34 Or Index = 35 Then
        If KeyCode = &H8 Then
            txtcadastro(2) = ""
            Formula = ""
            ForPint = ""
            Conta = 0
            CarregaDados (0)
            txtcadastro(2).SetFocus
        End If
        If Conta > 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                If Index = 2 Then CapVar
                If Index = 34 Then CapVar2
                Formula = Replace(Formula, "const0(1)", const0(1))
                Formula = Replace(Formula, "const0(2)", const0(2))
                Formula = Replace(Formula, "const0(3)", const0(3))
                Formula = Replace(Formula, "const0(4)", const0(4))
                Formula = Replace(Formula, "const0(5)", const0(5))
                Formula = Replace(Formula, "const0(6)", const0(6))
                Formula = Replace(Formula, "const0(7)", const0(7))
                Formula = Replace(Formula, "const0(8)", const0(8))
                Formula = Replace(Formula, "const0(9)", const0(9))
                Formula = Replace(Formula, "var0(1)", vAr0(1))
                Formula = Replace(Formula, "var0(2)", vAr0(2))
                Formula = Replace(Formula, "var0(3)", vAr0(3))
                Formula = Replace(Formula, "var0(4)", vAr0(4))
                Formula = Replace(Formula, "var0(5)", vAr0(5))
                Formula = Replace(Formula, "var0(6)", vAr0(6))
                Formula = Replace(Formula, "var0(7)", vAr0(7))
                Formula = Replace(Formula, "var0(8)", vAr0(8))
                Formula = Replace(Formula, "var0(9)", vAr0(9))
                Formula = Replace(Formula, ",", ".")
                
                QuantCJ = Val(txtcadastro(4))
                ForPint = Replace(ForPint, "const0(1)", const0(1))
                ForPint = Replace(ForPint, "const0(2)", const0(2))
                ForPint = Replace(ForPint, "const0(3)", const0(3))
                ForPint = Replace(ForPint, "const0(4)", const0(4))
                ForPint = Replace(ForPint, "const0(5)", const0(5))
                ForPint = Replace(ForPint, "const0(6)", const0(6))
                ForPint = Replace(ForPint, "const0(7)", const0(7))
                ForPint = Replace(ForPint, "const0(8)", const0(8))
                ForPint = Replace(ForPint, "const0(9)", const0(9))
                ForPint = Replace(ForPint, "var0(1)", vAr0(1))
                ForPint = Replace(ForPint, "var0(2)", vAr0(2))
                ForPint = Replace(ForPint, "var0(3)", vAr0(3))
                ForPint = Replace(ForPint, "var0(4)", vAr0(4))
                ForPint = Replace(ForPint, "var0(5)", vAr0(5))
                ForPint = Replace(ForPint, "var0(6)", vAr0(6))
                ForPint = Replace(ForPint, "var0(7)", vAr0(7))
                ForPint = Replace(ForPint, "var0(8)", vAr0(8))
                ForPint = Replace(ForPint, "var0(9)", vAr0(9))
                ForPint = Replace(ForPint, "quantcj", QuantCJ)
                ForPint = Replace(ForPint, ",", ".")
                If Index = 2 Then IncluirItem
                If Index = 34 Or Index = 35 Then IncluirItem2
                Conta = 0
                LimpaControles
                If Index = 2 Then txtcadastro(0).SetFocus
            End If
            If KeyCode = &H6D Then
                If Index = 2 Then CapVar
                If Index = 34 Or Index = 35 Then CapVar2
            End If
        ElseIf Conta = 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                vAr0(1) = Val(txtcadastro(Index))
                Text2.Text = vAr0(1)
                Y = X
                X = Len(txtcadastro(Index))
                Conta = Conta + 1
                Text2.Text = Formula
                txtCadastro_KeyDown Index, 13, 1
            End If
            If KeyCode = &H6D Then 'traço
                vAr0(1) = Val(txtcadastro(Index))
                Text2.Text = vAr0(1)
                Y = X
                X = Len(txtcadastro(Index))
                Conta = Conta + 1
            End If
        End If
        SomaListview
    ElseIf Index = 4 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If optCadastro(0).Value = True Then
                If Val(txtcadastro(0)) <> 0 Then txtcadastro(2).SetFocus
            Else
                txtcadastro(8).SetFocus
            End If
        End If
    ElseIf Index = 6 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtcadastro(6) = "" Or Val(txtcadastro(6)) = 0 Then Exit Sub
            CarregaFO
            txtcadastro(6) = Format(txtcadastro(6), "000000")
            If txtcadastro(13) <> "" Then txtCadastro_KeyDown 13, 13, 13
            If txtcadastro(24) <> "" Then txtCadastro_KeyDown 24, 13, 24
        End If
    ElseIf Index = 7 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            txtcadastro(0).SetFocus
        End If
    ElseIf Index = 8 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            chamCad_Click (0)
            LimpaControles
            txtcadastro(0).SetFocus
        End If
    ElseIf Index = 9 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtcadastro(9) <> "" Then
                CarregaTipoMat
            Else
                txtcadastro(9) = ""
                txtcadastro(10) = ""
            End If
            txtcadastro(0).SetFocus
        End If
    ElseIf Index = 13 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            ChamaGridCli
        End If
    ElseIf Index = 24 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            ChamaGridCont
        End If
    ElseIf Index = 31 Then
        'If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        '    CarregaDados (index)
        'End If
    End If

End Sub

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
End Sub

Private Sub txtCadastro_LostFocus(Index As Integer)
    If Index = 40 Then
        'CalcTotalProposta
    End If
End Sub


Private Sub CarregaDados(Index)
On Error GoTo Err
    Dim X As Integer
    If Index <> 31 Then
        If Val(txtcadastro(0)) = 0 Then
            txtcadastro(1).Enabled = True
            txtcadastro(2).Enabled = False
            txtcadastro(3).Enabled = True
            Check1.Enabled = False
            txtcadastro(5).Enabled = False
            optCadastro(1).Value = True
        
            txtcadastro(3) = "PÇ"
            txtcadastro(1) = "DIGITE O NOME DO MATERIAL"
            txtcadastro(2) = "-"
            Check1.Value = 0
            txtcadastro(1).SetFocus
            txtcadastro(1).BackColor = &HC0C0FF
            txtcadastro(3).BackColor = &HC0C0FF
            txtcadastro(4).BackColor = &HC0C0FF
            txtcadastro(8).BackColor = &HC0C0FF
            Text1.FontBold = True
            Text1.Text = "Item não cadastrado"
            
            If txtcadastro(0) = "000000" Then optCadastro_Click (1)
            Exit Sub
        Else
            txtcadastro(1).Enabled = False
            txtcadastro(2).Enabled = True
            txtcadastro(1).BackColor = &H80000005
            txtcadastro(3).BackColor = &H80000005
            txtcadastro(4).BackColor = &H80000005
            txtcadastro(8).BackColor = &H80000005
            Text1.FontBold = False
            If optCadastro(0).Value = True Then optCadastro_Click (0) Else optCadastro_Click (1)
            Check1.Enabled = True
            Check1.Value = 1
            txtcadastro(5).Enabled = True
        End If
    End If
    
    If Index = 0 Then SqlM = "Select tbMateriais.codmaterial, tbmateriais.descricao, tbMateriais.formula, tbmateriais.constpint, tbconstantes.valconst, tbmateriais.unidade, tbmateriais.forpint, tbmateriais.observacao from tbMateriais Inner Join tbconstantes on tbMateriais.codmaterial = tbConstantes.codmaterial where tbconstantes.codmaterial= '" & Val(txtcadastro(0)) & "'order by tbconstantes.codigo"
    If Index = 31 Then SqlM = "Select tbMateriais.codmaterial, tbmateriais.descricao, tbMateriais.formula, tbmateriais.constpint, tbconstantes.valconst, tbmateriais.unidade, tbmateriais.forpint, tbmateriais.observacao from tbMateriais Inner Join tbconstantes on tbMateriais.codmaterial = tbConstantes.codmaterial where tbconstantes.codmaterial= '" & Val(txtcadastro(31)) & "'order by tbconstantes.codigo"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsMaterial.EOF Then rsMaterial.MoveFirst
    
    If Index = 0 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtcadastro(0)) & "'"
    If Index = 31 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtcadastro(31)) & "'"
    
    If rsMaterial.EOF Then
        If Index = 0 Then txtcadastro(0).Text = Format(txtcadastro(0), "000000") & ""
        If Index = 31 Then txtcadastro(0).Text = Format(txtcadastro(31), "000000") & ""
        Msgbox "Código de material não cadastrado", vbInformation, "Zeus"
    Else
        If Index = 31 Then
            txtcadastro(31).Text = Format(rsMaterial.Fields(0), "000000") & ""
            txtcadastro(32).Text = rsMaterial.Fields(1)
            Formula = rsMaterial.Fields(2)
            ForPint = rsMaterial.Fields(6)
            Text4.Text = rsMaterial.Fields(7)
        End If
        
        If Index = 0 Then
            txtcadastro(0).Text = Format(rsMaterial.Fields(0), "000000") & ""
            txtcadastro(1).Text = rsMaterial.Fields(1)
            Formula = rsMaterial.Fields(2)
            ForPint = rsMaterial.Fields(6)
            txtcadastro(3) = rsMaterial(5)
            txtcadastro(5) = rsMaterial(3)
            Text1.Text = rsMaterial.Fields(7)
            txtcadastro(3).Enabled = False
            txtcadastro(4) = 1
        End If
        For X = 1 To rsMaterial.RecordCount
            const0(X) = rsMaterial.Fields(4)
            rsMaterial.MoveNext
        Next
        If Index = 0 Then txtcadastro(4).SetFocus
    End If
    rsMaterial.Close
    Set rsMaterial = Nothing
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

Private Sub CarregaTipoMat()
On Error GoTo Err
    Dim X As Integer
    Dim rsTipoMat As New ADODB.Recordset
    SqlM = "Select * from tbTipoMat order by tbTipoMat.codigo"
    rsTipoMat.Open SqlM, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsTipoMat.EOF Then rsTipoMat.MoveFirst
    rsTipoMat.Find "codigo=" & "'" & Val(Me.txtcadastro(9)) & "'"
    If rsTipoMat.EOF Then
        txtcadastro(9).Text = Format(txtcadastro(9), "000000") & ""
        Msgbox "Tipo de material não cadastrado", vbInformation, "Zeus"
    Else
        txtcadastro(9).Text = Format(rsTipoMat.Fields(0), "000000") & ""
        txtcadastro(10).Text = rsTipoMat.Fields(1)
        txtcadastro(10).Enabled = False
    End If
    rsTipoMat.Close
    Set rsTipoMat = Nothing
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

Private Sub CarregaFO()
On Error GoTo Err
    Dim rsFO As New ADODB.Recordset
    Dim rsClientes As New ADODB.Recordset
    Dim rsContatos As New ADODB.Recordset
    Dim sqlFO As String
    Dim sqlClientes As String
    Dim sqlContatos As String
    Dim rsLisview As New ADODB.Recordset
    Dim ItemLst As ListItem
    Dim sql As String
    PesoTotal = 0
    SomaTotal = 0
    SomaPint = 0
    
    sqlFO = "select * from tbfo where tbfo.codfo = '" & Val(txtcadastro(6)) & "'"
    rsFO.Open sqlFO, cnBanco, adOpenKeyset, adLockOptimistic
    If rsFO.RecordCount > 0 Then
        DTPicker1 = Format(rsFO(1), "dd/mm/yyyy")
        
        If rsFO.Fields(2) = 1 Then
            Label32 = Label32 & "Em orçamento"
        ElseIf rsFO.Fields(2) = 2 Then
            Label32 = Label32 & "Serviço"
        ElseIf rsFO.Fields(2) = 3 Then
            Label32 = Label32 & "Arquivado"
        End If
        If rsFO.Fields(9) <> "Null" Then DTPicker2 = Format(rsFO(9), "dd/mm/yyyy")
        txtcadastro(13) = Format(rsFO(5), "000000")
        txtcadastro(24) = Format(rsFO(6), "000000")
        txtcadastro(11) = rsFO(4)
        txtcadastro(28) = rsFO(7)
        txtcadastro(30) = rsFO(8)
        If rsFO.Fields(10) <> "Null" Then txtcadastro(39) = rsFO(10)
        txtcadastro(40) = Format(rsFO(11), "#,##0.000;(#,##0.000)")
        If rsFO.Fields(12) <> "Null" Then cboCadastro = rsFO(12)
        'If rsFO.Fields(13) <> "Null" Then mskCadastro(1) = rsFO(13)
        'If rsFO.Fields(13) <> "Null" Then mskCadastro(2) = Format(mskCadastro(1), "#,##0.000;(#,##0.000)") * Format(txtcadastro(40), "#,##0.000;(#,##0.000)")
        If rsFO.Fields(2) = 2 Then
            SkinLabel45 = Format(rsFO.Fields(3), "0000")
        Else
            SkinLabel45 = "####"
        End If
        BloqueiaControles
    Else
        LimpaControles1
        DesbloqueiaControles
        'Check2_Click
        rsFO.Close
        Set rsFO = Nothing
        Exit Sub
    End If
    
    sql = "select a.codfo,a.codseq,a.desenho,a.codmat,a.quantcj,a.dimensoes,a.pesounit,a.area,d.NOMEFANTASIA,d.CODUNDCONTROLE,a.TipoMat,a.revisao,c.descricao[DescTipoMat],a.observacao from tblistamaterial as a left join tbmateriais as b on a.codmat = b.idprd left join tbtipomat as c on a.TipoMat=c.codigo inner join " & vBancoTotvs & ".dbo.tprd as d on b.idprd = d.IDPRD where a.codfo = '" & Val(txtcadastro(6)) & "'order by a.codseq"
    rsLisview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    ListView1.ListItems.Clear
    If rsLisview.RecordCount > 0 Then
        While Not rsLisview.EOF
            'insere o item do arquivo de dados
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsLisview.Fields(3), "000000"))
            'cada item precisa de um subitem para exibir na lista
            ItemLst.SubItems(1) = "" & rsLisview.Fields(8)
            
            If rsLisview.Fields(10) <> 0 Then ItemLst.SubItems(2) = "" & Format(rsLisview.Fields(10), "000000") & "-" & rsLisview.Fields(12) Else ItemLst.SubItems(2) = "-"
            ItemLst.SubItems(3) = "" & rsLisview.Fields(5)
            ItemLst.SubItems(4) = "" & Format(rsLisview.Fields(6), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(5) = "" & rsLisview.Fields(9)
            ItemLst.SubItems(6) = "" & rsLisview.Fields(4)
                
            PesoTotal = rsLisview.Fields(4) * rsLisview.Fields(6)
            ItemLst.SubItems(7) = "" & Format(PesoTotal, "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(8) = "" & Format(rsLisview.Fields(7), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(9) = "" & rsLisview.Fields(2)
            'ItemLst.SubItems(10) = Format(rsLisview.Fields(3), "000000") & rsLisview.Fields(10)
            ItemLst.SubItems(11) = "" & rsLisview.Fields(11)
            
            If rsLisview.Fields(3) = 0 Then
                ItemLst.SubItems(1) = "" & rsLisview.Fields(13)
                ItemLst.SubItems(5) = "PÇ"
            End If
            If ItemLst.SubItems(2) <> "-" Then ItemLst.SubItems(10) = ItemLst.SubItems(1) & ItemLst.SubItems(2) Else ItemLst.SubItems(10) = ItemLst.SubItems(1)
            ItemLst.SubItems(12) = Format(rsLisview.Fields(1), "0000")
            SomaTotal = SomaTotal + PesoTotal
            SomaPint = SomaPint + rsLisview.Fields(7)
            PesoTotal = 0
            'vai para o proximo registro
            rsLisview.MoveNext
        Wend
    End If
    
    lblTotal = Format(SomaTotal, "#,##0.000;(#,##0.000)")
    lblTotPint = Format(SomaPint, "#,##0.000;(#,##0.000)")
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    rsLisview.Close
    ListView1.Refresh
    Set rsLisview = Nothing
    ListView2.ListItems.Clear
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        Debug.Print Err.Number & " - " & Err.Description
        Resume Next
    End If
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

Private Function ValidaCampo()
    ValidaCampo = False
    If txtcadastro(6).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(6).Tag, vbInformation, "Atenção"
        Me.txtcadastro(6).SetFocus
        Exit Function
    End If
    If txtcadastro(30).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(30).Tag, vbInformation, "Atenção"
        Me.txtcadastro(30).SetFocus
        Exit Function
    End If
    If txtcadastro(11).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(11).Tag, vbInformation, "Atenção"
        Me.txtcadastro(11).SetFocus
        Exit Function
    End If
    If txtcadastro(13).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(13).Tag, vbInformation, "Atenção"
        Me.txtcadastro(13).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function ValidaCampo2()
    ValidaCampo2 = False
    If txtcadastro(0).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(0).Tag, vbInformation, "Atenção"
        Me.txtcadastro(0).SetFocus
        Exit Function
    End If
    If txtcadastro(4).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(4).Tag, vbInformation, "Atenção"
        Me.txtcadastro(4).SetFocus
        Exit Function
    End If
    If txtcadastro(3).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(3).Tag, vbInformation, "Atenção"
        Me.txtcadastro(3).SetFocus
        Exit Function
    End If
    
    If optCadastro(0).Value = True And txtcadastro(2).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(2).Tag, vbInformation, "Atenção"
        Me.txtcadastro(2).SetFocus
        Exit Function
    End If
    If optCadastro(1).Value = True And txtcadastro(8).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(8).Tag, vbInformation, "Atenção"
        Me.txtcadastro(8).SetFocus
        Exit Function
    End If
    If Check1.Value = 1 And txtcadastro(5).Text = "" Then
        Msgbox "Favor informar o campo " & Me.txtcadastro(5).Tag, vbInformation, "Atenção"
        Me.txtcadastro(5).SetFocus
        Exit Function
    End If
    
    ValidaCampo2 = True
End Function

Private Sub CapVar()
    vAr0(Conta + 1) = Val(Mid$(txtcadastro(2), X + 2, Len(txtcadastro(2)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtcadastro(2))
    Conta = Conta + 1
End Sub

Private Sub CapVar2()
    vAr0(Conta + 1) = Val(Mid$(txtcadastro(34), X + 2, Len(txtcadastro(34)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtcadastro(34))
    Conta = Conta + 1
End Sub

Private Sub ChamaGrid()
    'Dim F As New frmpesqcli
    Sqlp = "Select * from tbmateriais where tbmateriais.descricao like '%" & txtcadastro(1) & "%'"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    F.Caption = "Pesquisa de Materiais"
    Pesquisa = frmFO.Tag
    F.Show 1
    If Pesquisa <> "0" Then
        txtcadastro(0) = Pesquisa
    End If
End Sub

Private Sub ChamaGridMat()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTipoMat order by descricao"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Tipo de Materiais"
    Pesquisa = frmFO.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "descricao=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(9).Text = Format(rsLocal.Fields(0), "000000")
        Else
            Msgbox "Tipo de material não cadastrado", vbInformation, "Zeus"
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

Private Sub ChamaGridCli()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbclifor order by nome"
    procnom = "nome"
    campo = 13
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa Clientes"
    Pesquisa = frmFO.Tag
    If txtcadastro(13) = "" Then F.Show 1
    Pesquisa = Mid$(Pesquisa, 7, 85)
    If Pesquisa <> "" And txtcadastro(13) = "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nome=" & "'" & Pesquisa & "'"
        If rsLocal.EOF Then
            Msgbox "Cliente não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    Else
        Sqlp = "Select * from tbclifor where tbclifor.codclifor = '" & Val(txtcadastro(13)) & "'"
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.EOF Then
            'MsgBox "Cliente não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    End If
    txtcadastro(13).Text = Format(rsLocal.Fields(0), "000000")
    txtcadastro(14).Text = rsLocal.Fields(13)
    txtcadastro(15).Text = rsLocal.Fields(1)
    txtcadastro(16).Text = rsLocal.Fields(2)
    txtcadastro(17).Text = rsLocal.Fields(3)
    txtcadastro(18).Text = rsLocal.Fields(4)
    txtcadastro(19).Text = rsLocal.Fields(5)
    txtcadastro(20).Text = Format(rsLocal.Fields(6), "(##)####-####")
    txtcadastro(21).Text = Format(rsLocal.Fields(7), "(##)####-####")
    txtcadastro(22).Text = rsLocal.Fields(8)
    txtcadastro(23).Text = rsLocal.Fields(9)
    rsLocal.Close
    Set rsLocal = Nothing
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

Private Sub ChamaGridCont()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcontatos where tbcontatos.codclifor= '" & Val(txtcadastro(13)) & "'order by nome"
    procnom = "nome"
    campo = 2
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa Contatos"
    Pesquisa = frmFO.Tag
    If txtcadastro(24) = "" Then F.Show 1
    If Pesquisa <> "" And txtcadastro(24) = "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nome=" & "'" & Mid$(Pesquisa, 7, 100) & "'"
        If rsLocal.EOF Then
            'MsgBox "Contato não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    Else
        Sqlp = "select * from tbcontatos where tbcontatos.codclifor = '" & Val(txtcadastro(13)) & "'" & _
        "and tbcontatos.codcontato=" & " '" & Val(txtcadastro(24)) & "'"
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.EOF Then
            'MsgBox "Contato não cadastrado", vbInformation, "Zeus"
            rsLocal.Close
            Set rsLocal = Nothing
            Exit Sub
        End If
    End If
    txtcadastro(24).Text = Format(rsLocal.Fields(1), "000000")
    txtcadastro(25).Text = rsLocal.Fields(2)
    txtcadastro(26).Text = Format(rsLocal.Fields(6), "(##)####-####")
    txtcadastro(27).Text = rsLocal.Fields(9)
    rsLocal.Close
    Set rsLocal = Nothing
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


Private Sub IncluirItem()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, ProxSeq As Integer
    If ValidaCampo2 = False Then Exit Sub
    
    'Variavel do sistema para calculo da area de pintura, consta na formula de pintura
    If optCadastro(0).Value = True Then
        PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(4), "#,##0.000;(#,##0.000)")
    Else
        PesoTotal = Format(txtcadastro(8) * txtcadastro(4), "#,##0.000;(#,##0.000)")
    End If
    Y = ListView1.ListItems.Count
    
    If Label36.Caption = "Alteração" Then
        ListView1.ListItems(Val(Label37)).Selected = True
        ListView1.ListItems(Val(Label37)).EnsureVisible
    
        'ListView1.SelectedItem.ListSubItems.Item (1)
        Label36.Caption = "Inclusão"
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            If Me.txtcadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Or Me.txtcadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(3) Then
                ListView1.SelectedItem.ListSubItems.Item(8) = Format(ScriptControl1.Eval(ForPint) * txtcadastro(5), "#,##0.000;(#,##0.000)")
            End If
        End If
        ListView1.SelectedItem.ListSubItems.Item(1) = Me.txtcadastro(1).Text
        ListView1.SelectedItem.ListSubItems.Item(2) = Me.txtcadastro(9).Text & "-" & Me.txtcadastro(10).Text
        
        If Me.txtcadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Then
'        If Me.txtcadastro(2) <> ListView1.SelectedItem.ListSubItems.Item(3) Or Me.txtcadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(4) Then
            If optCadastro(0).Value = True Then
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(4), "#,##0.000;(#,##0.000)")
            Else
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(txtcadastro(8), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            End If
        End If
        
        If Me.txtcadastro(8) <> ListView1.SelectedItem.ListSubItems.Item(4) Then
            If optCadastro(1).Value = True Then
                ListView1.SelectedItem.ListSubItems.Item(4) = Format(txtcadastro(8), "#,##0.000;(#,##0.000)")
                ListView1.SelectedItem.ListSubItems.Item(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            End If
        End If
        
        ListView1.SelectedItem.ListSubItems.Item(3) = Me.txtcadastro(2).Text
        ListView1.SelectedItem.ListSubItems.Item(5) = Me.txtcadastro(3).Text
        ListView1.SelectedItem.ListSubItems.Item(6) = Me.txtcadastro(4).Text
        ListView1.SelectedItem.ListSubItems.Item(9) = Me.txtcadastro(7).Text
        ListView1.SelectedItem.ListSubItems.Item(12) = Format(Label37, "0000") 'Me.txtcadastro(1).Text & Me.txtcadastro(10).Text
        ListView1.SelectedItem.ListSubItems.Item(11) = Me.txtcadastro(29).Text
    Else
        'Ordena Listview pela sequencia de cadastramento antes de gravar
        Me.ListView1.Sorted = True
        Me.ListView1.SortKey = 12
        Me.ListView1.SortOrder = lvwAscending
        '------
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(ListView1.ListItems.Count).Selected = True
            ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
            ProxSeq = Val(ListView1.SelectedItem.ListSubItems.Item(12)) + 1
        Else
            ProxSeq = 1
        End If
        
        Set ItemLst = ListView1.ListItems.Add(, , Format(txtcadastro(0), "000000"))
        Label36.Caption = "Inclusão"
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            ItemLst.SubItems(8) = Format(ScriptControl1.Eval(ForPint) * txtcadastro(5), "#,##0.000;(#,##0.000)")
        End If
        ItemLst.SubItems(1) = Me.txtcadastro(1).Text
        ItemLst.SubItems(2) = Me.txtcadastro(9).Text & "-" & Me.txtcadastro(10).Text
        ItemLst.SubItems(3) = Me.txtcadastro(2).Text
        If optCadastro(0).Value = True Then
            ItemLst.SubItems(4) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(7) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(4), "#,##0.000;(#,##0.000)")
        Else
            ItemLst.SubItems(4) = Format(txtcadastro(8), "#,##0.000;(#,##0.000)")
            ItemLst.SubItems(7) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
        End If
        ItemLst.SubItems(5) = Me.txtcadastro(3).Text
        ItemLst.SubItems(6) = Me.txtcadastro(4).Text
        ItemLst.SubItems(9) = Me.txtcadastro(7).Text
        ItemLst.SubItems(10) = Me.txtcadastro(1).Text & Me.txtcadastro(10).Text
        ItemLst.SubItems(11) = Me.txtcadastro(29).Text
        ItemLst.SubItems(12) = Format(ProxSeq, "0000")
        
        ListView1.ListItems(ListView1.ListItems.Count).Selected = True
        ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
        
    End If
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    If txtcadastro(3) <> "pç" And txtcadastro(3) <> "PÇ" Then
        If optCadastro(0).Value = True Then
            SomaTotal = SomaTotal + ScriptControl1.Eval(Formula) * Me.txtcadastro(4)
        Else
            SomaTotal = Format(SomaTotal + txtcadastro(8) * Me.txtcadastro(4), "#,##0.000;(#,##0.000")
        End If
        
        If Check1.Value = 1 Then SomaPint = SomaPint + ScriptControl1.Eval(ForPint) * Me.txtcadastro(5)
    End If
    'lblTotal.Caption = Format(SomaTotal, "#,##0.0;(#,##0.0)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    'lblTotPint.Caption = Format(SomaPint, "#,##0.00;(#,##0.00)")
    txtcadastro(0) = ""
    txtcadastro(1) = ""
    txtcadastro(2) = ""
    Text1.Text = ""
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, verifique se as dimensões digitadas estão de acordo com as referidas na fórmula!", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub IncluirItem2()
On Error GoTo TrataErro
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaCampo = False Then Exit Sub
    
    'Variavel do sistema para calculo da area de pintura, consta na formula de pintura
    If optCadastro(2).Value = True Then
        PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(37), "#,##0.000;(#,##0.000)")
    Else
        PesoTotal = Format(txtcadastro(35), "#,##0.000;(#,##0.000)")
    End If
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView2.ListItems.Item(X) = Me.txtcadastro(38) Then
                If optCadastro(2).Value = True Then
                    ListView2.SelectedItem.ListSubItems.Item(7) = Me.txtcadastro(34).Text
                    ListView2.SelectedItem.ListSubItems.Item(8) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(37), "#,##0.000;(#,##0.000)")
                    ListView2.SelectedItem.ListSubItems.Item(9) = 0
                    ListView2.SelectedItem.ListSubItems.Item(10) = txtcadastro(37)
                Else
                    ListView2.SelectedItem.ListSubItems.Item(7) = "-"
                    If Check3.Value = 1 Then ListView2.SelectedItem.ListSubItems.Item(8) = Format((txtcadastro(35) * txtcadastro(36) / 100) + txtcadastro(35), "#,##0.000;(#,##0.000)")
                    If Check3.Value = 0 Then ListView2.SelectedItem.ListSubItems.Item(8) = Format(txtcadastro(35), "#,##0.000;(#,##0.000)")
                    If Check3.Value = 1 Then ListView2.SelectedItem.ListSubItems.Item(9) = txtcadastro(36)
                    ListView2.SelectedItem.ListSubItems.Item(10) = 1
                End If
                ListView2.SelectedItem.ListSubItems.Item(11) = "Alterado"
                
                Exit For
            End If
        Next
    End If
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView1.ColumnHeaders(9).Alignment = lvwColumnRight
    Text4.Text = ""
    For X = 31 To 38
        txtcadastro(X) = ""
    Next
    LimpaControles
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, verifique se as dimensões digitadas estão de acordo com as referidas na fórmula!", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub ExcluirItem()
On Error GoTo TrataErro
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    ListView1.ListItems.Remove (X)
    Exit Sub
TrataErro:
    Msgbox "Ocorreu um erro, Selecione um item antes de excluir", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Sub AlterarItem()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtcadastro(31).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.txtcadastro(32).Text = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.txtcadastro(33).Text = ListView2.SelectedItem.ListSubItems.Item(4)
    If ListView2.SelectedItem.ListSubItems.Item(8) = 0 Then Me.txtcadastro(35).Text = ListView2.SelectedItem.ListSubItems.Item(5) Else Me.txtcadastro(35).Text = ListView2.SelectedItem.ListSubItems.Item(8)
    If ListView2.SelectedItem.ListSubItems.Item(9) = 0 Then Me.txtcadastro(36).Text = 5 Else Me.txtcadastro(36).Text = ListView2.SelectedItem.ListSubItems.Item(9)
    Me.txtcadastro(37).Text = ListView2.SelectedItem.ListSubItems.Item(10)
    Me.txtcadastro(38).Text = ListView2.ListItems.Item(X)
    CarregaDados (31)
    txtcadastro(31).Enabled = False
    txtcadastro(32).Enabled = False
    txtcadastro(33).Enabled = False
End Sub

Private Sub AlterarItem1()
    Label36.Caption = "Alteração"
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Label37 = Val(ListView1.SelectedItem.ListSubItems.Item(12))
    Me.txtcadastro(7) = ListView1.SelectedItem.ListSubItems.Item(9)
    Me.txtcadastro(29) = ListView1.SelectedItem.ListSubItems.Item(11)
    Me.txtcadastro(9) = Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 6)
    Me.txtcadastro(0) = ListView1.ListItems.Item(X)
    Me.txtcadastro(1) = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.txtcadastro(3) = ListView1.SelectedItem.ListSubItems.Item(5)
    Me.txtcadastro(2) = ListView1.SelectedItem.ListSubItems.Item(3)
    Me.txtcadastro(4) = ListView1.SelectedItem.ListSubItems.Item(6)
    
    If ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
        Me.optCadastro(1).Value = True
        optCadastro_Click (1)
        Me.txtcadastro(8) = ListView1.SelectedItem.ListSubItems.Item(4)
        txtcadastro(8).SetFocus
        Check1.Value = 0
        txtcadastro(5).Text = ""
        If Check1.Value = 0 Then txtcadastro(5).Enabled = False
        txtcadastro(0).BackColor = &HC0C0FF
    End If
    If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then
        Me.optCadastro(0).Value = True
        optCadastro_Click (0)
        txtcadastro(2).SetFocus
        txtcadastro(0).BackColor = &H80000005
    End If
    
    If Val(txtcadastro(0)) = 0 Then
        optCadastro_Click (1)
    End If
    
    If ListView1.SelectedItem.ListSubItems.Item(8) = "0,00" Then Check1.Value = 0 Else Check1.Value = 1
    txtCadastro_KeyDown 0, 13, 0
    txtCadastro_KeyDown 9, 13, 9
    
    If txtcadastro(0) = "000000" Then
        txtcadastro(8).BackColor = &HC0C0FF
        Me.txtcadastro(1) = ListView1.SelectedItem.ListSubItems.Item(1)
    End If

End Sub

Private Sub GravarDados()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    Dim rsDeleta As New ADODB.Recordset
    Dim rsGravaLM As New ADODB.Recordset
    Dim rsGravaFO As New ADODB.Recordset
    Dim rsGravaResumo As New ADODB.Recordset
    
    Dim sqlExc As String
    Dim sql As String
    Dim Y As Integer, X As Integer
10  cnBanco.BeginTrans

    sql = "Select * from tbListaMaterial order by codfo"
    rsGravaLM.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
    sqlExc = "Delete from tbListaMaterial where tbListaMaterial.codfo = '" & Val(txtcadastro(6)) & "'"
    rsDeleta.Open sqlExc, cnBanco
    
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        rsGravaLM.AddNew
        rsGravaLM(0) = txtcadastro(6)
        rsGravaLM(1) = ListView1.SelectedItem.ListSubItems.Item(12)
        rsGravaLM(2) = ListView1.SelectedItem.ListSubItems.Item(9)
        
        rsGravaLM(3) = Val(ListView1.ListItems.Item(X))
        rsGravaLM(4) = ListView1.SelectedItem.ListSubItems.Item(6)
        rsGravaLM(5) = ListView1.SelectedItem.ListSubItems.Item(3)
        rsGravaLM(6) = ListView1.SelectedItem.ListSubItems.Item(4)
        If ListView1.SelectedItem.ListSubItems.Item(8) <> "" Then rsGravaLM(7) = ListView1.SelectedItem.ListSubItems.Item(8) Else rsGravaLM(7) = 0
        rsGravaLM(8) = Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 6))
        rsGravaLM(9) = ListView1.SelectedItem.ListSubItems.Item(11)
        If Val(ListView1.ListItems.Item(X)) = 0 Then rsGravaLM(10) = ListView1.SelectedItem.ListSubItems.Item(1)
    Next
    If Not rsGravaLM.EOF Then rsGravaLM.Update
    
    sql = "Select * from tbFo where tbfo.codfo = '" & Val(txtcadastro(6)) & "'"
    rsGravaFO.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsGravaFO.EOF Then
        rsGravaFO.AddNew
    End If
    rsGravaFO(0) = txtcadastro(6)
    rsGravaFO(1) = Format(DTPicker1, "dd/mm/yyyy")

    If rsGravaFO(2) = "" Then rsGravaFO(2) = 1
    If rsGravaFO(2) = 1 Then rsGravaFO(2) = 1
    If rsGravaFO(2) = 2 Then rsGravaFO(2) = 2
    If rsGravaFO(2) = 3 Then rsGravaFO(2) = 3
    
    rsGravaFO(4) = txtcadastro(11)
    If txtcadastro(13) <> "" Then rsGravaFO(5) = txtcadastro(13)
    If txtcadastro(24) <> "" Then rsGravaFO(6) = txtcadastro(24)
    rsGravaFO(7) = txtcadastro(28)
    rsGravaFO(8) = txtcadastro(30)
    If DTPicker2 <> "" Then
        rsGravaFO(9) = Format(DTPicker2, "dd/mm/yyyy")
    End If
    rsGravaFO(10) = txtcadastro(39)
    If txtcadastro(40) <> "" Then rsGravaFO(11) = Format(txtcadastro(40), "#,##0.000;(#,##0.000)")
    rsGravaFO(12) = cboCadastro
    'mskCadastro(1).PromptInclude = False
    'If mskCadastro(1) <> "" Then rsGravaFO(13) = mskCadastro(1)
    'mskCadastro(1).PromptInclude = True
    rsGravaFO(14) = "S"
    
    If Not rsGravaFO.EOF Then rsGravaFO.Update

'***************** INICIO GRAVAR DADOS DA TABELA DE RESUMO ***************
'Grava apenas se houver alguma informação na tabela de resumo
    If ListView2.ListItems.Count <> 0 Then
        
        sqlExc = "Delete from tbResumo where tbResumo.codfo = '" & Val(txtcadastro(6)) & "'"
        rsDeleta.Open sqlExc, cnBanco
        
        sql = "Select * from tbResumo where tbResumo.codfo = '" & Val(txtcadastro(6)) & "'"
        rsResumo.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    
        
        Y = ListView2.ListItems.Count
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            
            
            rsResumo.AddNew
            rsResumo.Fields(0) = txtcadastro(6)
            rsResumo.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
            rsResumo.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(7)
            rsResumo.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(8)
            rsResumo.Fields(4) = ListView2.SelectedItem.ListSubItems.Item(9)
            rsResumo.Fields(5) = ListView2.SelectedItem.ListSubItems.Item(10)
            rsResumo.Fields(6) = ListView2.SelectedItem.ListSubItems.Item(5)
            If ListView2.SelectedItem.ListSubItems.Item(4) <> "-" Then rsResumo.Fields(7) = Val(Mid$(ListView2.SelectedItem.ListSubItems.Item(3), 1, 6)) Else rsResumo.Fields(7) = 0
            rsResumo.Fields(8) = X
'            If Val(ListView2.SelectedItem.ListSubItems.Item(1)) = 0 Then rsResumo.Fields(8) = X Else rsResumo.Fields(8) = Val(ListView2.SelectedItem.ListSubItems.Item(1))
            If Val(ListView2.SelectedItem.ListSubItems.Item(1)) = 0 Then rsResumo.Fields(9) = ListView2.SelectedItem.ListSubItems.Item(2)
        
        Next
        If Not rsResumo.EOF Then rsResumo.Update
        rsResumo.Close
    End If

'***************** FIM GRAVAR DADOS DA TABELA DE RESUMO ******************
    
    cnBanco.CommitTrans
    rsGravaFO.Close
    rsGravaLM.Close
    'AtualizaListview
    Msgbox "Dados gravados com sucesso", vbInformation, "Zeus"
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

Private Sub SomaListview()
    Dim SomaT As Currency, SomaP As Currency
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).SubItems(5) <> "pç" And ListView1.ListItems(i).SubItems(5) <> "PÇ" Then
            If ListView1.ListItems(i).SubItems(7) <> "" Then SomaT = SomaT + CCur(ListView1.ListItems(i).SubItems(7)) 'coluna de valores
            If ListView1.ListItems(i).SubItems(8) <> "" Then SomaP = SomaP + CCur(ListView1.ListItems(i).SubItems(8)) 'coluna de valores
        End If
    Next
    lblTotal.Caption = Format(SomaT, "#,##0.000;(#,##0.000)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    lblTotPint.Caption = Format(SomaP, "#,##0.000;(#,##0.000)")
End Sub

