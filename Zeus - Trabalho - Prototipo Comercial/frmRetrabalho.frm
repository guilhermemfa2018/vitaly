VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmRetrabalho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OS de RETRABALHO"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18480
   Icon            =   "frmRetrabalho.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   18480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   13
      Left            =   720
      Picture         =   "frmRetrabalho.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9840
      Width           =   615
   End
   Begin VB.Frame Frame6 
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
      Height          =   1455
      Left            =   8160
      TabIndex        =   24
      Top             =   120
      Width           =   8055
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   5520
         TabIndex        =   105
         Text            =   "Desenho"
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtformula 
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
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   480
         Width           =   1695
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
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":1994
         TabIndex        =   30
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtformula 
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
         TabIndex        =   29
         Top             =   1080
         Width           =   5415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":1A08
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
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
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   480
         Width           =   375
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "frmRetrabalho.frx":1A74
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame21 
      Caption         =   "Dados Retrabalho"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   7440
         TabIndex        =   22
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   7440
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtformula 
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
         HelpContextID   =   1
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtformula 
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
         Index           =   12
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtformula 
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
         Index           =   13
         Left            =   3240
         TabIndex        =   12
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtformula 
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
         Left            =   1680
         TabIndex        =   11
         Tag             =   "Nome do Responsável"
         ToolTipText     =   "Nome do Responsável"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtformula 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   4200
         TabIndex        =   10
         Text            =   "ID Projeto"
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   90112001
         CurrentDate     =   41554
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmRetrabalho.frx":1AD8
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "frmRetrabalho.frx":1B48
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   1680
         OleObjectBlob   =   "frmRetrabalho.frx":1BB0
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":1C10
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":1C86
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   120
      Picture         =   "frmRetrabalho.frx":1CFC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9840
      Width           =   615
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
      Left            =   16320
      TabIndex        =   5
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
         TabIndex        =   6
         Text            =   "-"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox label53 
      Height          =   330
      Left            =   7080
      TabIndex        =   4
      Text            =   "-"
      Top             =   10080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   9960
      Width           =   2655
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
      Left            =   13800
      TabIndex        =   1
      Top             =   9840
      Visible         =   0   'False
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":29C6
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   375
      Left            =   5160
      OleObjectBlob   =   "frmRetrabalho.frx":2A20
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   8295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   32
      Top             =   1680
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   14208
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
      TabPicture(0)   =   "frmRetrabalho.frx":2AEE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdCadastro(4)"
      Tab(0).Control(1)=   "cmdCadastro(8)"
      Tab(0).Control(2)=   "cmdCadastro(7)"
      Tab(0).Control(3)=   "Frame31"
      Tab(0).Control(4)=   "Frame11"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Recursos"
      TabPicture(1)   =   "frmRetrabalho.frx":2B0A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "aicAlphaImage2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ScriptControl1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ListView1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "SkinLabel6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "SkinLabel7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SkinLabel18"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "SkinLabel16"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame17"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdCadastro(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdCadastro(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Frame7"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Frame13"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame14"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtformula(0)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtformula(1)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdCadastro(10)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtformula(25)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtDB"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtLV"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtformula(26)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Combo1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "cmdCadastro(17)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmdCadastro(2)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmdCadastro(3)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Frame25"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Ordem de Serviço"
      TabPicture(2)   =   "frmRetrabalho.frx":2B26
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame15"
      Tab(2).Control(1)=   "Frame16"
      Tab(2).ControlCount=   2
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
         Left            =   15240
         TabIndex        =   106
         Top             =   7320
         Width           =   2295
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":2B42
            TabIndex        =   107
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   3
         Left            =   1320
         Picture         =   "frmRetrabalho.frx":2B9C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   720
         Picture         =   "frmRetrabalho.frx":3866
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "Agregar"
         Height          =   495
         Index           =   17
         Left            =   8280
         TabIndex        =   98
         Tag             =   "Inclui um item selecionado no LV à uma OS"
         ToolTipText     =   "Inclui um item selecionado no LV à uma OS"
         Top             =   7440
         Width           =   1455
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
         ItemData        =   "frmRetrabalho.frx":4530
         Left            =   14400
         List            =   "frmRetrabalho.frx":4552
         TabIndex        =   97
         Tag             =   "Operação nº"
         ToolTipText     =   "Operação nº"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtformula 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Index           =   26
         Left            =   120
         TabIndex        =   96
         Tag             =   "Observação"
         ToolTipText     =   "Observação"
         Top             =   1200
         Width           =   9735
      End
      Begin VB.TextBox txtLV 
         Height          =   330
         Left            =   3000
         TabIndex        =   95
         Text            =   "LV"
         Top             =   7320
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtDB 
         Height          =   330
         Left            =   1560
         TabIndex        =   94
         Text            =   "DB"
         Top             =   7680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtformula 
         Height          =   330
         Index           =   25
         Left            =   1560
         TabIndex        =   93
         Text            =   "ID OS"
         Top             =   7320
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
         Height          =   7335
         Left            =   -64440
         TabIndex        =   78
         Top             =   600
         Width           =   7575
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   16
            Left            =   6720
            Picture         =   "frmRetrabalho.frx":457F
            Style           =   1  'Graphical
            TabIndex        =   87
            Tag             =   "Cadastrar Serviços de Terceiros"
            ToolTipText     =   "Cadastrar Serviços de Terceiros"
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   11
            Left            =   120
            Picture         =   "frmRetrabalho.frx":5249
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox txtformula 
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
            Index           =   24
            Left            =   120
            TabIndex        =   85
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox txtformula 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   23
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   84
            Top             =   1200
            Width           =   7335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   255
            Left            =   7080
            TabIndex        =   83
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtformula 
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
            Index           =   22
            Left            =   840
            TabIndex        =   82
            Top             =   480
            Width           =   6135
         End
         Begin VB.TextBox txtformula 
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
            Index           =   21
            Left            =   120
            TabIndex        =   81
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   14
            Left            =   720
            Picture         =   "frmRetrabalho.frx":5F13
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   2760
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   15
            Left            =   1320
            Picture         =   "frmRetrabalho.frx":6BDD
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   2760
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   3735
            Left            =   120
            TabIndex        =   88
            Top             =   3480
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   6588
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
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":78A7
            TabIndex        =   89
            Top             =   2160
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":7915
            TabIndex        =   90
            Top             =   960
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmRetrabalho.frx":7981
            TabIndex        =   91
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":79E9
            TabIndex        =   92
            Top             =   240
            Width           =   1575
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
         Height          =   7455
         Left            =   -74880
         TabIndex        =   67
         Top             =   480
         Width           =   10335
         Begin VB.TextBox txtformula 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6255
            Index           =   20
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   71
            Top             =   1080
            Width           =   10095
         End
         Begin VB.TextBox txtformula 
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
            Index           =   19
            Left            =   4440
            TabIndex        =   70
            Top             =   480
            Width           =   5775
         End
         Begin VB.TextBox txtformula 
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
            Index           =   18
            Left            =   1920
            TabIndex        =   69
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtformula 
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
            Index           =   17
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":7A49
            TabIndex        =   72
            Top             =   840
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   2760
            TabIndex        =   73
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
            Format          =   215220225
            CurrentDate     =   41568
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "frmRetrabalho.frx":7AB7
            TabIndex        =   74
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   4440
            OleObjectBlob   =   "frmRetrabalho.frx":7B19
            TabIndex        =   75
            Top             =   240
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "frmRetrabalho.frx":7B91
            TabIndex        =   76
            Top             =   240
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":7BF9
            TabIndex        =   77
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   10
         Left            =   240
         Picture         =   "frmRetrabalho.frx":7C57
         Style           =   1  'Graphical
         TabIndex        =   66
         Tag             =   "Gerar OS - Ordem de Serviço"
         ToolTipText     =   "Gerar OS - Ordem de Serviço"
         Top             =   7320
         Visible         =   0   'False
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
         TabIndex        =   65
         Top             =   600
         Width           =   11535
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
         TabIndex        =   64
         Tag             =   "ID Centro de Custo"
         ToolTipText     =   "ID Centro de Custo"
         Top             =   600
         Width           =   1935
      End
      Begin VB.Frame Frame14 
         Caption         =   "Tempo total "
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
         Left            =   11880
         TabIndex        =   62
         Top             =   7320
         Width           =   3255
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
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   195
            Width           =   2895
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
         Height          =   855
         Left            =   4320
         TabIndex        =   60
         Top             =   3000
         Width           =   5535
         Begin ACTIVESKINLibCtl.SkinLabel Label8 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "frmRetrabalho.frx":8921
            TabIndex        =   61
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data Prevista"
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
         TabIndex        =   58
         Top             =   3000
         Width           =   1815
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   405
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
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
            Format          =   215220225
            CurrentDate     =   41556
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   4
         Left            =   -69720
         Picture         =   "frmRetrabalho.frx":897B
         Style           =   1  'Graphical
         TabIndex        =   57
         Tag             =   "Limpar Controles"
         ToolTipText     =   "Limpar Controles"
         Top             =   7200
         Width           =   615
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   56
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastro 
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmRetrabalho.frx":9645
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3960
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
         Height          =   3615
         Left            =   9960
         TabIndex        =   53
         Top             =   960
         Width           =   4215
         Begin VB.PictureBox Picture1 
            Height          =   3255
            Left            =   120
            ScaleHeight     =   3195
            ScaleWidth      =   3915
            TabIndex        =   54
            Top             =   240
            Width           =   3975
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   3255
               Left            =   0
               Top             =   0
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   5741
               Image           =   "frmRetrabalho.frx":A30F
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
         Height          =   3615
         Left            =   14280
         TabIndex        =   50
         Top             =   960
         Width           =   3855
         Begin VB.TextBox txtformula 
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   51
            Top             =   3000
            Visible         =   0   'False
            Width           =   1815
         End
         Begin MSComctlLib.TreeView TreeView3 
            Height          =   3135
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   5530
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
      Begin VB.Frame Frame5 
         Caption         =   "Tempo previsto"
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
         TabIndex        =   48
         Top             =   3000
         Width           =   2175
         Begin VB.TextBox txtformula 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Tag             =   "Insira as variáveis de acordo com a Observação acima"
            ToolTipText     =   "Insira as variáveis de acordo com a Observação acima"
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   "<"
         Height          =   615
         Index           =   8
         Left            =   -69720
         TabIndex        =   47
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdCadastro 
         Caption         =   ">"
         Height          =   615
         Index           =   7
         Left            =   -69720
         TabIndex        =   46
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
         Left            =   -68880
         TabIndex        =   41
         Top             =   480
         Width           =   12015
         Begin VB.TextBox txtLvw 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   10680
            TabIndex        =   104
            Top             =   120
            Visible         =   0   'False
            Width           =   615
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
            TabIndex        =   44
            Top             =   6600
            Width           =   11775
            Begin ACTIVESKINLibCtl.SkinLabel Label3 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmRetrabalho.frx":A327
               TabIndex        =   45
               Top             =   240
               Width           =   11535
            End
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Tag             =   "Itens selecionados"
            ToolTipText     =   "Itens selecionados"
            Top             =   6120
            Visible         =   0   'False
            Width           =   11655
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   6375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   11245
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
         TabIndex        =   37
         Top             =   480
         Width           =   5055
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
            TabIndex        =   38
            Top             =   6600
            Width           =   4815
            Begin ACTIVESKINLibCtl.SkinLabel Label6 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmRetrabalho.frx":A381
               TabIndex        =   39
               Top             =   240
               Width           =   4455
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   6255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   4815
            _ExtentX        =   8493
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
      Begin VB.Frame Frame17 
         Caption         =   "Status (0-nada/1-aberta/2-andamento/3-fechada)"
         Height          =   615
         Left            =   3600
         TabIndex        =   35
         Top             =   7320
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox txtformula 
            Height          =   330
            Index           =   27
            Left            =   120
            TabIndex        =   36
            Text            =   "0"
            Top             =   240
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   14400
         OleObjectBlob   =   "frmRetrabalho.frx":A3DB
         TabIndex        =   99
         Top             =   360
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":A44B
         TabIndex        =   100
         Top             =   960
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   2760
         OleObjectBlob   =   "frmRetrabalho.frx":A4B9
         TabIndex        =   101
         Top             =   360
         Width           =   2055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmRetrabalho.frx":A525
         TabIndex        =   102
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   120
         TabIndex        =   103
         Top             =   4680
         Width           =   18015
         _ExtentX        =   31776
         _ExtentY        =   4471
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
         Left            =   9120
         Top             =   4080
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
         Height          =   435
         Left            =   9960
         Top             =   7440
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   767
         Image           =   "frmRetrabalho.frx":A58D
         Props           =   5
      End
   End
End
Attribute VB_Name = "frmRetrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'Acima - usado para poder editar o listview --------------------
'***********************************************************************************
'***********************************************************************************
'***********************************************************************************
'***********************************************************************************
'***********************************************************************************


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
Private rsRetrab As New ADODB.Recordset
Private SqlRetrab As String
Private vTime As String

Private Sub cmdCadastro_Click(Index As Integer)
On Error GoTo Err
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    Select Case Index
    Case 0
        'Chama CC - Centro de Custo
        Label8 = "-"
        ChamaGrid "" & vBancoTotvs & ".dbo.GCCUSTO", "codreduzido", txtformula(0), frmMPCompleto, "codreduzido", "nome"
        CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtformula(0), txtformula(1), 7, 2, txtformula(0), "S", txtformula(1), "1"
        montaEstrutTreeview
        
        LimpaVariaveis
        'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtformula(7), txtformula(8), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
        'compoeAutomatico
        compoeDadosLVs
    Case 1
        'Primeiro verifica se há um codigo de CD de retrabalho válido
        If txtformula(2).Text = "" Then
            mobjMsg.Abrir "Não foi informado um código válido de CD para retrabalho", Ok, critico
            txtformula(2).SetFocus
            Exit Sub
        End If
        
        'Depois gera o ID da Programação para lançar no LV
        Dim CodID As String
        Dim rsGeraID As New ADODB.Recordset
        Dim sqlGeraID As String
        
        sqlGeraID = "Select * from tbMP where tbMP.idprogramacao= " & Val(Me.txtformula(11))
        rsGeraID.Open sqlGeraID, cnBanco, adOpenKeyset, adLockOptimistic
        CodID = 0
        If txtformula(11).Text = "" Then 'Código do Cliente/Comitente
            rsGeraID.AddNew
            CodID = Format(GeraCodigoTB("tbMP", "idprogramacao", "", ""), "000000")
            rsGeraID.Fields(0) = CodID
            txtformula(11) = CodID
            'rsGeraID.Close
            Set rsGeraID = Nothing
        Else
            CodID = txtformula(11)
            txtformula(11) = CodID
        End If
        
        'Gera o ID do Retrabalho Somente quando insere algum item no LV
        Dim CodIDRET As String
        Dim rsGeraIDRET As New ADODB.Recordset
        Dim sqlGeraIDRET As String
        
        sqlGeraIDRET = "Select * from tbRetrabalho where tbRetrabalho.idretrabalho = " & Val(Me.txtformula(6))
        rsGeraIDRET.Open sqlGeraIDRET, cnBanco, adOpenKeyset, adLockOptimistic
        CodIDRET = 0
        If txtformula(6).Text = "" Then
            rsGeraIDRET.AddNew
            CodIDRET = Format(GeraCodigoTB("tbRetrabalho", "idRetrabalho", "", ""), "000000")
            rsGeraIDRET.Fields(0) = CodIDRET
            txtformula(6) = CodIDRET
            'rsGeraID.Close
            Set rsGeraIDRET = Nothing
        Else
            CodIDRET = txtformula(6)
            txtformula(6) = CodIDRET
        End If
        
        'Cria TextBox em tempo de Execução
        txtLV = Val(txtformula(11).Text) & Val(vPonte1) & Val(txtformula(15)) & Val(Combo1.Text)
        If txtformula(26) = "" Then txtformula(26) = "-"
        If vPonte1 = "" Then vPonte1 = Format(txtformula(17).Text, "000000000") & "/" & txtformula(18).Text
        vPonte2.Text = DTPicker2.Value
        vPonte3.Text = Label8.Caption
        vPonte4.Text = Combo1.Text
        vPonte5.Text = vPonte1 & Format(Combo1.Text, "000")
        If ValidaCampos(ListView1, txtformula(0), vPonte3, txtformula(5), vPonte4, txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0)) = False Then Exit Sub
        
        
        'Text1.Text recebe o valor abaixo quando não há nenhum desenho/posição/item para a operação
        If Text1.Text = "" Then Text1.Text = "009999"
        If IncluirLV(ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtformula(5), vPonte3, txtformula(11), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), vPonte5) = False Then
            Exit Sub
        End If
        
        
        
        'IncluirLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), vPonte5
        
        
        ListView1.Sorted = True
        ListView1.SortKey = 14
        ListView1.SortOrder = lvwAscending
        
        'Salva dados do LV2 na tabela tbMPItensRet
        Salvar_Dados_Ret
        
        vPonte1.Text = "0"
        txtformula(27) = "0"
        LimpaVariaveis
        LimpaControles txtformula(0), txtformula(1), vPonte4, txtformula(26), txtformula(5), vPonte3, txtformula(1), txtformula(1), txtformula(1), txtformula(1)
        LimpaControles txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1), txtformula(1)
        Combo1.Text = ""
        txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        SomaLV ListView1, 6, Text2
        TreeView3.Nodes.Clear
        aicAlphaImage1.ClearImage
        
        'SALVA OS DADOS A CADA VEZ QUE UM ITEM É INCLUIDO NO LISTVIEW
        'DEIXA O SISTEMA BEM MAIS LENTO
        
        salvar_Dados
    Case 2
        ListView2.ListItems.Clear
        EditaLVMP
        compoeDadosLV2
        lebel3 = ""
        SomaLV ListView2, 1, vPonte5
        If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 3
        ExcluirItemLV ListView1
        'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtformula(7), txtformula(8), txtformula(26), txtformula(2), txtformula(2), txtformula(2)
        txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        SomaLV ListView1, 6, Text2
    Case 4
        ListView2.ListItems.Clear
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        Label3 = ""
        SomaLV ListView2, 1, vPonte5
        If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 5
        ChamaGridFCE
        CarregaFCE
    Case 6
        If txtformula(12).Text <> "" Then
            ChamaGridProjeto
            CarregaProjeto
            mostraDesenhos "tbitemlm", TreeView1
            'txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
        End If
    Case 7
        sqlDeletar = "Delete from tbMPDesSel" & vTime
        rsDeletar.Open sqlDeletar, cnBanco
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        vAcumulaTempo = 0
        buscaChecado2 TreeView1
        
        insereDadosListview
        compoeText1
        'mostraDesenhos "tbMPDesSel" & vTime, TreeView2
        
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
        Label3 = ""
        SomaLV ListView2, 1, vPonte5
        If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    Case 8
        vPesoTotal2 = 0
        Text1.Text = ""
        vAcumula = ""
        excluiChecados
        compoeText1
        Label3 = ""
        SomaLV ListView2, 1, vPonte5
        If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
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
        txtformula(18).Text = Val(txtformula(18)) '+ 1
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
    Case 18
        'ExcluirItemLV ListView4
        'SomaLV ListView4, 2, Text7
        Timer1.Enabled = True
    End Select
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
    'Grava dados do formulário na tabela tbMP
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    'Se algum campo não for preciso gravar ou alterar os dados (identifique-o, mas nos dois parâmetros deixe apenas as aspas sem nada)
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
        
    
    'Limpa dados da Matriz vQualquerDado
    limpaQualquerDado
    'Grava dados do formulário na tabela tbRetrabalho
    'O 1º parametro é o valor que sera gravado no campo
    'O 2º parametro é o tipo de dado que o campo armazena
    vQualquerDado(1, 1) = txtformula(6).Text
    vQualquerDado(1, 2) = "I"
    vQualquerDado(2, 1) = txtformula(11).Text
    vQualquerDado(2, 2) = "I"
    vQualquerDado(3, 1) = txtformula(2).Text
    vQualquerDado(3, 2) = "I"
    GravaDados "tbRetrabalho", "idretrabalho", "I", txtformula(6), 3, "", "", txtformula(6)
        
    'Grava dados ListView1
    limpaQualquerDado
    ordenaLVArray ListView1, "8", "0", "2", "3", "4", "5", "6", "7", "9", "1", "10", "11", "12", "13", "", ""
    GravaDadosLV "tbMPItens", "idprogramacao", "I", txtformula(11)
    complementaDadosLV
        
        
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
        sqlOSstatus = "SELECT IDOS,RASTREABILIDADE,OBSERVACAO,DATAOS,REVISAO,STATUS FROM TBOS as a where a.idos = '" & Val(txtformula(17).Text) & "' and revisao = '" & Val(txtformula(18).Text) & "'"
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
        vQualquerDado(5, 1) = txtformula(18).Text
        vQualquerDado(5, 2) = "S"
        vQualquerDado(6, 1) = statusOS
        vQualquerDado(6, 2) = "I"
        GravaDados "tbOS", "idos", "I", txtformula(17), 6, "revisao", "S", txtformula(18)
        
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
    End If
End Function

Private Sub complementaDadosLV()
    Dim rsUpdate As New ADODB.Recordset
    Dim SqlUpdate As String
    ListView1.ListItems.Item(1).Selected = True
    SqlUpdate = "Update tbMPItens set idos = '" & Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9)) & "',revisaoos = '" & Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 11, 3)) & "' where idprogramacao = '" & Val(ListView1.SelectedItem.ListSubItems.Item(8)) & "'"
    rsUpdate.Open SqlUpdate, cnBanco
End Sub

Private Sub Salvar_Dados_Ret()
On Error GoTo Err
    'Grava dados do ListView2 na tabela tbMPItensRet
    'limpaQualquerDado
    'ordenaLVArray ListView2, "7", "8", "0", "10", "11", "", "", "", "", "", "", "", "", "", "", ""
    'GravaDadosLV "tbMPItensRet", "idprogramacao", "I", txtformula(11)
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim X As Integer
    sqlDeletar = "delete from tbMPItensRet where idprogramacao = '" & Val(txtformula(11).Text) & "' and idoperacao = '" & Combo1.Text & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView2.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbMPItensRet"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView2.ListItems.Count
            ListView2.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = Val(ListView2.SelectedItem.ListSubItems.Item(7))
            rsSalvar.Fields(1) = Val(ListView2.SelectedItem.ListSubItems.Item(8))
            rsSalvar.Fields(2) = Val(ListView2.ListItems.Item(X))
            If ListView2.SelectedItem.ListSubItems.Item(10) <> "" Then
                rsSalvar.Fields(3) = Val(ListView2.SelectedItem.ListSubItems.Item(10))
            Else
                rsSalvar.Fields(3) = Val(txtformula(11).Text)
            End If
            rsSalvar.Fields(4) = Combo1.Text 'Val(ListView2.SelectedItem.ListSubItems.Item(11))
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
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
        'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtDecoder, txtResultado, txtformula(7), txtformula(8), txtformula(26), txtformula(2)
        AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtformula(5), vPonte3, txtformula(26), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), txtLV
        
'       AlteraLV ListView1, txtformula(15), vPonte1, txtformula(0), txtformula(1), Text1, vPonte2, txtResultado, vPonte3, txtformula(11), txtformula(5), txtformula(26), vPonte4, txtLV, txtformula(27), txtLV
        
        montaEstrutTreeview
        compoeDadosLVs
        'If txtformula(11) = Val(varGlobal) Then txtformula(11) = ""
        txtformula(17).Text = Mid$(vPonte1.Text, 1, 9)
        If vPonte2.Text <> "" Then
            DTPicker2.Value = vPonte2.Text
        End If
        Label8.Caption = vPonte3.Text
        Combo1.Text = vPonte4.Text
        EditaTreeview
        CompoeControles
        separaDadosText1 Text1
        vPesoTotal2 = 0
        Text1 = ""
        
        'Calcula o tempo automaticamente se vPAutomatico for igual a "S"
        'Caso contrário o procedimento requer entrada manual de parâmetros
        If vPAutomatico = "S" Then
            vAcumula = ""
            vAcumulaTempo = 0
            
            insereDadosListview
            'mostraDesenhos "tbMPDesSel" & vTime, TreeView2
            txtResultado = Format(vAcumulaTempo, "#,##0.00;(#,##0.00)")
            'txtformula(5).Text = ""
            txtformula(5).Text = "AUTOMÁTICO"
        Else
            insereDadosListview
            'mostraDesenhos "tbMPDesSel" & vTime, TreeView2
        End If
        verOS 'Verifica se há alguma OS ativa
        compoeControlesOS 2
        If vPesoTotal2 <> 0 Then Label3 = Format(vPesoTotal2, "#,##0.00;(#,##0.00)") Else Label3 = "-"
'---------------------
        LimpaLV ListView3
        chamaSQL "select a.idservterc,b.nmserv,a.observacao,a.quantidade,a.idos from tbServTercOS as a inner join tbServTerc as b on a.idservterc = b.idservterc where a.idos = '" & Val(vPonte1.Text) & "'"
        Compoe_Listview ListView3, Sqlp, "00"
'---------------------
'---------------------
        'LimpaLV ListView4
        'chamaSQL "select a.* from tbMPHist as a where a.programacao = '" & Val(txtformula(11)) & "' and a.seqprog = '" & Val(txtformula(15)) & "'"
        'Compoe_Listview ListView4, Sqlp, "00"
        'SomaLV ListView4, 2, Text7
'---------------------
        compoeText1
        txtformula(11) = Format(txtformula(11), "000000")
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
        GeraOSLV = LV.SelectedItem.ListSubItems.Item(1) + 1
        LV.SortKey = 0
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
        SqlTransf = "Insert into tbOSItens(idos,revisao,fce,projeto,codlm,codseq,idcc,idprogramacao,status,codigobarra,idoperacao) Values('" & Val(txtformula(17)) & "','" & txtformula(18) & "','" & Val(txtformula(12)) & "','" & txtformula(13) & "','" & Val(vCodLM) & "','" & Val(vCodSeq) & "','" & ListView1.SelectedItem.ListSubItems.Item(2) & "','" & Val(txtformula(11)) & "',1,'" & ListView1.SelectedItem.ListSubItems.Item(12) & "','" & ListView1.SelectedItem.ListSubItems.Item(11) & "')"
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

Private Sub CarregaFCE()
On Error GoTo Err
    Dim X As Integer
'    sqlFCE = "Select * from tbprojetos where fce = '" & txtformula(12) & "' order by fce"
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
'    Sqlp = "Select fce,MAX(oc) from tbprojetos group by FCE order by fce"
    Sqlp = "Select a.fce,MAX(a.oc) from tbprojetos as a inner join tbFCE as b on a.fce = b.fce where b.status <> 1 group by a.FCE order by a.fce"
    procnom = "fce"
    campo = 0
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de FCE"
    Pesquisa = frmMPCompleto.Tag
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
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
    Pesquisa = frmRetrabalho.Tag
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

Private Sub Command3_Click()
    ChamaGridCD
    chamaCD
End Sub

Private Sub Form_Activate()
    'vTime = Time
    'vTime = RemoveMask(vTime)
    excluiTabela
    criaTabela
'    verOS 'Verifica se há alguma OS ativa
'    listview_cabecalho
'    Status = Pesquisa
'    If Status = "novo" Then
'        txtformula(11) = ""
'    ElseIf Status = "editar" Then
'        ResultPesq
'    End If
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
    DTPicker2 = Date
    SSTab1.Tab = 0
    listview_cabecalho
    
    Status = Pesquisa
    ResultPesq
    If Status <> "novo" And Status <> "" Then
        txtformula(11) = varGlobal
    End If
    
'    compoeControlesOS 1
    txtformula(18) = Val(txtformula(18)) + 1
    mostraDesenhos "tbitemlm", TreeView1
    compoeDadosLV
    
    compoeControlesOS 1
    
    
    SomaLV ListView1, 6, Text2

    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_Unload(Cancel As Integer)
    excluiTabela
End Sub

Private Sub ListView1_Click()
    ListView2.ListItems.Clear
    EditaLVMP
    compoeDadosLV2
    Label3 = ""
    SomaLV ListView2, 1, vPonte5
    calculaTempoApropriado ListView1.SelectedItem.ListSubItems.Item(12), ListView1.SelectedItem.ListSubItems.Item(6)
    If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
End Sub

Private Sub ListView1_DblClick()
    ListView2.ListItems.Clear
    EditaLVMP
    compoeDadosLV2
    Label3 = ""
    SomaLV ListView2, 1, vPonte5
    If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    ListView2.ListItems.Clear
    EditaLVMP
    compoeDadosLV2
    Label3 = ""
    SomaLV ListView2, 1, vPonte5
    If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    ListView2.ListItems.Clear
    EditaLVMP
    compoeDadosLV2
    Label3 = ""
    SomaLV ListView2, 1, vPonte5
    If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
End Sub

Private Sub ListView3_DblClick()
    AlteraLV ListView3, txtformula(21), txtformula(22), txtformula(23), txtformula(24), txtformula(17), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21), txtformula(21)
End Sub

Private Sub ListView4_Click()
    EditaLVHist
    CompoeHist
End Sub

Private Sub Timer1_Timer()
    txtResultado = Text7
    Timer1.Enabled = False
End Sub

Private Function chamaCD()
On Error GoTo Err
    chamaChapa = False
    Dim rsCD As New ADODB.Recordset
    Dim SqlCD As String
    SqlCD = "select a.idcd,a.observacao,b.idrnc from tbComunicacaoDesvio as a inner join tbRNC as b on a.idcd = b.idcd where a.idcd = '" & Val(txtformula(2).Text) & "' and b.gerouretrabalho = 'S' and a.idos = '" & Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9)) & "'"
    rsCD.Open SqlCD, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCD.EOF Then
        txtformula(2).Text = Format(rsCD.Fields(0), "000000")
        txtformula(3).Text = rsCD.Fields(1)  'Observação
    Else
        mobjMsg.Abrir "CD de retrabalho não identificada no sistema", Ok, critico, "Atenção"
        txtformula(2).Text = ""
        txtformula(3).Text = ""
    End If
    rsCD.Close
    Set rsCD = Nothing
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

Private Sub ChamaGridCD()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "select a.idcd,a.observacao,b.idrnc from tbComunicacaoDesvio as a inner join tbRNC as b on a.idcd = b.idcd where b.gerouretrabalho = 'S' and a.idos = '" & Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(1), 1, 9)) & "'"
    procnom = "chamaCD"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa Comunicação de Desvio"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "idcd=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtformula(2).Text = Pesquisa
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
        rsLocal.Close
        Set rsLocal = Nothing
        txtformula(2).Text = ""
        txtformula(3).Text = ""
    End If
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

Private Sub excluiChecados()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        If X > Y Then Exit For
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Remove (X)
            Y = ListView2.ListItems.Count
            X = 1
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

Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Seq.", ListView1.Width / 26
    ListView1.ColumnHeaders.Add , , "OS nº/Rev.", ListView1.Width / 13
    ListView1.ColumnHeaders.Add , , "ID. C.Custo", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Nome C. Custo", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Desenhos/Itens", ListView1.Width / 4.5
    ListView1.ColumnHeaders.Add , , "Data Prevista", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "T. Calculado", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Grupo", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "ID Programação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Variáveis", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Observação", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Operação", ListView1.Width / 16
    ListView1.ColumnHeaders.Add , , "Código de Barras", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Ordenação", ListView1.Width / 10000
    Me.ListView1.ColumnHeaders(7).Alignment = lvwColumnRight
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Qt Ret.", ListView2.Width / 13
    ListView2.ColumnHeaders.Add , , "PT Ret.", ListView2.Width / 13 'Peso Total LM
    ListView2.ColumnHeaders.Add , , "Desenho/Posição/item", ListView2.Width / 3.8
    ListView2.ColumnHeaders.Add , , "Material", ListView2.Width / 3.2
    ListView2.ColumnHeaders.Add , , "Qt LM", ListView2.Width / 13
    ListView2.ColumnHeaders.Add , , "PU LM", ListView2.Width / 13 'Peso Unitario LM
    ListView2.ColumnHeaders.Add , , "PT LM", ListView2.Width / 13 'Peso Total LM
    ListView2.ColumnHeaders.Add , , "codLM", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "codSEQ", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "cod Material", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "idProgramacao", ListView2.Width / 10000
    ListView2.ColumnHeaders.Add , , "idOperacao", ListView2.Width / 10000
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "ID", ListView3.Width / 6
    ListView3.ColumnHeaders.Add , , "Serviço", ListView3.Width / 2.5
    ListView3.ColumnHeaders.Add , , "Descrição", ListView3.Width / 3.5
    ListView3.ColumnHeaders.Add , , "Qtd.", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "OS", ListView3.Width / 10000
    
    'ListView4.ColumnHeaders.Clear
    'ListView4.ColumnHeaders.Add , , "ID", ListView4.Width / 6
    'ListView4.ColumnHeaders.Add , , "Variáveis", ListView4.Width / 2.5
    'ListView4.ColumnHeaders.Add , , "Tempo", ListView4.Width / 2.5
    'ListView4.ColumnHeaders.Add , , "Grupo", ListView4.Width / 10000
    'ListView4.ColumnHeaders.Add , , "Programação", ListView4.Width / 10000
    'ListView4.ColumnHeaders.Add , , "Seq", ListView4.Width / 10000
    'ListView4.ColumnHeaders.Add , , "codreduzido", ListView4.Width / 10000
    'ListView4.ColumnHeaders.Add , , "idFormula", ListView4.Width / 10000
    'Me.ListView4.ColumnHeaders(3).Alignment = lvwColumnRight
    
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    'ListView4.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ResultPesq()
On Error GoTo Err
    SqlProg = "Select a.idprogramacao,a.dataprogramacao,a.codprojeto,a.responsavel,a.ativo,b.fce,b.projeto from tbMP as a inner join tbProjetos as b on a.codprojeto = b.codprojeto where a.idprogramacao = '" & Val(varGlobal) & "'"
    rsProg.Open SqlProg, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlRetrab = "Select a.idretrabalho, a.idcd from tbRetrabalho as a where a.idprogramacao = '" & Val(varGlobal) & "'"
    rsRetrab.Open SqlRetrab, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsProg.RecordCount > 0 Then
        compoeControlesForm
        'mostraDesenhos "tbitemlm", TreeView1
        'compoeDadosLV
        'SomaLV ListView1, 6, Text2
    End If
    rsProg.Close
    Set rsProg = Nothing
    rsRetrab.Close
    Set rsRetrab = Nothing
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
    'txtformula(11).Text = Format(rsProg.Fields(0), "000000")
    DTPicker1.Value = rsProg.Fields(1) 'Data da Programação
    txtformula(16).Text = rsProg.Fields(2) 'ID do Projeto
    txtformula(14).Text = rsProg.Fields(3) 'Responsável
    txtformula(12).Text = rsProg.Fields(5) 'FCE
    txtformula(13).Text = rsProg.Fields(6) 'Projeto
    If rsRetrab.RecordCount > 0 Then
        txtformula(6).Text = Format(rsRetrab.Fields(0), "000000000") 'ID Retrabalho
        txtformula(2).Text = Format(rsRetrab.Fields(1), "000000000") 'ID CD
    End If

End Sub

Private Sub insereDadosListview()
On Error GoTo Err
    'EM DESENVOLVIMENTO
    
    Dim rsListView2 As New ADODB.Recordset
    Dim SqlListView2 As String

'    SqlListView2 = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(h.idos) as OS " & _
'    "from tbitemlm as a inner join CORPORERM.dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
'    "left join CORPORERM.dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel" & RemoveMask(vTime) & " as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq inner join tbProjetos as g on g.codprojeto = c.codprojeto left join tbositens as h on a.fce = h.fce and a.codlm = h.codlm and a.codseq = h.codseq Where a.fce = '" & Val(txtformula(12)) & "' and g.projeto = '" & txtformula(13) & "'" & _
'    "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    
    SqlListView2 = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(h.idos) as OS " & _
    "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
    "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel" & RemoveMask(vTime) & " as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq inner join tbProjetos as g on g.codprojeto = c.codprojeto left join tbositens as h on a.fce = h.fce and a.codlm = h.codlm and a.codseq = h.codseq Where a.fce = '" & Val(txtformula(12)) & "'" & _
    "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    
    
    rsListView2.Open SqlListView2, cnBanco, adOpenKeyset, adLockReadOnly
    If rsListView2.RecordCount = 0 Then Exit Sub
    
    'ListView2.ListItems.Clear
    
    'O LAÇO ABAIXO PEGA DADOS DA TABELA TEMPORÁRIA E ENVIA PARA O LISTVIEW2
    Do While Not rsListView2.EOF
        Set ItemLst = ListView2.ListItems.Add(, , "-")   'Quantidade Retrabalho
        ItemLst.SubItems(1) = "-" 'Peso Total Retrabalho
        ItemLst.SubItems(2) = LTrim(rsListView2.Fields(0) & "/" & rsListView2.Fields(10) & " - " & rsListView2.Fields(4) & "/" & rsListView2.Fields(5))   'Desenho/Posição/Item
        ItemLst.SubItems(3) = "" & LTrim(rsListView2.Fields(3)) 'Nome Material
        ItemLst.SubItems(4) = "" & rsListView2.Fields(7) 'Quantidade LM
        ItemLst.SubItems(5) = "" & rsListView2.Fields(9) 'Peso Unitário LM
        ItemLst.SubItems(6) = "" & rsListView2.Fields(7) * rsListView2.Fields(9) 'Quantidade LM * Peso Unitário LM = Peso Total LM
        ItemLst.SubItems(7) = "" & rsListView2.Fields(11) 'Codigo LM
        ItemLst.SubItems(8) = "" & rsListView2.Fields(12) 'Codigo Sequencial da LM
        ItemLst.SubItems(9) = "" & rsListView2.Fields(13) 'Codigo do Material
        ItemLst.SubItems(10) = "" & txtformula(11) 'Identificador da Programacao
        ItemLst.SubItems(11) = "" & Combo1.Text 'Identificador da Operação
        'ItemLst.ListSubItems(3).Bold = True
        rsListView2.MoveNext
        X = X + 1
    Loop
    rsListView2.Close
    Set rsListView2 = Nothing
    Me.ListView2.ColumnHeaders(2).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(7).Alignment = lvwColumnRight
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
       
    TV.Nodes.Clear

    If vTabela = "tbitemlm" Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(g.idos) as OS " & _
        "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbProjetos as f on f.codprojeto = c.codprojeto left join tbositens as g on a.fce = g.fce and a.codlm = g.codlm and a.codseq = g.codseq Where a.fce = '" & Val(txtformula(12)) & "' and f.projeto = '" & txtformula(13) & "'" & _
        "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    ElseIf vTabela = "tbMPDesSel" & vTime Then
        SqlTreeview = "select c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat,b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq,MAX(h.idos) as OS " & _
        "from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
        "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT inner join tbMPDesSel" & RemoveMask(vTime) & " as f on a.fce = f.fce and a.codlm = f.codlm and a.codseq = f.codseq inner join tbProjetos as g on g.codprojeto = c.codprojeto left join tbositens as h on a.fce = h.fce and a.codlm = h.codlm and a.codseq = h.codseq Where a.fce = '" & Val(txtformula(12)) & "' and g.projeto = '" & txtformula(13) & "'" & _
        "Group by c.desenho,c.revisao,b.CODIGOPRD + ' - ' + cast(a.codmat as varchar),b.NOMEFANTASIA,d.posicao,d.item,a.quantcj,a.quantunit,a.dimensoes,a.pesounit,d.descposicao,a.codlm,a.codseq Order by c.desenho,d.posicao,b.NOMEFANTASIA"
    End If
    
    rsTreeview.Open SqlTreeview, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTreeview.RecordCount = 0 Then Exit Sub
    
    Text10.Text = rsTreeview.Fields(0)
    
    vJuntaNome = ""
              vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & " - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
    separaDadosTree vJuntaNome
    vNome1 = vNomeA
    vNome2 = vNomeB
    vNome3 = vNomeC
    vNo = 0
    
    'ABAIXO - INSERE DADOS NO TREEVIEW
    On Error Resume Next
    Do While Not rsTreeview.EOF
        'PRIMEIRO NO
        Set nd = TV.Nodes.Add(, , vNome1, vNome1)
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
                    'TESTE DE COR --------------
                    If Not IsNull(rsTreeview.Fields(13)) Then
                        nd.ForeColor = &H8000&
                    End If
                    '----------------------------
                    If Not rsTreeview.EOF Then rsTreeview.MoveNext Else Exit Do
                    vJuntaNome = rsTreeview.Fields(0) & " (" & rsTreeview.Fields(1) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(10) & " - " & rsTreeview.Fields(4) & " - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ");" & rsTreeview.Fields(5) & " - " & rsTreeview.Fields(3) & " (" & rsTreeview.Fields(8) & ") - (OS: " & Format(rsTreeview.Fields(13), "000000000") & ") - ID: " & Format(rsTreeview.Fields(11), "00") & Format(rsTreeview.Fields(12), "000")
                    separaDadosTree vJuntaNome
                    vPula = 1
                Loop
                
                'Utiliza o peso calculado da posição para classificar o tipo de estrutura e calcular o tempo
                'em seguida acumula o tempo encontrado para determinar o tempo real de fabricação
                'If vPAutomatico = "S" Then
                '    txtformula(5) = Format(vPesoPosicao, "#,##0.00;(#,##0.00)")
                '    txtformula_KeyDown 5, 13, 5
                '    vAcumulaTempo = vAcumulaTempo + txtResultado
                'End If
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

Private Sub compoeText1()
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        If Text1.Text = "" Then
            Text1 = Format(ListView2.SelectedItem.ListSubItems.Item(7), "00") & Format(ListView2.SelectedItem.ListSubItems.Item(8), "000")
        Else
            Text1 = Text1.Text & ";" & Format(ListView2.SelectedItem.ListSubItems.Item(7), "00") & Format(ListView2.SelectedItem.ListSubItems.Item(8), "000")
        End If
    Next
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    'SqlCompoe = "Select a.parametros,a.formula from tbFormula as a where a.idprd = '" & txtformula(0) & "' and idform = '" & Val(txtformula(4)) & "'"
    sqlCompoe = "Select a.parametros,a.formula,a.observacao,a.imagem from tbFormula as a where a.codreduzido = '" & txtformula(0) & "' and a.idform = '" & Val(txtformula(4)) & "'"
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
'        txtformula(2).Text = rsCompoe.Fields(0) 'Parâmetros
'        txtformula(3).Text = rsCompoe.Fields(1) 'Formula
'       If Not IsNull(rsCompoe.Fields(2)) Then txtformula(6).Text = rsCompoe.Fields(2) 'Observação
        If Not IsNull(rsCompoe.Fields(3)) Then Label53 = rsCompoe.Fields(3) Else Label53 = "-" 'Imagem
    Else
'        txtformula(2).Text = "" 'Parâmetros
'        txtformula(3).Text = "" 'Formula
'        txtformula(6).Text = "" 'Observação
        Label53 = "-" 'Imagem
    End If
'    If Mid$(txtformula(2).Text, 1, 7) = "formula" Then
'        localizaFormula Mid$(txtformula(2).Text, 9, 1), 1
'    End If
'    If Mid$(txtformula(2).Text, 12, 7) = "formula" Then
'        localizaFormula Mid$(txtformula(2).Text, 20, 1), 2
'    End If
    
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

Private Sub compoeDadosLVs()
    'Faz referências a Funções que estão no: Module1.bas
    'Listview2 - Constantes
    'LimpaLV ListView2
    'chamaSQL "Select a.idseq,a.valconst,a.descricao from tbconstantesCC as a where a.idprd = '" & txtformula(0) & "'"
    'Compoe_Listview ListView2, Sqlp, "000"
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
    AlteraTreeview
    'LimpaVariaveis
    'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
    'compoeDadosLVs
    CompoeControles
    compoeAutomatico
    If vPAutomatico = "S" Then
        txtformula(5).Enabled = False
        If Text1.Text <> "" Then
            vAcumula = ""
            vAcumulaTempo = 0
            Text1 = ""
            'mostraDesenhos "tbMPDesSel" & vTime, TreeView2
            txtResultado = Format(vAcumulaTempo, "#,##0.00;(#,##0.00)")
        Else
            Msgbox "Nenhum DESENHO selecionado na guia de Desenhos"
        End If
        txtformula(5).Text = "AUTOMÁTICO"
    Else
        txtformula(5).Enabled = True
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
            'LimpaControles txtformula(2), txtformula(3), txtformula(5), txtformula(6), txtResultado, txtformula(7), txtformula(8), txtformula(2), txtformula(2)
            compoeDadosLVs
        End If
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If chamaCD = False Then Exit Sub
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'preparaDados
            txtResultado = ""
            calculaValores 1
            
            'IncluiHistorico
            'SomaLV ListView4, 2, Text7
            'Timer1.Enabled = True
            
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
    'vPonte3.Text = Label8.Caption
    'If txtformula(28) = "" Then
    '    txtformula(28) = Format(GeraCodigoLV(ListView4), "00")
    'End If
    'IncluirLV ListView4, txtformula(28), txtformula(5), txtResultado, vPonte3, txtformula(11), txtformula(15), txtformula(0), txtformula(4), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28), txtformula(28)
    'txtformula(28) = ""
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
On Error GoTo Err
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
    'O ScriptControl é um componente. Ele interpreta e executa a formula/expressão numérica de um textbox
    If vQual = 1 Then
        txtResultado = Format(ScriptControl1.Eval(txtDecoder), "#,##0.00;(#,##0.00)")
        'SOMENTE PARA O CENTRO DE CUSTO SOLDA
        'CONVERTE O RESULTADO EM HORAS PARA MINUTOS
        
'        If Mid$(txtformula(0).Text, 1, 12) = "3000.3104.SC" Then
'            txtResultado = Format(txtResultado * 60, "#,##0.00;(#,##0.00)")
'        End If
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
'    LimpaVariaveis
'    If txtformula(5) = "" Then
'        Msgbox "Favor informar o campo: " & txtformula(5).Tag, vbInformation, "Atenção"
'        txtformula(5).SetFocus
'        Exit Sub
'    End If
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
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
    Else
        vTxtRetorno.Text = "-"
    End If
End Sub

Private Sub txtformula_LostFocus(Index As Integer)
    voltaCorText txtformula(Index)
    Select Case Index
    Case 2
        If chamaCD = False Then Exit Sub
    Case 5
        'preparaDados
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

Private Sub compoeControlesOS(vOrdem As Integer)
On Error GoTo Err
'vOrdem serve para saber se é ou não a primeira vez que se executa a função
    Dim rsCompoeOS As New ADODB.Recordset
    Dim SqlCompoeOS As String
    If Pesquisa <> "novo" Then
        SqlCompoeOS = "select top 1 * from tbos where idos = '" & Val(Mid(vPonte1, 1, 9)) & "' and revisao = '" & Val(Mid(vPonte1, 11, 2)) & "' order by idos Desc"
    Else
        'SqlCompoeOS = "select top 1 * from tbos where idos = '" & Val(Mid(vPonte1, 1, 9)) & "' order by idos,revisao Desc"
        'SqlCompoeOS = "select max(cast(a.revisao as int)) as revisao from tbos as a where a.idos = '" & Val(Mid(vPonte1, 1, 9)) & "'"
        
        SqlCompoeOS = "Declare @revisao as int SET @revisao = 0 " & _
                      "SELECT @revisao = max(cast(a.revisao as int)) from tbos as a where idos = '" & Val(Mid(vPonte1, 1, 9)) & "'" & _
                      "select a.idos,a.rastreabilidade,a.observacao,a.dataos,@revisao as revisao from tbos as a where idos = '" & Val(Mid(vPonte1, 1, 9)) & "' and revisao = @revisao"
    End If
    rsCompoeOS.Open SqlCompoeOS, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoeOS.EOF Then
        txtformula(17).Text = rsCompoeOS.Fields(0) 'Id OS
        If vOrdem = 1 Then
            If Pesquisa = "novo" Then
                txtformula(18).Text = rsCompoeOS.Fields(4) + 1 'Revisão
            Else
                txtformula(18).Text = rsCompoeOS.Fields(4) 'Revisão
            End If
        End If
        txtformula(19).Text = rsCompoeOS.Fields(1) 'Rastreabilidade
        txtformula(20).Text = rsCompoeOS.Fields(2) 'Observação
        DTPicker3.Value = rsCompoeOS.Fields(3)   'Data da OS
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

Private Sub compoeDadosLV2()
On Error GoTo Err
    Dim rsCompoeLV2 As New ADODB.Recordset
    Dim SqlCompoeLV2 As String
    Dim Y As Integer
    Y = ListView2.ListItems.Count
    SqlCompoeLV2 = "Select * from tbMPItensRet as a where a.idprogramacao = '" & Val(txtformula(11)) & "' and a.idoperacao = '" & Val(Combo1.Text) & "' order by a.codlm,a.codseq"
    rsCompoeLV2.Open SqlCompoeLV2, cnBanco, adOpenKeyset, adLockReadOnly
    While Not rsCompoeLV2.EOF
        For X = 1 To Y
            ListView2.ListItems(X).Selected = True
            If Val(ListView2.SelectedItem.ListSubItems.Item(7)) = rsCompoeLV2.Fields(0) And Val(ListView2.SelectedItem.ListSubItems.Item(8)) = rsCompoeLV2.Fields(1) Then
                ListView2.ListItems.Item(X) = rsCompoeLV2.Fields(2)
                ListView2.SelectedItem.ListSubItems.Item(1) = rsCompoeLV2.Fields(2) * ListView2.SelectedItem.ListSubItems.Item(5)
            End If
        Next
        rsCompoeLV2.MoveNext
    Wend
    rsCompoeLV2.Close
    Set rsCompoeLV2 = Nothing
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
On Error GoTo Err
    'EM TESTE
    Dim rsCompoeOS As New ADODB.Recordset
    Dim SqlCompoeOS As String
    SqlCompoeOS = "Declare @revisao as int, @OS as int SET @revisao = 0 SET @OS = 0 " & _
                  "SELECT @OS = MAX(a.idos) from tbMPItens as a where a.idprogramacao = '" & Val(varGlobal) & "' " & _
                  "SELECT @revisao = max(cast(a.revisao as int)) from tbos as a where idos = @OS " & _
                  "select a.idprogramacao,b.idos,@revisao as revisao from tbMPItens as a inner join tbos as b on a.idos = b.idos " & _
                  "where a.idprogramacao = 282 group by a.idprogramacao,b.idos"
    rsCompoeOS.Open SqlCompoeOS, cnBanco, adOpenKeyset, adLockReadOnly
    If Pesquisa = "novo" Then
        txtformula(18).Text = rsCompoeOS.Fields(2) + 1 'Revisão
    End If
    rsCompoeOS.Close
    Set rsCompoeOS = Nothing
    'EM TESTE
    
    
    LimpaLV ListView1
    If Pesquisa = "novo" Then
        chamaSQL "select a.idsequencia,RIGHT('000000000'+ CONVERT(VARCHAR,a.idos),9) + '/' + '" & txtformula(18) & "',a.idcc,a.nomecc,a.desenhos,a.dataprevista,a.tempocalc,a.grupo,a.idprogramacao,a.variaveis,a.observacao,a.idoperacao,a.codigobarra,a.status,Replicate ('0',9 - Len(Cast(a.idos as varchar))) + Cast(a.idos as varchar) + Replicate ('0',3 - Len(Cast(a.idoperacao as varchar))) + Cast(a.idoperacao as varchar)  as ordenação  from tbMPItens as a where a.idprogramacao = '" & Val(varGlobal) & "'"
    Else
        chamaSQL "select a.idsequencia,RIGHT('000000000'+ CONVERT(VARCHAR,a.idos),9) + '/' + a.revisaoos,a.idcc,a.nomecc,a.desenhos,a.dataprevista,a.tempocalc,a.grupo,a.idprogramacao,a.variaveis,a.observacao,a.idoperacao,a.codigobarra,a.status,Replicate ('0',9 - Len(Cast(a.idos as varchar))) + Cast(a.idos as varchar) + Replicate ('0',3 - Len(Cast(a.idoperacao as varchar))) + Cast(a.idoperacao as varchar)  as ordenação  from tbMPItens as a where a.idprogramacao = '" & Val(varGlobal) & "'"
    End If
    Compoe_Listview ListView1, Sqlp, "000"
    txtformula(15) = Format(GeraCodigoLV(ListView1), "000")
    
    ListView1.Sorted = True
    ListView1.SortKey = 14
    ListView1.SortOrder = lvwAscending
    MudaCorLV1
    If vStatus > 1 Then
        'bloqueiaEdicao
    End If
    ListView1.ListItems(1).Selected = True
    vPonte1 = ListView1.SelectedItem.ListSubItems.Item(1)
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
        Msgbox Err.Number & " - " & Err.Description
        Resume
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
    'If SkinLabel12 = "-" Or SkinLabel12 = "" Then
    '    SkinLabel13 = "-"
    'Else
        SkinLabel26 = vHorasApropriadas
        SkinLabel27 = vTempoOrcadoConvertido
    'End If
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

'**********************************************
'**********************************************
'**********************************************
'**********************************************
'**********************************************

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView2.SelectedItem Is Nothing Then
        ListView2.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 1 To 1
            leftPos = ListView2.Left + ListView2.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView2.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView2.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView2.SelectedItem.Text
                    Else
                        .Text = ListView2.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    If txtLvw.Enabled = True Then .SetFocus 'Else txtModelo(1).SetFocus
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
   
    Set lvwCol = ListView2.ColumnHeaders(ListView2.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView2.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView2.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView2.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView2.Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView2.Left Then txtLeft = ListView2.Left + 48
    
        txtRightMax = ListView2.Left + ListView2.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView2.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView2.Left + ListView2.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView2.Top
        If Not ListView2.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView2.Top + ListView2.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView2.Top + ListView2.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
        With txtLvw '.move produces runtime error with -ve values
            .Left = ListView2.ColumnHeaders.Item(m_ColIndex).Left + 460 '+ 330
            .Top = txtTop '+ 165
            .Width = 600
            .Height = 285
        End With
    End If
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView2 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
        'other keys can be used for navigation
    End Select
    If txtLvw.Text = "-" Then txtLvw.Text = ""
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLvw_LostFocus()
On Error GoTo TrataErro
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "-"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView2.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView2.ListItems(ListView2.ListItems.Count) <> c_EntryTxt Then ListView2.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > ListView2.ListItems(m_RowIndex).SubItems(4) Then
         ListView2.ListItems(m_RowIndex).Text = "-"
         ListView2.ListItems(m_RowIndex).SubItems(1) = "-"
    Else
        ListView2.ListItems(m_RowIndex).SubItems(1) = ListView2.ListItems(m_RowIndex).Text * ListView2.ListItems(m_RowIndex).SubItems(5)
    End If
    Label3 = ""
    SomaLV ListView2, 1, vPonte5
    If Val(vPonte5) <> 0 Then Label3 = Format(vPonte5, "#,##0.00;(#,##0.00)") Else Label3 = "-"
    
    
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView2's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView2.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

