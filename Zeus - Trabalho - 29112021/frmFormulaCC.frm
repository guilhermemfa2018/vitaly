VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmFormulaCC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F�rmulas"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13830
   Icon            =   "frmFormulaCC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00B7B7B7&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9120
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   75
      Top             =   9840
      Visible         =   0   'False
      Width           =   975
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   6600
      Top             =   9840
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.Frame Frame10 
      Caption         =   "Controle de Produ��o"
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
      Left            =   11160
      TabIndex        =   71
      Top             =   9720
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "Registra no TAOS"
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
         Left            =   240
         TabIndex        =   72
         Tag             =   "Marque caso for necess�rio realizar apontamento de horas do CC no sistema TAOS"
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   13
      Left            =   840
      Picture         =   "frmFormulaCC.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "Sair"
      Top             =   9720
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   240
      Picture         =   "frmFormulaCC.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   46
      Tag             =   "Salvar"
      Top             =   9720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados Centro de Custo "
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
      TabIndex        =   45
      Top             =   120
      Width           =   13575
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
         Height          =   345
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Tag             =   "Nome do Centro de Custo"
         Top             =   480
         Width           =   11295
      End
      Begin VB.TextBox txtformula 
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Identificador do Centro de Custo"
         Top             =   480
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "frmFormulaCC.frx":265E
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmFormulaCC.frx":26C0
         TabIndex        =   52
         Top             =   240
         Width           =   375
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   33
      Top             =   1200
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14843
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
      TabCaption(0)   =   "Centro de Custo"
      TabPicture(0)   =   "frmFormulaCC.frx":271E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "label53"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Padr�o T�cnico"
      TabPicture(1)   =   "frmFormulaCC.frx":273A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      Begin VB.TextBox label53 
         Height          =   285
         Left            =   7680
         TabIndex        =   51
         Text            =   "-"
         Top             =   8040
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame Frame6 
         Caption         =   "Imagem"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Index           =   0
         Left            =   7680
         TabIndex        =   47
         Top             =   5160
         Width           =   5775
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   15
            Left            =   720
            Picture         =   "frmFormulaCC.frx":2756
            Style           =   1  'Graphical
            TabIndex        =   49
            Tag             =   "Excluir imagem"
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   14
            Left            =   120
            Picture         =   "frmFormulaCC.frx":3420
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "Inserir imagem para refer�ncia"
            Top             =   2400
            Width           =   615
         End
         Begin VB.PictureBox Picture1 
            Height          =   2775
            Left            =   2520
            ScaleHeight     =   2715
            ScaleWidth      =   3075
            TabIndex        =   48
            Top             =   240
            Width           =   3135
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
               Height          =   2655
               Left            =   0
               Top             =   0
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   4683
               Image           =   "frmFormulaCC.frx":40EA
            End
         End
         Begin MSComDlg.CommonDialog cdlFoto 
            Left            =   1320
            Top             =   2520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tabela de Classifica��o "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7815
         Left            =   -74880
         TabIndex        =   37
         Top             =   420
         Width           =   13335
         Begin MSComctlLib.ListView ListView3 
            Height          =   4095
            Left            =   120
            TabIndex        =   32
            Top             =   3120
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   7223
            LabelEdit       =   1
            Sorted          =   -1  'True
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
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   11
            Left            =   1320
            Picture         =   "frmFormulaCC.frx":4102
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Excluir"
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   10
            Left            =   720
            Picture         =   "frmFormulaCC.frx":4DCC
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "Editar"
            Top             =   2400
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   9
            Left            =   120
            Picture         =   "frmFormulaCC.frx":5A96
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "Inserir"
            Top             =   2400
            Width           =   615
         End
         Begin VB.Frame Frame9 
            Caption         =   "Defini��es "
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
            Left            =   6960
            TabIndex        =   42
            Top             =   1320
            Width           =   4815
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
               Height          =   345
               Index           =   16
               Left            =   3240
               TabIndex        =   28
               Tag             =   "Organiza��o"
               Top             =   480
               Width           =   1455
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
               Height          =   345
               Index           =   15
               Left            =   1680
               TabIndex        =   27
               Tag             =   "Fadiga"
               Top             =   480
               Width           =   1455
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
               Height          =   345
               Index           =   14
               Left            =   120
               TabIndex        =   26
               Tag             =   "Tempo M�dio"
               Top             =   480
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
               Height          =   255
               Left            =   3240
               OleObjectBlob   =   "frmFormulaCC.frx":6760
               TabIndex        =   69
               Top             =   240
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
               Height          =   255
               Left            =   1680
               OleObjectBlob   =   "frmFormulaCC.frx":67D0
               TabIndex        =   68
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFormulaCC.frx":6836
               TabIndex        =   67
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Intervalos "
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
            Left            =   3600
            TabIndex        =   41
            Top             =   1320
            Width           =   3255
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
               Height          =   345
               Index           =   13
               Left            =   1680
               TabIndex        =   25
               Tag             =   "Intervalo2"
               Top             =   480
               Width           =   1455
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
               Height          =   345
               Index           =   12
               Left            =   120
               TabIndex        =   24
               Tag             =   "Intervalo1"
               Top             =   480
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
               Height          =   255
               Left            =   1680
               OleObjectBlob   =   "frmFormulaCC.frx":68A6
               TabIndex        =   66
               Top             =   240
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFormulaCC.frx":6914
               TabIndex        =   65
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Dimens�es "
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
            TabIndex        =   40
            Top             =   1320
            Width           =   3375
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
               Height          =   345
               Index           =   11
               Left            =   1680
               TabIndex        =   23
               Tag             =   "Dimens�o2"
               Top             =   480
               Width           =   1455
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
               Height          =   345
               Index           =   10
               Left            =   120
               TabIndex        =   22
               Tag             =   "Dimens�o1"
               Top             =   480
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   1680
               OleObjectBlob   =   "frmFormulaCC.frx":6982
               TabIndex        =   64
               Top             =   240
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFormulaCC.frx":69EC
               TabIndex        =   63
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame6 
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
            Height          =   975
            Index           =   1
            Left            =   1560
            TabIndex        =   39
            Top             =   240
            Width           =   10215
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
               Height          =   345
               Index           =   8
               Left            =   120
               TabIndex        =   19
               Tag             =   "ID do Grupo"
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
               Height          =   345
               Index           =   9
               Left            =   960
               TabIndex        =   20
               Tag             =   "Nome do Grupo"
               Top             =   480
               Width           =   7695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   960
               OleObjectBlob   =   "frmFormulaCC.frx":6A58
               TabIndex        =   62
               Top             =   240
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmFormulaCC.frx":6ABA
               TabIndex        =   61
               Top             =   240
               Width           =   495
            End
            Begin VB.CommandButton cmdCadastro 
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
               Index           =   7
               Left            =   8760
               TabIndex        =   44
               Top             =   480
               Width           =   375
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   8
               Left            =   9480
               Picture         =   "frmFormulaCC.frx":6B18
               Style           =   1  'Graphical
               TabIndex        =   21
               Tag             =   "Cadastrar Novo Grupo"
               ToolTipText     =   "Cadastrar Novo Grupo"
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame12 
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
            Height          =   975
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1335
            Begin VB.TextBox txtformula 
               Alignment       =   2  'Center
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
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
               HideSelection   =   0   'False
               Index           =   7
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   70
               Text            =   "ID"
               Top             =   480
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informa��es Gerais "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   36
         Top             =   5220
         Width           =   7455
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
            Height          =   1215
            Index           =   20
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1680
            Width           =   7215
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
            Height          =   1335
            Index           =   6
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   240
            Width           =   7215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contantes "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   7680
         TabIndex        =   35
         Top             =   420
         Width           =   5775
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
            Height          =   345
            Index           =   19
            Left            =   120
            TabIndex        =   12
            Tag             =   "Descri��o da constante"
            Top             =   1200
            Width           =   5535
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
            Height          =   345
            Index           =   17
            Left            =   840
            TabIndex        =   10
            Tag             =   "Constante da f�rmula"
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
            Height          =   345
            Index           =   18
            Left            =   120
            TabIndex        =   11
            Tag             =   "ID Constante"
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFormulaCC.frx":77E2
            TabIndex        =   60
            Top             =   960
            Width           =   5055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFormulaCC.frx":785E
            TabIndex        =   59
            Top             =   240
            Width           =   2055
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFormulaCC.frx":78DC
            TabIndex        =   58
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   6
            Left            =   1320
            Picture         =   "frmFormulaCC.frx":793A
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "Excluir"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   5
            Left            =   720
            Picture         =   "frmFormulaCC.frx":8604
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "Editar"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   4
            Left            =   120
            Picture         =   "frmFormulaCC.frx":92CE
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "Inserir"
            Top             =   1680
            Width           =   615
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   120
            TabIndex        =   16
            Tag             =   "Constantes"
            ToolTipText     =   "Constantes"
            Top             =   2400
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   3625
            LabelEdit       =   1
            Sorted          =   -1  'True
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
      Begin VB.Frame Frame2 
         Caption         =   "F�rmulas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   34
         Top             =   420
         Width           =   7455
         Begin VB.Frame Frame11 
            Caption         =   "Par�metros Autom�tico"
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
            Left            =   3000
            TabIndex        =   73
            Top             =   1680
            Width           =   4215
            Begin VB.CheckBox Check2 
               Caption         =   "Usu�rio n�o informar� par�metros"
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
               Left            =   240
               TabIndex        =   74
               Tag             =   "Marcar quando o sistema utilizar apenas o PESO como par�metro para c�lculo de tempo de fabrica��o"
               Top             =   240
               Width           =   3735
            End
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
            Height          =   345
            Index           =   5
            Left            =   3000
            TabIndex        =   5
            Tag             =   "F�rmula"
            Top             =   1200
            Width           =   4215
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
            Height          =   345
            Index           =   4
            Left            =   120
            TabIndex        =   4
            Tag             =   "Par�metros"
            Top             =   1200
            Width           =   2775
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
            Height          =   345
            Index           =   3
            Left            =   840
            TabIndex        =   3
            Tag             =   "Nome da f�rmula"
            Top             =   480
            Width           =   6495
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
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Tag             =   "Identificador"
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   3000
            OleObjectBlob   =   "frmFormulaCC.frx":9F98
            TabIndex        =   57
            Top             =   960
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFormulaCC.frx":A000
            TabIndex        =   56
            Top             =   960
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   840
            OleObjectBlob   =   "frmFormulaCC.frx":A06E
            TabIndex        =   55
            Top             =   240
            Width           =   2415
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmFormulaCC.frx":A0D0
            TabIndex        =   54
            Top             =   240
            Width           =   375
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   120
            TabIndex        =   9
            Top             =   2400
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   3625
            LabelEdit       =   1
            Sorted          =   -1  'True
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
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   1320
            Picture         =   "frmFormulaCC.frx":A12E
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "Excluir"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   2
            Left            =   720
            Picture         =   "frmFormulaCC.frx":ADF8
            Style           =   1  'Graphical
            TabIndex        =   7
            Tag             =   "Editar"
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   1
            Left            =   120
            Picture         =   "frmFormulaCC.frx":BAC2
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "Inserir"
            Top             =   1680
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmFormulaCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsLocal As New ADODB.Recordset
Private Caminho1 As String
Private vPAutomatico As TextBox

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        'Chama CC - Centro de Custo
'        ChamaGrid "tbCCusto", "nome", txtformula(0), frmFormulaCC, "idprd", "nome"
'        CarregaTxt "tbCCusto", "idprd", "S", "", "", txtformula(0), txtformula(1), 0, 1, txtformula(0), "S", txtformula(1)
'        compoeDadosLVs
    Case 1
        If Check2.Value = 1 Then
            vPAutomatico = "S"
        Else
            vPAutomatico = "N"
        End If
        'Incluir F�rmula no ListView1
        IncluirLV ListView1, txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), vPAutomatico, txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
        LimpaControles txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), txtformula(2), txtformula(2), txtformula(2)
        Check2.Value = 0
        label53.Text = "-"
        txtformula(2) = Format(GeraCodigoLV(ListView1), "000")
        aicAlphaImage1.ClearImage
        label53.Text = "-"
    Case 2
        'Altera f�rmula no ListView1
        LimpaControles txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), txtformula(2), txtformula(2), txtformula(2)
        label53.Text = "-"
        AlteraLV ListView1, txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), vPAutomatico, txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
        If vPAutomatico = "S" Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        
        aicAlphaImage1.ClearImage
        If label53.Text <> "-" Then
            aicAlphaImage1.LoadImage_FromFile (label53.Text)
        End If
    Case 3
        'Excluir F�rmulas no ListView1
        ExcluirItemLV ListView1
        LimpaControles txtformula(0), txtformula(1), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0)
        label53.Text = "-"
        txtformula(2) = Format(GeraCodigoLV(ListView1), "000")
    Case 4
        'Incluir Constantes no ListView2
        IncluirLV ListView2, txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
        LimpaControles txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
        txtformula(18) = Format(GeraCodigoLV(ListView2), "000")
    Case 5
        'Alterar Constante no ListView2
        AlteraLV ListView2, txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
    Case 6
        'Exclui Constantes do ListView2
        ExcluirItemLV ListView2
        LimpaControles txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
        txtformula(18) = Format(GeraCodigoLV(ListView2), "000")
    Case 7
        'Chama Grid Grupo
        If txtformula(0).Text = "" Then
            Msgbox "Selecione primeiro um CC (Centro de Custo)"
            Exit Sub
        End If
        ChamaGrid "tbGrupoClass", "nmgrupo", txtformula(8), frmFormulaCC, "idgrupo", "nmgrupo"
        CarregaTxt "tbGrupoClass", "idprd", "S", "idgrupo", "I", txtformula(0), txtformula(8), 1, 2, txtformula(8), "I", txtformula(9), "1"
    Case 8
        'Cadastra Grupos para o CC (Centro de Custo) selecionado
        If txtformula(0).Text <> "" Then
            frmGrupo.Show 1
        End If
    Case 9
        'Inclui itens na tabela de classifica��o
        IncluirLV ListView3, txtformula(7), txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16), txtformula(7), txtformula(7), txtformula(7), txtformula(7), txtformula(7)
        LimpaControles txtformula(7), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16), txtformula(7), txtformula(7)
        txtformula(7) = Format(GeraCodigoLV(ListView3), "000")
    Case 10 'Altera dados do Item na Tabela de Classifica��o
        AlteraLV ListView3, txtformula(7), txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16), txtformula(7), txtformula(7), txtformula(7), txtformula(7), txtformula(7)
    Case 11
        'Excluir dados na Tabela de Classifica��o
        ExcluirItemLV ListView3
        LimpaControles txtformula(7), txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16)
        txtformula(7) = Format(GeraCodigoLV(ListView3), "000")
    Case 12
        GravaDados
        'Grava dados do formul�rio
        'limpaQualquerDado
        'vQualquerDado(1, 1) = txtformula(0).Text
        'vQualquerDado(1, 2) = "S"
        'vQualquerDado(2, 1) = txtformula(6).Text
        'vQualquerDado(2, 2) = "S"
        'GravaDados "tbProduto", "idprd", "S", txtformula(0), 2
        
        'Grava dados ListView1
        limpaQualquerDado
        ordenaLVArray ListView1, txtformula(0).Text, "0", "1", "2", "3", "4", "5", "6", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbformula", "codreduzido", "S", txtformula(0)
    
        'Grava dados ListView2
        limpaQualquerDado
        ordenaLVArray ListView2, txtformula(0).Text, "0", "1", "2", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbconstantesCC", "idprd", "S", txtformula(0)
    
        'Grava dados ListView3
        limpaQualquerDado
        ordenaLVArray ListView3, txtformula(0).Text, "1", "0", "3", "4", "5", "6", "7", "8", "9", "", "", "", "", "", ""
        GravaDadosLV "tbClassificacao", "idprd", "S", txtformula(0)
        Msgbox "Dados Salvos com sucesso!", vbInformation, "PrototipoX"
    Case 13 'Sair do formul�rio
        Unload Me
    Case 14
        'carregar imagem para o Picture
        With cdlFoto
            .Filter = "(Arquivo *.PNG)|*.png"
            .ShowOpen
            Caminho1 = .FileName
        End With
        'mostra a figura
        aicAlphaImage1.LoadImage_FromFile (Caminho1)
        label53 = Caminho1
    Case 15
        aicAlphaImage1.ClearImage
        label53 = "-"
    End Select
End Sub

Private Sub GravaDados()
On Error GoTo Err
    Dim rsGravaDados As New ADODB.Recordset
    Dim SqlGravaDados As String
    Dim rsGravaDadosPar As New ADODB.Recordset
    Dim SqlGravaDadosPar As String
    
    
    If Check1.Value = 1 Then
        SqlGravaDados = "Select * from tbApropriacao where codreduzido = '" & txtformula(0) & "'"
        rsGravaDados.Open SqlGravaDados, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaDados.RecordCount = 0 Then
            rsGravaDados.AddNew
            rsGravaDados.Fields(0) = txtformula(0)
            rsGravaDados.Update
            rsGravaDados.Close
            Set rsGravaDados = Nothing
        End If
    Else
        SqlGravaDados = "delete from tbApropriacao where codreduzido = '" & txtformula(0) & "'"
        rsGravaDados.Open SqlGravaDados, cnBanco
    End If
    
    
    Dim ItemLst As ListItem
    Dim y As Integer, x As Integer
    y = ListView1.ListItems.Count
    For x = 1 To y
        ListView1.ListItems.Item(x).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(7) = "S" Then
            SqlGravaDadosPar = "Select * from tbParametrosAut where codreduzido = '" & txtformula(0) & "' and idform = '" & Val(ListView1.ListItems.Item(x)) & "'"
            rsGravaDadosPar.Open SqlGravaDadosPar, cnBanco, adOpenKeyset, adLockOptimistic
            If rsGravaDadosPar.RecordCount = 0 Then
                rsGravaDadosPar.AddNew
                rsGravaDadosPar.Fields(0) = txtformula(0)
                rsGravaDadosPar.Fields(1) = Val(ListView1.ListItems.Item(x))
                rsGravaDadosPar.Update
                rsGravaDadosPar.Close
                Set rsGravaDadosPar = Nothing
            End If
        Else
            SqlGravaDadosPar = "delete from tbParametrosAut where codreduzido = '" & txtformula(0) & "' and idform = '" & Val(ListView1.ListItems.Item(x)) & "'"
            rsGravaDadosPar.Open SqlGravaDadosPar, cnBanco
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

Private Sub Form_Load()
    inicializa_tabs SSTab1, Picture2
    Set vPAutomatico = Me.Controls.Add("VB.TextBox", "vPAutomatico")
'    Status = Pesquisa
    SSTab1.Tab = 0
    listview_cabecalho
    LimpaControles txtformula(0), txtformula(1), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0), txtformula(0)
    LimpaControles txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
    LimpaControles txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
    LimpaControles txtformula(7), txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16)
    txtformula(2) = Format(GeraCodigoLV(ListView1), "000")
    txtformula(18) = Format(GeraCodigoLV(ListView2), "000")
    txtformula(7) = Format(GeraCodigoLV(ListView3), "000")
    txtformula(0) = varGlobal
    CarregaTxt "CORPORERM.dbo.GCCUSTO", "codreduzido", "S", "", "", txtformula(0), txtformula(1), 7, 2, txtformula(0), "S", txtformula(1), "1"
    compoeDadosLVs
    carregarIconBotao
    
    
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    'Me.BackColor = vColorThema(Principal.Ribbon.Theme)
    'Me.Refresh
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte T�cnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadastro(1), 1, 46 'Inserir
    carregaImagemBotao cmdCadastro(2), 2, 32 'Editar
    carregaImagemBotao cmdCadastro(3), 3, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(4), 4, 46 'Inserir
    carregaImagemBotao cmdCadastro(5), 5, 32 'Editar
    carregaImagemBotao cmdCadastro(6), 6, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(9), 9, 46 'Inserir
    carregaImagemBotao cmdCadastro(10), 10, 32 'Editar
    carregaImagemBotao cmdCadastro(11), 11, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(14), 14, 47 'Add User
    carregaImagemBotao cmdCadastro(15), 15, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(12), 12, 45 'Salvar
    carregaImagemBotao cmdCadastro(13), 13, 34 'Sair
End Sub

Private Sub ListView1_DblClick()
    LimpaControles txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), txtformula(2), txtformula(2), txtformula(2)
    label53.Text = "-"
    vPAutomatico = ""
    AlteraLV ListView1, txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), label53, txtformula(20), vPAutomatico, txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
    If vPAutomatico <> "" And vPAutomatico <> "N" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    aicAlphaImage1.ClearImage
    If label53.Text <> "-" Then
        aicAlphaImage1.LoadImage_FromFile (label53.Text)
    End If
End Sub

Private Sub ListView2_DblClick()
    AlteraLV ListView2, txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
End Sub

Private Sub ListView3_DblClick()
    AlteraLV ListView3, txtformula(7), txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16), txtformula(7), txtformula(7), txtformula(7), txtformula(7), txtformula(7)
End Sub

Private Sub txtformula_GotFocus(Index As Integer)
    mudaCorText txtformula(Index)
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
            'Abaixo Compoe Listview =========================
            'compoeDadosLVs esta neste formul�rio
            compoeDadosLVs
            '================================================
        End If
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaTxt "tbGrupoClass", "idprd", "S", "idgrupo", "I", txtformula(0), txtformula(8), 1, 2, txtformula(8), "I", txtformula(9), "1"
            'CarregaGrupoClass
        End If
    End Select
End Sub

Private Sub txtformula_LostFocus(Index As Integer)
    voltaCorText txtformula(Index)
End Sub

Private Sub CompoeControles()
On Error GoTo Err
    Dim rsCompoe As New ADODB.Recordset
    Dim sqlCompoe As String
    sqlCompoe = "Select a.observacao from tbProduto as a where a.idprd = '" & txtformula(0) & "'"
    rsCompoe.Open sqlCompoe, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsCompoe.EOF Then
        txtformula(6).Text = rsCompoe.Fields(0) 'Observa��o
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

Private Sub compoeChk()
On Error GoTo Err
    Dim rsChk As New ADODB.Recordset
    Dim SqlChk As String
    SqlChk = "Select * from tbApropriacao where codreduzido = '" & txtformula(0) & "'"
    rsChk.Open SqlChk, cnBanco, adOpenKeyset, adLockReadOnly
    If rsChk.RecordCount > 0 Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    rsChk.Close
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
    LimpaControles txtformula(2), txtformula(3), txtformula(4), txtformula(5), txtformula(6), txtformula(20), txtformula(2), txtformula(2), txtformula(2), txtformula(2)
    LimpaControles txtformula(18), txtformula(17), txtformula(19), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18), txtformula(18)
    LimpaControles txtformula(8), txtformula(9), txtformula(10), txtformula(11), txtformula(12), txtformula(13), txtformula(14), txtformula(15), txtformula(16), txtformula(8)
    compoeChk
    'compoeControles
    'Faz refer�ncias a Fun��es que est�o no: Module1.bas
    'Listview1 - Formulas
    LimpaLV ListView1
    chamaSQL "select a.idform,a.nmform,a.parametros,a.formula,a.observacao,a.imagem,a.observacao2,b.idform from tbFormula as a LEFT join tbParametrosAut as b on a.codreduzido = b.codreduzido and a.idform = b.idform where a.codreduzido = '" & txtformula(0) & "'"
    Compoe_Listview ListView1, Sqlp, "000"
    txtformula(2) = Format(GeraCodigoLV(ListView1), "000")
    
    'Listview2 - Constantes
    LimpaLV ListView2
    chamaSQL "Select a.idseq,a.valconst,a.descricao from tbconstantesCC as a where a.idprd = '" & txtformula(0) & "'"
    Compoe_Listview ListView2, Sqlp, "000"
    txtformula(18) = Format(GeraCodigoLV(ListView2), "000")
    
    'Listview3 - Classifica��o
    LimpaLV ListView3
    chamaSQL "select a.idseq,a.idgrupo,b.nmgrupo,a.dim1,a.dim2,a.inter1,a.inter2,a.tmedio,a.fadiga,a.organizacao from tbClassificacao as a inner join tbgrupoclass as b on b.idprd = a.idprd and a.idgrupo = b.idgrupo where a.idprd = '" & txtformula(0) & "'"
    Compoe_Listview ListView3, Sqlp, "000"
    txtformula(7) = Format(GeraCodigoLV(ListView3), "000")
End Sub


Private Sub listview_cabecalho()
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 14
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 3.8
    ListView1.ColumnHeaders.Add , , "Par�metros", ListView1.Width / 7
    ListView1.ColumnHeaders.Add , , "F�rmula", ListView1.Width / 2.3
    ListView1.ColumnHeaders.Add , , "Dica", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "imagem", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Observa��o", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "PAutomatico", ListView1.Width / 10000
    
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 6
    ListView2.ColumnHeaders.Add , , "Valor constante", ListView2.Width / 4
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 2
    
    ListView3.ColumnHeaders.Add , , "Seq.", ListView3.Width / 18
    ListView3.ColumnHeaders.Add , , "IdGrupo", ListView3.Width / 16
    ListView3.ColumnHeaders.Add , , "Grupo", ListView3.Width / 7
    ListView3.ColumnHeaders.Add , , "Dim1", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Dim2", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Intervalo1", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Intervalo2", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "T. M�dio", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Fadiga", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Organiza��o", ListView3.Width / 10
    
    Me.ListView3.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(5).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(8).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(9).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(10).Alignment = lvwColumnRight
    
    ListView1.View = lvwReport 'Modo de Exibi��o do seu Listview
    ListView2.View = lvwReport
    ListView3.View = lvwReport
End Sub

