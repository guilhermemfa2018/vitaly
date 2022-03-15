VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmConfSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema"
   ClientHeight    =   11355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmConfSistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11355
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   840
      Picture         =   "frmConfSistema.frx":37E04
      Style           =   1  'Graphical
      TabIndex        =   175
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   10680
      Width           =   615
   End
   Begin VB.CommandButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   240
      Picture         =   "frmConfSistema.frx":38ACE
      Style           =   1  'Graphical
      TabIndex        =   176
      Tag             =   "Salvar Grupo"
      ToolTipText     =   "Salvar Grupo"
      Top             =   10680
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10335
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   18230
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Configurações"
      TabPicture(0)   =   "frmConfSistema.frx":39798
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame25"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Parametrizações"
      TabPicture(1)   =   "frmConfSistema.frx":397B4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Empresa/Coligadas"
      TabPicture(2)   =   "frmConfSistema.frx":397D0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Servidor - email"
      TabPicture(3)   =   "frmConfSistema.frx":397EC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame9"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Menu"
      TabPicture(4)   =   "frmConfSistema.frx":39808
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame25 
         Caption         =   "Controle de Atividades de Desenvolvimento "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   120
         TabIndex        =   162
         Top             =   4560
         Width           =   9015
         Begin VB.CommandButton cmdCadAtividade 
            Height          =   615
            Index           =   3
            Left            =   1920
            Picture         =   "frmConfSistema.frx":39824
            Style           =   1  'Graphical
            TabIndex        =   166
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdCadAtividade 
            Height          =   615
            Index           =   2
            Left            =   1320
            Picture         =   "frmConfSistema.frx":3A4EE
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdCadAtividade 
            Height          =   615
            Index           =   1
            Left            =   720
            Picture         =   "frmConfSistema.frx":3B1B8
            Style           =   1  'Graphical
            TabIndex        =   168
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdCadAtividade 
            Height          =   615
            Index           =   0
            Left            =   120
            Picture         =   "frmConfSistema.frx":3BE82
            Style           =   1  'Graphical
            TabIndex        =   169
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   165
            ToolTipText     =   "Informe a última alteração realizada no sistema"
            Top             =   480
            Width           =   8775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":3CB4C
            TabIndex        =   164
            Top             =   240
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView7 
            Height          =   1815
            Left            =   120
            TabIndex        =   163
            ToolTipText     =   "Mantém o histórico das últimas alterações realizadas no sistema"
            Top             =   2160
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3201
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
         Caption         =   "Tipo de FCE (Define o nível de complexidade da FCE)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   143
         Top             =   480
         Width           =   9015
         Begin VB.Frame Frame24 
            Caption         =   "Cor de Identificação do Tipo "
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
            Left            =   3840
            TabIndex        =   149
            Top             =   1560
            Width           =   3135
            Begin VB.TextBox txtCorTipo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
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
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   151
               Text            =   "8421631"
               Top             =   300
               Width           =   1935
            End
            Begin VB.CommandButton cmdCad 
               Caption         =   "..."
               Height          =   375
               Index           =   3
               Left            =   2520
               TabIndex        =   150
               Top             =   240
               Width           =   495
            End
         End
         Begin MSComDlg.CommonDialog dlgCores 
            Left            =   3120
            Top             =   1800
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdCadEmail 
            Height          =   615
            Index           =   16
            Left            =   1920
            Picture         =   "frmConfSistema.frx":3CBB6
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadEmail 
            Height          =   615
            Index           =   17
            Left            =   1320
            Picture         =   "frmConfSistema.frx":3D880
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadEmail 
            Height          =   615
            Index           =   18
            Left            =   720
            Picture         =   "frmConfSistema.frx":3E54A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadEmail 
            Height          =   615
            Index           =   19
            Left            =   120
            Picture         =   "frmConfSistema.frx":3F214
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1680
            Width           =   615
         End
         Begin VB.Frame Frame23 
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
            Left            =   7800
            TabIndex        =   147
            Top             =   1560
            Width           =   1095
            Begin VB.CheckBox Check10 
               Caption         =   "Ativo"
               Height          =   255
               Left            =   120
               TabIndex        =   148
               Top             =   300
               Value           =   1  'Checked
               Width           =   735
            End
         End
         Begin VB.TextBox txtCadTipo 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Tag             =   "Código  do treinamento"
            ToolTipText     =   "Código  do treinamento"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtCadTipo 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            Tag             =   "Conteúdo da avaliação"
            ToolTipText     =   "Conteúdo da avaliação"
            Top             =   600
            Width           =   7575
         End
         Begin VB.TextBox txtCadTipo 
            Height          =   285
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   1200
            Width           =   8775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":3FEDE
            TabIndex        =   144
            Top             =   960
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Index           =   1
            Left            =   1320
            OleObjectBlob   =   "frmConfSistema.frx":3FF4A
            TabIndex        =   145
            Top             =   360
            Width           =   1695
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Index           =   1
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":3FFAC
            TabIndex        =   146
            Top             =   360
            Width           =   495
         End
         Begin MSComctlLib.ListView ListView6 
            Height          =   1335
            Left            =   120
            TabIndex        =   7
            Top             =   2400
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   2355
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
         Caption         =   "Estrutura do Menu"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9855
         Left            =   -74880
         TabIndex        =   86
         Top             =   360
         Width           =   9015
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   4
            Left            =   1920
            Picture         =   "frmConfSistema.frx":4000A
            Style           =   1  'Graphical
            TabIndex        =   177
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   5
            Left            =   1320
            Picture         =   "frmConfSistema.frx":40CD4
            Style           =   1  'Graphical
            TabIndex        =   179
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   720
            Picture         =   "frmConfSistema.frx":4199E
            Style           =   1  'Graphical
            TabIndex        =   178
            Top             =   1680
            Width           =   615
         End
         Begin VB.CommandButton cmdCadastro 
            Height          =   615
            Index           =   2
            Left            =   120
            Picture         =   "frmConfSistema.frx":42668
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   1680
            Width           =   615
         End
         Begin VB.Frame Frame27 
            Caption         =   "Selecione a colação de ícones para o sistema"
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
            Left            =   4200
            TabIndex        =   173
            Top             =   1080
            Width           =   4695
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
               ItemData        =   "frmConfSistema.frx":43332
               Left            =   240
               List            =   "frmConfSistema.frx":43342
               TabIndex        =   174
               Text            =   "Padão"
               Top             =   360
               Width           =   4095
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Identificador "
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
            Left            =   7200
            TabIndex        =   104
            Top             =   360
            Width           =   1695
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":43398
               TabIndex        =   106
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "      Ícone "
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
            Left            =   5640
            TabIndex        =   95
            Top             =   360
            Width           =   1335
            Begin ZEUS.chameleonButton cmdCadastro1 
               Height          =   255
               Index           =   6
               Left            =   840
               TabIndex        =   107
               Top             =   240
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   450
               BTYPE           =   3
               TX              =   "..."
               ENAB            =   0   'False
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
               MICON           =   "frmConfSistema.frx":433F2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox Text7 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   103
               Tag             =   "Ícone"
               ToolTipText     =   "Ícone"
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox Check9 
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   102
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "      Tipo "
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
            Left            =   120
            TabIndex        =   92
            Top             =   360
            Width           =   1215
            Begin VB.ComboBox Combo2 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmConfSistema.frx":4340E
               Left            =   120
               List            =   "frmConfSistema.frx":4341B
               TabIndex        =   94
               Tag             =   "Tipo"
               Text            =   "TAB"
               ToolTipText     =   "Tipo"
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox Check8 
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "      Botão "
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
            Height          =   615
            Left            =   4200
            TabIndex        =   91
            Top             =   360
            Width           =   1335
            Begin VB.ComboBox Combo4 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   101
               Tag             =   "Botão"
               ToolTipText     =   "Botão"
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox Check7 
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "      Submenu "
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
            Height          =   615
            Left            =   2640
            TabIndex        =   90
            Top             =   360
            Width           =   1455
            Begin VB.ComboBox Combo3 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmConfSistema.frx":4342E
               Left            =   120
               List            =   "frmConfSistema.frx":43430
               TabIndex        =   99
               Tag             =   "Submenu"
               ToolTipText     =   "Submenu"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox Check3 
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "      Menu "
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
            Height          =   615
            Left            =   1440
            TabIndex        =   89
            Top             =   360
            Width           =   1095
            Begin VB.ComboBox Combo1 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "frmConfSistema.frx":43432
               Left            =   120
               List            =   "frmConfSistema.frx":43454
               TabIndex        =   97
               Tag             =   "Menu"
               Text            =   "01"
               ToolTipText     =   "Menu"
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox Check2 
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   105
            Tag             =   "Nome"
            ToolTipText     =   "Nome"
            Top             =   1320
            Width           =   3975
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   7335
            Left            =   120
            TabIndex        =   87
            Top             =   2400
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   12938
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
         Begin VB.Label Label2 
            Caption         =   "Nome:"
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
            TabIndex        =   88
            Top             =   1080
            Width           =   975
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   9735
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   17171
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
         TabCaption(0)   =   "Empresa/coligada ativa"
         TabPicture(0)   =   "frmConfSistema.frx":43480
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Empresas/Coligadas"
         TabPicture(1)   =   "frmConfSistema.frx":4349C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chameleonButton5"
         Tab(1).Control(1)=   "chameleonButton4"
         Tab(1).Control(2)=   "imgColigada"
         Tab(1).Control(3)=   "ListView3"
         Tab(1).ControlCount=   4
         Begin VB.CommandButton chameleonButton5 
            Height          =   615
            Left            =   -74160
            Picture         =   "frmConfSistema.frx":434B8
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton chameleonButton4 
            Height          =   615
            Left            =   -74760
            Picture         =   "frmConfSistema.frx":44182
            Style           =   1  'Graphical
            TabIndex        =   186
            Top             =   480
            Width           =   615
         End
         Begin MSComctlLib.ImageList imgColigada 
            Left            =   -74760
            Top             =   1560
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmConfSistema.frx":44E4C
                  Key             =   "OK"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmConfSistema.frx":4585E
                  Key             =   "EXC"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   8415
            Left            =   -74880
            TabIndex        =   35
            Top             =   1200
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   14843
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "imgColigada"
            SmallIcons      =   "imgColigada"
            ColHdrIcons     =   "imgColigada"
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
         Begin VB.Frame Frame2 
            Caption         =   "Dados da empresa/coligada ativa"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   8775
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   16
               Left            =   720
               Picture         =   "frmConfSistema.frx":46270
               Style           =   1  'Graphical
               TabIndex        =   182
               Top             =   3720
               Width           =   615
            End
            Begin VB.CommandButton cmdCadastro 
               Height          =   615
               Index           =   15
               Left            =   120
               Picture         =   "frmConfSistema.frx":46F3A
               Style           =   1  'Graphical
               TabIndex        =   181
               Top             =   3720
               Width           =   615
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   10
               Left            =   3720
               TabIndex        =   48
               Top             =   3240
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   3360
               OleObjectBlob   =   "frmConfSistema.frx":47C04
               TabIndex        =   63
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   8
               Left            =   3720
               TabIndex        =   46
               Top             =   2880
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   3360
               OleObjectBlob   =   "frmConfSistema.frx":47C64
               TabIndex        =   62
               Top             =   3000
               Width           =   495
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   5
               Left            =   1320
               TabIndex        =   43
               Top             =   2160
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   4
               Left            =   2640
               TabIndex        =   42
               Top             =   1800
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   2160
               OleObjectBlob   =   "frmConfSistema.frx":47CC4
               TabIndex        =   61
               Top             =   1920
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47D24
               TabIndex        =   60
               Top             =   3360
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47D86
               TabIndex        =   59
               Top             =   3000
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47DF0
               TabIndex        =   58
               Top             =   2640
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47E52
               TabIndex        =   57
               Top             =   2280
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47EB6
               TabIndex        =   56
               Top             =   1920
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47F16
               TabIndex        =   55
               Top             =   1560
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47F7C
               TabIndex        =   54
               Top             =   1200
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":47FE2
               TabIndex        =   53
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtDadosEmpresa 
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   1320
               TabIndex        =   36
               Tag             =   "Código da coligada"
               ToolTipText     =   "Código da coligada"
               Top             =   360
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Index           =   0
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":4804C
               TabIndex        =   52
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   0
               Left            =   2160
               TabIndex        =   37
               Tag             =   "Razão social"
               ToolTipText     =   "Razão social"
               Top             =   360
               Width           =   3495
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   38
               Top             =   720
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   2
               Left            =   1320
               TabIndex        =   39
               Top             =   1080
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   3
               Left            =   1320
               TabIndex        =   40
               Top             =   1440
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   6
               Left            =   1320
               TabIndex        =   44
               Top             =   2520
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   7
               Left            =   1320
               TabIndex        =   45
               Top             =   2880
               Width           =   1935
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   9
               Left            =   1320
               TabIndex        =   47
               Top             =   3240
               Width           =   1935
            End
            Begin VB.ComboBox cboDadosEmpresa 
               Height          =   315
               ItemData        =   "frmConfSistema.frx":480BE
               Left            =   1320
               List            =   "frmConfSistema.frx":48113
               TabIndex        =   41
               Top             =   1800
               Width           =   735
            End
            Begin VB.Frame Frame6 
               Caption         =   "Logo"
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
               Index           =   0
               Left            =   5760
               TabIndex        =   30
               Top             =   240
               Width           =   2895
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   13
                  Left            =   720
                  Picture         =   "frmConfSistema.frx":48183
                  Style           =   1  'Graphical
                  TabIndex        =   183
                  Top             =   3120
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadastro 
                  Height          =   615
                  Index           =   12
                  Left            =   120
                  Picture         =   "frmConfSistema.frx":48E4D
                  Style           =   1  'Graphical
                  TabIndex        =   184
                  Top             =   3120
                  Width           =   615
               End
               Begin VB.PictureBox Picture2 
                  Height          =   2775
                  Left            =   120
                  ScaleHeight     =   2715
                  ScaleWidth      =   2595
                  TabIndex        =   49
                  Top             =   240
                  Width           =   2655
                  Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
                     Height          =   2655
                     Left            =   120
                     Top             =   0
                     Width           =   2415
                     _ExtentX        =   4260
                     _ExtentY        =   4683
                     Image           =   "frmConfSistema.frx":49B17
                  End
                  Begin VB.Label Label59 
                     Alignment       =   2  'Center
                     Caption         =   "A Imagem não se encontra no local especificado"
                     Height          =   495
                     Left            =   240
                     TabIndex        =   31
                     Top             =   1200
                     Visible         =   0   'False
                     Width           =   2055
                  End
               End
               Begin MSComDlg.CommonDialog cdlFoto 
                  Left            =   1800
                  Top             =   3240
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
            End
            Begin VB.Label Label53 
               BackColor       =   &H8000000C&
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   3840
               Visible         =   0   'False
               Width           =   5415
            End
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9735
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   17171
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
         TabCaption(0)   =   "Gerais"
         TabPicture(0)   =   "frmConfSistema.frx":49B2F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame8"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Integração"
         TabPicture(1)   =   "frmConfSistema.frx":49B4B
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "SSTab3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Check4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame15"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame3"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.Frame Frame3 
            Caption         =   "SGBD "
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
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   120
            TabIndex        =   67
            Top             =   840
            Width           =   4095
            Begin VB.OptionButton optIntegra 
               Caption         =   "SQL Server"
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
               Index           =   0
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Value           =   -1  'True
               Width           =   2895
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Sistema "
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
            ForeColor       =   &H8000000D&
            Height          =   615
            Left            =   4320
            TabIndex        =   65
            Top             =   840
            Width           =   4455
            Begin VB.OptionButton chkIntegra 
               Caption         =   "Totvs - RM Labore (11.40)"
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
               Index           =   0
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Integrar o Zeus ao Totvs "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   480
            Width           =   4695
         End
         Begin VB.Frame Frame8 
            Caption         =   "Gerais "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   9135
            Left            =   -74880
            TabIndex        =   20
            Top             =   420
            Width           =   8775
            Begin VB.Frame Frame26 
               Caption         =   "Tabs Ativas"
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
               TabIndex        =   170
               Top             =   1200
               Width           =   3135
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":49B67
                  TabIndex        =   172
                  Top             =   600
                  Width           =   2895
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   4
                  Left            =   120
                  TabIndex        =   171
                  Text            =   "1"
                  Top             =   240
                  Width           =   2775
               End
            End
            Begin VB.Frame Frame22 
               Caption         =   "E-mails SRM- Solicitação de Retirada deMat."
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
               Left            =   4320
               TabIndex        =   135
               Top             =   5760
               Width           =   4095
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   12
                  Left            =   1920
                  Picture         =   "frmConfSistema.frx":49BF9
                  Style           =   1  'Graphical
                  TabIndex        =   136
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   13
                  Left            =   1320
                  Picture         =   "frmConfSistema.frx":4A8C3
                  Style           =   1  'Graphical
                  TabIndex        =   137
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   14
                  Left            =   720
                  Picture         =   "frmConfSistema.frx":4B58D
                  Style           =   1  'Graphical
                  TabIndex        =   138
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   3
                  Left            =   120
                  TabIndex        =   140
                  Top             =   480
                  Width           =   3855
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   15
                  Left            =   120
                  Picture         =   "frmConfSistema.frx":4C257
                  Style           =   1  'Graphical
                  TabIndex        =   139
                  Top             =   960
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":4CF21
                  TabIndex        =   141
                  Top             =   240
                  Width           =   1815
               End
               Begin MSComctlLib.ListView ListView5 
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   142
                  Top             =   1680
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   2566
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
            End
            Begin VB.Frame Frame21 
               Caption         =   "E-mails SI- Solicitação de Inspeção"
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
               Left            =   120
               TabIndex        =   127
               Top             =   5760
               Width           =   4095
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   8
                  Left            =   1920
                  Picture         =   "frmConfSistema.frx":4CF87
                  Style           =   1  'Graphical
                  TabIndex        =   128
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   9
                  Left            =   1320
                  Picture         =   "frmConfSistema.frx":4DC51
                  Style           =   1  'Graphical
                  TabIndex        =   129
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   10
                  Left            =   720
                  Picture         =   "frmConfSistema.frx":4E91B
                  Style           =   1  'Graphical
                  TabIndex        =   130
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   11
                  Left            =   120
                  Picture         =   "frmConfSistema.frx":4F5E5
                  Style           =   1  'Graphical
                  TabIndex        =   132
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   2
                  Left            =   120
                  TabIndex        =   131
                  Top             =   480
                  Width           =   3855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":502AF
                  TabIndex        =   133
                  Top             =   240
                  Width           =   1815
               End
               Begin MSComctlLib.ListView ListView4 
                  Height          =   1455
                  Left            =   120
                  TabIndex        =   134
                  Top             =   1680
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   2566
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
            End
            Begin VB.Frame Frame19 
               Caption         =   "E-mails RNC- Registro de Não Conformidade"
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
               Left            =   4320
               TabIndex        =   119
               Top             =   2280
               Width           =   4095
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   4
                  Left            =   1920
                  Picture         =   "frmConfSistema.frx":50315
                  Style           =   1  'Graphical
                  TabIndex        =   120
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   5
                  Left            =   1320
                  Picture         =   "frmConfSistema.frx":50FDF
                  Style           =   1  'Graphical
                  TabIndex        =   121
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   6
                  Left            =   720
                  Picture         =   "frmConfSistema.frx":51CA9
                  Style           =   1  'Graphical
                  TabIndex        =   122
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   0
                  Left            =   120
                  TabIndex        =   124
                  Top             =   480
                  Width           =   3855
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   7
                  Left            =   120
                  Picture         =   "frmConfSistema.frx":52973
                  Style           =   1  'Graphical
                  TabIndex        =   123
                  Top             =   960
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":5363D
                  TabIndex        =   125
                  Top             =   240
                  Width           =   1815
               End
               Begin MSComctlLib.ListView ListView2 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   126
                  Top             =   1680
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   2778
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
            End
            Begin VB.Frame Frame20 
               Caption         =   "E-mails CD - Comunicação de Desvio"
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
               TabIndex        =   111
               Top             =   2280
               Width           =   4095
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   3
                  Left            =   1920
                  Picture         =   "frmConfSistema.frx":536A3
                  Style           =   1  'Graphical
                  TabIndex        =   118
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   2
                  Left            =   1320
                  Picture         =   "frmConfSistema.frx":5436D
                  Style           =   1  'Graphical
                  TabIndex        =   117
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   1
                  Left            =   720
                  Picture         =   "frmConfSistema.frx":55037
                  Style           =   1  'Graphical
                  TabIndex        =   116
                  Top             =   960
                  Width           =   615
               End
               Begin VB.CommandButton cmdCadEmail 
                  Height          =   615
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmConfSistema.frx":55D01
                  Style           =   1  'Graphical
                  TabIndex        =   115
                  Top             =   960
                  Width           =   615
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   1
                  Left            =   120
                  TabIndex        =   112
                  Top             =   480
                  Width           =   3855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":569CB
                  TabIndex        =   113
                  Top             =   240
                  Width           =   1815
               End
               Begin MSComctlLib.ListView ListView1 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   114
                  Top             =   1680
                  Width           =   3855
                  _ExtentX        =   6800
                  _ExtentY        =   2778
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
            End
            Begin VB.Frame Frame18 
               Caption         =   "Iniciar relatórios de Inspeção/Expedição em:"
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
               Left            =   3360
               TabIndex        =   108
               Top             =   1200
               Width           =   5295
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":56A31
                  TabIndex        =   110
                  Top             =   600
                  Width           =   4935
               End
               Begin VB.TextBox Text4 
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   109
                  Top             =   240
                  Width           =   5055
               End
            End
            Begin VB.Frame Frame17 
               Caption         =   "           Atualizações automáticas"
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
               Left            =   3360
               TabIndex        =   50
               Top             =   240
               Width           =   5295
               Begin VB.CommandButton cmdCad 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  Height          =   255
                  Index           =   2
                  Left            =   4680
                  TabIndex        =   25
                  Top             =   360
                  Width           =   375
               End
               Begin MSComDlg.CommonDialog cdlTXT2 
                  Left            =   4800
                  Top             =   240
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin VB.CheckBox Check6 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   23
                  Top             =   0
                  Width           =   375
               End
               Begin VB.TextBox Text2 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   120
                  TabIndex        =   24
                  Text            =   "Informe o caminho do executável: AtualizaSGCH.exe"
                  Top             =   360
                  Width           =   4455
               End
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Exibir avisos ao logar"
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
               TabIndex        =   22
               Top             =   720
               Width           =   2175
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Ativar arquivo de LOG"
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
               TabIndex        =   21
               Top             =   300
               Width           =   2775
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   3375
            Left            =   120
            TabIndex        =   69
            Top             =   1680
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5953
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
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
            TabCaption(0)   =   "RM Sistemas"
            TabPicture(0)   =   "frmConfSistema.frx":56AE3
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Label25"
            Tab(0).Control(1)=   "Label24"
            Tab(0).Control(2)=   "Label23"
            Tab(0).Control(3)=   "Label22"
            Tab(0).Control(4)=   "txtIntegra(0)"
            Tab(0).Control(5)=   "txtIntegra(1)"
            Tab(0).Control(6)=   "txtIntegra(2)"
            Tab(0).Control(7)=   "txtIntegra(3)"
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Relógio de Ponto"
            TabPicture(1)   =   "frmConfSistema.frx":56AFF
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cdlTXT3"
            Tab(1).Control(1)=   "SkinLabel25"
            Tab(1).Control(2)=   "SkinLabel22"
            Tab(1).Control(3)=   "SkinLabel21"
            Tab(1).Control(4)=   "SkinLabel20"
            Tab(1).Control(5)=   "SkinLabel19"
            Tab(1).Control(6)=   "txtIntegra(4)"
            Tab(1).Control(7)=   "txtIntegra(5)"
            Tab(1).Control(8)=   "txtIntegra(6)"
            Tab(1).Control(9)=   "txtIntegra(7)"
            Tab(1).Control(10)=   "txtIntegra(10)"
            Tab(1).Control(11)=   "cmdCad(5)"
            Tab(1).ControlCount=   12
            TabCaption(2)   =   "Banco FlexJR.FDB"
            TabPicture(2)   =   "frmConfSistema.frx":56B1B
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Label1"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "SkinLabel24"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "cdlTXT4"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "SkinLabel23"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "Text5"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "cmdCad(4)"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "txtIntegra(8)"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).Control(7)=   "txtIntegra(9)"
            Tab(2).Control(7).Enabled=   0   'False
            Tab(2).ControlCount=   8
            Begin VB.CommandButton cmdCad 
               Caption         =   "..."
               Height          =   255
               Index           =   5
               Left            =   -71760
               TabIndex        =   161
               ToolTipText     =   "Localizar arquivo"
               Top             =   2280
               Width           =   375
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   10
               Left            =   -74760
               TabIndex        =   159
               ToolTipText     =   "Caminho onde será gravado o arquivo com os dados capturados no relógio de ponto"
               Top             =   2280
               Width           =   2895
            End
            Begin VB.TextBox txtIntegra 
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
               IMEMode         =   3  'DISABLE
               Index           =   9
               Left            =   3360
               PasswordChar    =   "*"
               TabIndex        =   81
               ToolTipText     =   "Senha de acesso ao banco FDB"
               Top             =   1560
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   8
               Left            =   240
               TabIndex        =   80
               ToolTipText     =   "Usuário de acesso ao banco FDB"
               Top             =   1560
               Width           =   2895
            End
            Begin VB.CommandButton cmdCad 
               Caption         =   "..."
               Height          =   255
               Index           =   4
               Left            =   7200
               TabIndex        =   79
               ToolTipText     =   "Localizar arquivo"
               Top             =   840
               Width           =   375
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   240
               TabIndex        =   78
               Text            =   "Informe o caminho do arquivo FDB"
               ToolTipText     =   "Caminho onde se encontra o arquivo FDB do Firebird"
               Top             =   840
               Width           =   6855
            End
            Begin VB.TextBox txtIntegra 
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
               IMEMode         =   3  'DISABLE
               Index           =   7
               Left            =   -71640
               PasswordChar    =   "*"
               TabIndex        =   77
               ToolTipText     =   "Senha de acesso ao relógio de ponto"
               Top             =   1560
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   6
               Left            =   -74760
               TabIndex        =   76
               ToolTipText     =   "CPF do responsável cadastrado no relógio de ponto"
               Top             =   1560
               Width           =   2895
            End
            Begin VB.TextBox txtIntegra 
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
               Left            =   -71640
               TabIndex        =   75
               ToolTipText     =   "IP do relógio de ponto"
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   4
               Left            =   -74760
               TabIndex        =   74
               ToolTipText     =   "Identificação do relógio de ponto"
               Top             =   840
               Width           =   2895
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
               Height          =   255
               Left            =   -74760
               OleObjectBlob   =   "frmConfSistema.frx":56B37
               TabIndex        =   152
               Top             =   600
               Width           =   2175
            End
            Begin VB.TextBox txtIntegra 
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
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   -71640
               PasswordChar    =   "*"
               TabIndex        =   73
               Top             =   1560
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   2
               Left            =   -74760
               TabIndex        =   72
               Top             =   1560
               Width           =   2895
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   1
               Left            =   -71640
               TabIndex        =   71
               Top             =   840
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
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
               Index           =   0
               Left            =   -74760
               TabIndex        =   70
               Top             =   840
               Width           =   2895
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
               Height          =   255
               Left            =   -71640
               OleObjectBlob   =   "frmConfSistema.frx":56BC1
               TabIndex        =   153
               ToolTipText     =   "IP do relógio"
               Top             =   600
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
               Height          =   255
               Left            =   -74760
               OleObjectBlob   =   "frmConfSistema.frx":56C3D
               TabIndex        =   154
               ToolTipText     =   "CPF do responsável cadastrado no relógio"
               Top             =   1320
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
               Height          =   375
               Left            =   -71640
               OleObjectBlob   =   "frmConfSistema.frx":56CBB
               TabIndex        =   155
               ToolTipText     =   "Senha de acesso ao relógio"
               Top             =   1320
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "frmConfSistema.frx":56D1F
               TabIndex        =   156
               Top             =   600
               Width           =   3615
            End
            Begin MSComDlg.CommonDialog cdlTXT4 
               Left            =   7800
               Top             =   720
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
               Height          =   255
               Left            =   240
               OleObjectBlob   =   "frmConfSistema.frx":56DC1
               TabIndex        =   157
               ToolTipText     =   "CPF do responsável cadastrado no relógio"
               Top             =   1320
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
               Height          =   255
               Left            =   -74760
               OleObjectBlob   =   "frmConfSistema.frx":56E29
               TabIndex        =   160
               ToolTipText     =   "Caminho onde será gravado o arquivo com os dados capturados do relógio de ponto"
               Top             =   2040
               Width           =   2415
            End
            Begin MSComDlg.CommonDialog cdlTXT3 
               Left            =   -71160
               Top             =   2160
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Label Label1 
               Caption         =   "Senha:"
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
               Left            =   3360
               TabIndex        =   158
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label Label22 
               Caption         =   "Senha:"
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
               Left            =   -71640
               TabIndex        =   85
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label Label23 
               Caption         =   "Usuário:"
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
               Left            =   -74760
               TabIndex        =   84
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label Label24 
               Caption         =   "Nome do BANCO:"
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
               Left            =   -71640
               TabIndex        =   83
               Top             =   600
               Width           =   2775
            End
            Begin VB.Label Label25 
               Caption         =   "Nome do SERVIDOR:"
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
               Left            =   -74760
               TabIndex        =   82
               Top             =   600
               Width           =   2175
            End
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Autenticação de email "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   16
         Top             =   1980
         Width           =   8895
         Begin VB.TextBox txtEmail 
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
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   4920
            PasswordChar    =   "*"
            TabIndex        =   11
            Tag             =   "Senha do usuario de autenticação"
            ToolTipText     =   "Senha do usuario de autenticação"
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox txtEmail 
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
            Left            =   120
            TabIndex        =   10
            Tag             =   "usuario de autenticação"
            ToolTipText     =   "usuario de autenticação"
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label19 
            Caption         =   "Senha:"
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
            Left            =   4920
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Usuário:"
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
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Servidor SMTP"
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   900
         Width           =   8895
         Begin ACTIVESKINLibCtl.SkinLabel Label16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":56EAF
            TabIndex        =   51
            Top             =   600
            Width           =   8655
         End
         Begin VB.TextBox txtEmail 
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
            TabIndex        =   9
            Tag             =   "Endereço do servidor de SMTP"
            ToolTipText     =   "Endereço do servidor de SMTP"
            Top             =   240
            Width           =   8655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione a tabela a ser importada "
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
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   8760
         Width           =   9015
         Begin ZEUS.chameleonButton chameleonButton1 
            Height          =   735
            Left            =   1560
            TabIndex        =   8
            Tag             =   "Importar dados"
            ToolTipText     =   "Importar dados"
            Top             =   360
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1296
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
            MICON           =   "frmConfSistema.frx":56F59
            PICN            =   "frmConfSistema.frx":56F75
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame16 
            Caption         =   "Importar colaboradores"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   5160
            TabIndex        =   26
            Top             =   240
            Width           =   3735
            Begin VB.CommandButton cmdCad 
               Caption         =   "Importar"
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
               Index           =   1
               Left            =   1560
               TabIndex        =   34
               Tag             =   "Importar"
               ToolTipText     =   "Importar"
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmdCad 
               Caption         =   "Localizar..."
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
               Index           =   0
               Left            =   120
               TabIndex        =   33
               Tag             =   "Localizar"
               ToolTipText     =   "Localizar"
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   3495
            End
            Begin MSComDlg.CommonDialog cdlTXT 
               Left            =   3000
               Top             =   480
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Materiais"
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
            TabIndex        =   14
            Top             =   360
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmConfSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Coloque estas declarações em um módulo
'Se colocar em um formulário lembre-se de não usar como 'Public'

'Existem outras flags para parametrizar a pesquisa

'Abaixo: declarações realizadas para selecionar pasta
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" ( _
lpbi As BrowseInfo _
) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" ( _
ByVal pidList As Long, _
ByVal lpBuffer As String _
) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" ( _
ByVal lpString1 As String, _
ByVal lpString2 As String _
) As Long
'Tipo para def
Private Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Caminho2 As String
Public Caminho3 As String
Public Caminho4 As String
Public Caminho5 As String
Public vTab As String
Public vCAT As String
Public X As Integer
Public no As Node
Public vChave As String
Public vChaveTAB As String
Private vPonte1 As TextBox
Private vPonte2 As TextBox

Private Sub chameleonButton1_Click()
On Error GoTo Err
    mobjMsg.Abrir "Deseja realmente importar os dados das tabelas selecionadas?", YesNo, pergunta, "ZEUS"
    If Tp = 2 Then
        Exit Sub
    End If
    
    If Option1.Value = True Then ImportaDadosCargo
    If Option2.Value = True Then ImportaDadosHabilidade
    If Option3.Value = True Then ImportaDadosAvaliacao
    If Option4.Value = True Then ImportaDadosEscolaridade
    If Option5.Value = True Then ImportaDadosDepartamento
    If Option6.Value = True Then ImportaDadosSetor
    
    If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False Then
        mobjMsg.Abrir "Nenhuma tabela selecionada. Marque a tabela a ser importada", Ok, critico, "ZEUS"
        Exit Sub
    End If
    
    'A ROTINA ABAIXO VC SELECIONA UM PROCESSO Q ESTA NA MEMORIA P SER REMOVIDO
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    '--------------------------------------------------------------------------
    mobjMsg.Abrir "Dados importados com sucesso. Para vizualisar os dados feche a tabela e abra novamente", Ok, informacao, "ZEUS"
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
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

Private Sub ImportaDadosCargo()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsCargos As New ADODB.Recordset
    Dim SqlCargos As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Cargos").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela de CARGOS
    SqlCargos = "select * from tbCargos where codcoligada = '" & vCodcoligada & "'"
    rsCargos.Open SqlCargos, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de CARGOS..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de CARGOS..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsCargos.AddNew
            rsCargos.Fields(0) = .Range("A" & j).Value 'Código do CARGO
            rsCargos.Fields(1) = .Range("B" & j).Value 'Código do CBO
            rsCargos.Fields(2) = .Range("C" & j).Value 'Nome do CARGO
            rsCargos.Fields(5) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsCargos.Update
    rsCargos.Close
    Set rsCargos = Nothing
'---------------------------------------------
    
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    Legenda = "Dados importados com sucesso!"
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

Private Sub ImportaDadosHabilidade()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidade As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Habilidades").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela de HABILIDADES
    sqlHabilidade = "select * from tbHabilidades where codcoligada = '" & vCodcoligada & "'"
    rsHabilidade.Open sqlHabilidade, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de Habilidades..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Habilidade..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsHabilidade.AddNew
            rsHabilidade.Fields(0) = .Range("A" & j).Value 'Código da Habilidade
            rsHabilidade.Fields(1) = .Range("B" & j).Value 'Habilidade
            rsHabilidade.Fields(2) = .Range("D" & j).Value 'Peso da Habilidade
            rsHabilidade.Fields(3) = .Range("C" & j).Value 'Descrição da Habilidade
            rsHabilidade.Fields(4) = "S" 'Status
            rsHabilidade.Fields(5) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsHabilidade.Update
    rsHabilidade.Close
    Set rsHabilidade = Nothing
'---------------------------------------------
    
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    
    Legenda = "Dados importados com sucesso!"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Existem dados cadastrados na tabela de Habilidades do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
        Legenda = "ERRO na importação de dados"
        Principal.StatusBar1.Panels(3).Text = Legenda
        Exit Sub
    End If
End Sub

Private Sub ImportaDadosAvaliacao()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsAvaliacao As New ADODB.Recordset
    Dim sqlAvaliacao As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Avaliacao").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela de AVALIACAO
    sqlAvaliacao = "select * from tbAvaliacao where codcoligada = '" & vCodcoligada & "'"
    rsAvaliacao.Open sqlAvaliacao, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de Avaliacão do Treinamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - AVALIACAO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Avaliação do Treinamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsAvaliacao.AddNew
            rsAvaliacao.Fields(0) = .Range("A" & j).Value 'Código da avaliação
            rsAvaliacao.Fields(1) = .Range("B" & j).Value 'Nome da avaliação
            rsAvaliacao.Fields(2) = .Range("C" & j).Value 'Tipo da avaliação
            rsAvaliacao.Fields(3) = .Range("D" & j).Value 'Peso da avaliação
            rsAvaliacao.Fields(4) = "S" 'Status
            rsAvaliacao.Fields(5) = .Range("E" & j).Value 'Descrição da avaliação
            rsAvaliacao.Fields(6) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsAvaliacao.Update
    rsAvaliacao.Close
    Set rsAvaliacao = Nothing
'---------------------------------------------
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    Legenda = "Dados importados com sucesso!"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Existem dados cadastrados na tabela de Avaliação do Treinamento do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
        Legenda = "ERRO na importação de dados"
        Principal.StatusBar1.Panels(3).Text = Legenda
        Exit Sub
    End If
End Sub

Private Sub ImportaDadosEscolaridade()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsEscolaridade As New ADODB.Recordset
    Dim sqlEscolaridade As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Escolaridade").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela de ESCOLARIDADE
    sqlEscolaridade = "select * from tbEscolaridade where codcoligada = '" & vCodcoligada & "'"
    rsEscolaridade.Open sqlEscolaridade, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de Escolaridade..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - ESCOLARIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Escolaridade..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsEscolaridade.AddNew
            rsEscolaridade.Fields(0) = .Range("A" & j).Value 'Código da escolaridade
            rsEscolaridade.Fields(1) = .Range("B" & j).Value 'Nome da escolaridade
            rsEscolaridade.Fields(2) = .Range("C" & j).Value 'Peso da escolaridade
            rsEscolaridade.Fields(3) = "S" 'Status
            rsEscolaridade.Fields(4) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsEscolaridade.Update
    rsEscolaridade.Close
    Set rsEscolaridade = Nothing
'---------------------------------------------
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    Legenda = "Dados importados com sucesso!"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Existem dados cadastrados na tabela de Escolaridade do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
        Legenda = "ERRO na importação de dados"
        Principal.StatusBar1.Panels(3).Text = Legenda
        Exit Sub
    End If
End Sub

Private Sub ImportaDadosDepartamento()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsDepartamento As New ADODB.Recordset
    Dim sqlDepartamento As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Departamento").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela do DEPARTAMENTO
    sqlDepartamento = "select * from tbDepartamentos where codcoligada = '" & vCodcoligada & "'"
    rsDepartamento.Open sqlDepartamento, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de Departamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - DEPARTAMENTO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Departamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsDepartamento.AddNew
            rsDepartamento.Fields(0) = .Range("A" & j).Value 'Código do departamento
            rsDepartamento.Fields(1) = .Range("B" & j).Value 'Nome do departamento
            rsDepartamento.Fields(2) = .Range("C" & j).Value 'descrição do departamento
            rsDepartamento.Fields(3) = "S" 'Status
            rsDepartamento.Fields(4) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsDepartamento.Update
    rsDepartamento.Close
    Set rsDepartamento = Nothing
'---------------------------------------------
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    Legenda = "Dados importados com sucesso!"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Existem dados cadastrados na tabela de Departamentos do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
        Legenda = "ERRO na importação de dados"
        Principal.StatusBar1.Panels(3).Text = Legenda
        Exit Sub
    End If
End Sub

Private Sub ImportaDadosSetor()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    
    Dim rsSetor As New ADODB.Recordset
    Dim SqlSetor As String
    
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\tabela de importação.xls"
    Plan.Visible = False 'Indica q a planilha do Excel a ser utilizada nao estará visível
    Plan.UserControl = False
    Plan.Sheets("Setor").Select ' Seleciona a planilha q vc vai trabalhar

'----> Importa Dados para tabela de SETOR
    SqlSetor = "select * from tbSetores where codcoligada = '" & vCodcoligada & "'"
    rsSetor.Open SqlSetor, cnBanco, adOpenKeyset, adLockOptimistic
    Legenda = "Aguarde, analisando tabela de Setores..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    'ABAIXO - CONTA A QUANTIDADE DE CONSTANTES CADASTRADOS NA PLANILHA ANTES DE IMPORTAR
    'PARA O PROGRESSBAR PODER TRABALHAR
    '**********************************************************************
    j = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & j).Value = "" Then Exit For
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Max = j
    
    'PREENCHE CÉLULAS DESEJADAS - SETORES
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Setores..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    j = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & j).Value = "" Then Exit For
            rsSetor.AddNew
            rsSetor.Fields(2) = .Range("A" & j).Value 'Código do departamento
            rsSetor.Fields(0) = .Range("B" & j).Value 'Código do setor
            rsSetor.Fields(1) = .Range("C" & j).Value 'Nome do setor
            rsSetor.Fields(3) = .Range("C" & j).Value 'Descrição do setor
            rsSetor.Fields(4) = "S" 'Status
            rsSetor.Fields(5) = vCodcoligada 'Codigo da coligada
            j = j + 1
        End With
    Next
    Principal.ProgressBar1.Value = 0
    rsSetor.Update
    rsSetor.Close
    Set rsSetor = Nothing
'---------------------------------------------
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.Application.Quit ' Fecha o Excel automaticamente
    Set Plan = Nothing ' Libera o espaço reservado na memoria para esta variavel
    Legenda = "Dados importados com sucesso!"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Existem dados cadastrados na tabela de Setores do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "ZEUS"
        Legenda = "ERRO na importação de dados"
        Principal.StatusBar1.Panels(3).Text = Legenda
        Exit Sub
    End If
End Sub

Private Sub chameleonButton4_Click()
    AlteraColigada
    SSTab4.Tab = 0
End Sub

Private Sub chameleonButton5_Click()
    mobjMsg.Abrir "Rotina em desenvolvimento", Ok, critico, "Atenção"
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Combo1.Enabled = True
        If Combo2.Text <> "TAB" Then
            Frame7.Enabled = True
            Check3.Enabled = True
            CompoeCombo3
        End If
    Else
        Combo1.Enabled = False
        Check3.Enabled = False
        Check9.Enabled = False
        Check7.Enabled = False
        Check9.Value = 0
        Check7.Value = 0
        Check3.Value = 0
        Frame7.Enabled = False
        Frame11.Enabled = False
        Frame13.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Text7.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Combo3.Enabled = True
        If Combo2.Text <> "TAB" And Combo2.Text <> "CAT" Then
            Frame11.Enabled = True
            Check7.Enabled = True
        End If
    Else
        Check7.Value = 0
        Check9.Value = 0
        Check7.Enabled = False
        Check9.Enabled = False
        Frame11.Enabled = False
        Frame13.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Text7.Enabled = False
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Frame3.Enabled = True
        Frame15.Enabled = True
        SSTab3.Enabled = True
    Else
        Frame3.Enabled = False
        Frame15.Enabled = False
        SSTab3.Enabled = False
    End If
End Sub

Private Sub Check6_Click()
    If Check6.Value = 1 Then
        Text2.Enabled = True
        cmdCad(2).Enabled = True
    Else
        Text2.Enabled = False
        'cmdCadastro(17).Enabled = False
        Text2 = "Informe o caminho do executável: AtualizaZEUSH.exe"
    End If
End Sub

Private Sub Check7_Click()
    If Check7.Value = 1 Then
        Combo4.Enabled = True
        Frame13.Enabled = True
        Check9.Enabled = True
        CompoeCombo4
    Else
        Combo4.Enabled = False
        Frame13.Enabled = False
        Check9.Enabled = False
        Check9.Value = 0
    End If
End Sub

Private Sub Check8_Click()
    If Check8.Value = 1 Then
        Combo2.Enabled = True
        Check2.Enabled = True
        Frame5.Enabled = True
        Frame12.Enabled = True
        Check8.Enabled = True
    Else
        Check9.Enabled = False
        Check7.Enabled = False
        Check3.Enabled = False
        Check2.Enabled = False
        Check9.Value = 0
        Check7.Value = 0
        Check3.Value = 0
        Check2.Value = 0
        Frame5.Enabled = False
        Frame7.Enabled = False
        Frame11.Enabled = False
        Frame13.Enabled = False
        Combo1.Enabled = False
        Combo2.Enabled = False
        Combo3.Enabled = False
        Combo4.Enabled = False
        Text7.Enabled = False
    End If
End Sub

Private Sub Check9_Click()
    If Check9.Value = 1 Then
        Text7.Enabled = True
        cmdCadastro(6).Enabled = True
    Else
        Text7.Enabled = False
        cmdCadastro(6).Enabled = False
    End If
End Sub

Private Sub cmdCad_Click(Index As Integer)
    Select Case Index
    Case 0
        'carregar arquivo texto
        With cdlTXT
            .Filter = "(Arquivo *.TXT)|*.txt"
            .ShowOpen
            Caminho2 = .FileName
        End With
        Text1 = Caminho2
        If Text1.Text <> "" Then cmdCad(1).Enabled = True
    Case 1
        importaColaboradores
    Case 2
'        carregaPasta
        With cdlTXT2
            .Filter = "(AtualizaZEUSH *.EXE)|*.exe"
            .ShowOpen
            Caminho3 = .FileName
        End With
        Text2 = Caminho3
    Case 3
        selectColor txtCorTipo, dlgCores
        txtCorTipo.Text = txtCorTipo.BackColor
    Case 4
        With cdlTXT4
            .Filter = "(Arquivo *.FDB)|*.fdb"
            .ShowOpen
            Caminho5 = .FileName
        End With
        Text5 = Caminho5
    Case 5
        With cdlTXT3
            .DialogTitle = "Selecione um diretório"
            .InitDir = App.Path
            .FileName = "Selecione um diretório"
            .flags = cdlOFNNoValidate + cdlOFNHideReadOnly
            .Filter = "(Diretórios|*.~#~"
            .CancelError = True
            .ShowSave
            Caminho4 = .FileName
        End With
        txtIntegra(10) = CurDir + "\"
    
    End Select
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados de parametrização?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            GravaParametros
            'gravaLog "Mádia para aprovação: " & txtCadParametro(0), "Gerar introdutório: " & Check3.Value, "Aprovação com restrição: " & txtCadParametro(1)
            Pesquisa = 0
            'Unload Me
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela configurações do sistema?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            Pesquisa = 0
            Unload Me
            Set frmConfSistema = Nothing
        End If
    Case 2
        IncluiTreeview
        CompoeTAB
        limpaControlTree
    Case 3
        AlteraTreeview
        CompoeDados
    Case 4
        DeletaTreeview
        CompoeTAB
'        LimpaControlesAvaliacao
    Case 5
        limpaControlTree
'        IncluirAvaliacao
    Case 6
        frmIcones.Show 1
 
 '       mobjMsg.Abrir "Deseja EXCLUIR essa ocorrência?", YesNo, pergunta, "ZEUS"
 '       If Tp = 1 Then
 '           ExcluirItemLV ListView2
 '       End If
    Case 7
 '       AlteraABS
    Case 8
 '       LimpaControlesABS
    Case 9
 '       IncluirABS
    Case 10
 '       If ListView1.ListItems.Count > 0 Then
 '           carregaADP
 '           mobjMsg.Abrir "Rotina de Avaliação de Desempenho efetuada com sucesso!", Ok, informacao, "ZEUS"
 '       Else
 '           mobjMsg.Abrir "É necessário cadastrar primeiramente os períodos de Avaliação de Desempenho Profissional", Ok, informacao, "ZEUS"
 '       End If
    Case 12
        'carregar imagem para o Picture
        With cdlFoto
            .Filter = "(Arquivo *.JPG)|*.jpg"
            .ShowOpen
            Caminho1 = .FileName
        End With
        'mostra a figura
        'Image1.Picture = LoadPicture(Caminho1)
        aicAlphaImage1.LoadImage_FromFile (Caminho1)
        Label53 = Caminho1
    Case 13
        aicAlphaImage1.ClearImage
        Label53 = "-"
    Case 15
        LimpaControlesColigada
    Case 16
        IncluirColigada
        'criaUsuEMenu Val(txtDadosEmpresa(11) - 1)
    End Select
End Sub

Private Sub IncluiTreeview()
    'If ValidaCampoTree = False Then Exit Sub
On Error GoTo Err
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    
10  cnBanco.BeginTrans
    
    SqlMenu = "Select * from tbMenuConf as a where a.id = '" & Val(SkinLabel13) & "' and a.codcoligada = '" & Val(vCodcoligada) & "'"
    rsMenu.Open SqlMenu, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsMenu.RecordCount = 0 Then
        rsMenu.AddNew
        rsMenu.Fields(4) = GeraCodigoMenu
    End If
    
    rsMenu.Fields(0) = Val(Combo1.Text)
    If Combo2 = "TAB" Then
        rsMenu.Fields(1) = Combo1.Text
    End If
    If Combo2 = "CAT" Then
        rsMenu.Fields(1) = Combo1.Text & Combo3.Text
    End If
    If Combo2 = "BUT" Then
        rsMenu.Fields(1) = Combo1.Text & Combo3.Text & Combo4.Text
    End If
    rsMenu.Fields(2) = Combo2.Text
    rsMenu.Fields(3) = Text3.Text ' Nome
    rsMenu.Fields(5) = vCodcoligada
    If Check9.Value = 1 Then
        rsMenu.Fields(6) = Text7.Text ' Icone
    Else
        rsMenu.Fields(6) = 0
    End If

    rsMenu.Update
    rsMenu.Close
    Set rsMenu = Nothing
    SkinLabel13 = Format(GeraCodigoMenu, "000000")
    
'---------------------------------
'EM TESTE
    
'    If Val(SkinLabel13) <> 0 Then
'
'        'tbConfGrupo = idgrupo/idmenu/idsub/tipo
'        'tbMenuConf  = ......./idMenu/idsub/idtipo
'        'TENHO QUE ATUALIZAR A TABELA tbConfGrupo AO ADICIONAR/ATUALIZAR OU REMOVER DADOS NA TABELA tbMenuConf
'        If Check7.Value = 1 Then
'            SqlMenu = "Update tbConfGrupo set nome = '" & Text3.Text & "', icon = '" & Text7.Text & "' where idmenu = '" & Val(Combo1) & "' and idsub = '" & Combo3 & Combo4 & "'"
 '       Else
 '           SqlMenu = "Update tbConfGrupo set nome = '" & Text3.Text & "', icon = '" & Text7.Text & "' where idmenu = '" & Val(Combo1) & "' and idsub = '" & Combo3 & "'"
 '       End If
'        rsMenu.Open SqlMenu, cnBanco
'    End If
'EM TESTE
'---------------------------------
    
    cnBanco.CommitTrans
    Exit Sub
    'CompoeTreeview
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Não é permitido duplicação de registros", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub limpaControlTree()
    Check8.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Check4.Value = 0
    Check9.Value = 0
    Check8.Value = 0
    Text3.Text = ""
    SkinLabel13 = 0
    SkinLabel13 = Format(GeraCodigoMenu, "000000")
End Sub

Private Sub CompoeTAB()
On Error GoTo Err
    Dim rsTAB As New ADODB.Recordset
    Dim SqlTAB
    Dim Y As Integer, Contador As Integer
    Dim vProc As String
    TreeView1.Nodes.Clear
    Dim vTes, vTexto As String

    SqlTAB = "Select * from tbmenuconf where codcoligada = '" & vCodcoligada & "' order by idsub"
    rsTAB.Open SqlTAB, cnBanco, adOpenKeyset, adLockReadOnly
    'On Error Resume Next
    Do While Not rsTAB.EOF
        X = rsTAB.Fields(0)
        vTexto = rsTAB.Fields(3)
        If rsTAB.Fields(2) = "TAB" Then
            vChaveTAB = rsTAB.Fields(2) & rsTAB.Fields(1)
            TreeView1.Nodes.Add , , vChaveTAB, Format(rsTAB.Fields(4), "000000") & " - " & vTexto
        End If
        If rsTAB.Fields(2) = "CAT" Then
            vTes = rsTAB.Fields(3)
            TreeView1.Nodes.Add vChaveTAB, tvwChild, vTes, Format(rsTAB.Fields(4), "000000") & " - " & vTexto
            vChave = vTes
        End If
        If rsTAB.Fields(2) = "BUT" Then
            vTes = rsTAB.Fields(2) & Right$(rsTAB.Fields(1), 5)
            TreeView1.Nodes.Add vChave, tvwChild, vTes, Format(rsTAB.Fields(4), "000000") & " - " & vTexto
        End If
        If Not rsTAB.EOF Then rsTAB.MoveNext Else Exit Do
    Loop
    rsTAB.Close
    Set rsTAB = Nothing
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
            If Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 0 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, 1, 6)
            ElseIf Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 1 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 6)
            ElseIf Len(TreeView1.Nodes(llng_Contador).FullPath) - Len(Replace(TreeView1.Nodes(llng_Contador).FullPath, "\", "")) = 2 Then
                SkinLabel13 = Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + InStr(Mid$(TreeView1.Nodes(llng_Contador).FullPath, InStr(TreeView1.Nodes(llng_Contador).FullPath, "\") + 1, 100), "\") + 1, 6)
            End If
        End If
    Next
End Sub

Private Sub CompoeDados()
On Error GoTo Err
    Dim rsCompoeDados As New ADODB.Recordset
    Dim SqlCompoeDados As String
    SqlCompoeDados = "Select * from tbMenuConf where id = '" & Val(SkinLabel13) & "'"
    rsCompoeDados.Open SqlCompoeDados, cnBanco, adOpenKeyset, adLockOptimistic
    
    Check8.Enabled = False
    Combo2.Enabled = False
    Frame12.Enabled = False

    Check2.Enabled = False
    Combo1.Enabled = False
    Frame5.Enabled = False
    
    Check3.Enabled = False
    Combo3.Enabled = False
    Frame7.Enabled = False
    
    Check7.Enabled = False
    Combo4.Enabled = False
    Frame11.Enabled = False
    
    Check9.Enabled = False
    Text7.Enabled = False
    Frame13.Enabled = False
    
    If rsCompoeDados.Fields(2) = "TAB" Then
        Check8.Enabled = True
        Combo2.Enabled = True
        Frame12.Enabled = True
        Check8.Value = 1
        Combo2.Text = rsCompoeDados.Fields(2)
        
        Check2.Enabled = True
        Combo1.Enabled = True
        Frame5.Enabled = True
        Check2.Value = 1
        Combo1.Text = Format(rsCompoeDados.Fields(0), "00")
        Text3.Text = rsCompoeDados.Fields(3)
    End If
    If rsCompoeDados.Fields(2) = "CAT" Then
        Check8.Enabled = True
        Combo2.Enabled = True
        Frame12.Enabled = True
        Check8.Value = 1
        Combo2.Text = rsCompoeDados.Fields(2)
        
        Check2.Enabled = True
        Combo1.Enabled = True
        Frame5.Enabled = True
        Check2.Value = 1
        Combo1.Text = Format(rsCompoeDados.Fields(0), "00")
        
        Check3.Enabled = True
        Combo3.Enabled = True
        Frame7.Enabled = True
        Check3.Value = 1
        Combo3.Text = Mid$(rsCompoeDados.Fields(1), 3, 3)
        Text3.Text = rsCompoeDados.Fields(3)
    End If
    If rsCompoeDados.Fields(2) = "BUT" Then
        Check8.Enabled = True
        Combo2.Enabled = True
        Frame12.Enabled = True
        Check8.Value = 1
        Combo2.Text = rsCompoeDados.Fields(2)
        
        Check2.Enabled = True
        Combo1.Enabled = True
        Frame5.Enabled = True
        Check2.Value = 1
        Combo1.Text = Format(rsCompoeDados.Fields(0), "00")
        
        Check3.Enabled = True
        Combo3.Enabled = True
        Frame7.Enabled = True
        Check3.Value = 1
        Combo3.Text = Mid$(rsCompoeDados.Fields(1), 3, 3)
        
        Check7.Enabled = True
        Combo4.Enabled = True
        Frame11.Enabled = True
        Check7.Value = 1
        Combo4.Text = Mid$(rsCompoeDados.Fields(1), 6, 2)
        
        Check9.Enabled = True
        Text7.Enabled = True
        Frame13.Enabled = True
        Check9.Value = 1
        Text7.Text = rsCompoeDados.Fields(6)
        
        
        Text3.Text = rsCompoeDados.Fields(3)
    End If
    rsCompoeDados.Close
    Set rsCompoeDados = Nothing
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

Private Sub DeletaTreeview()
On Error GoTo Err
    Dim rsDeleta As New ADODB.Recordset
    Dim SqlDeleta As String
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Msgbox("Confirma Exclusão", vbQuestion + vbYesNo, "ZEUS") = vbYes Then
                SqlDeleta = "Delete from tbMenuConf where id = '" & Val(SkinLabel13) & "' and codcoligada= '" & vCodcoligada & "'"
                rsDeleta.Open SqlDeleta, cnBanco, adOpenKeyset, adLockOptimistic
                
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

Private Sub carregaPasta()
    'carregar arquivo texto
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    'Personaliza a procura
    szTitle = "Titulo da procura"
    With tBrowseInfo
        .hWndOwner = Me.HWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_EDITBOX
    End With

    'Abre a janela de procura
    'E retorna o caminho da pasta selecionada
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    'Se existir alguma pasta selecionada extrair
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Text2.Text = sBuffer
    End If
End Sub

Private Sub cmdINTD_Click(Index As Integer)
    vCodModeloAval = Val(Label14)
    frmADPModelo.Show 1
End Sub

Private Sub CompoeCombo3()
    'If Combo3 <> "" Then Combo3.Clear
    Combo3.AddItem "001"
    Combo3.AddItem "011"
    Combo3.AddItem "021"
    Combo3.AddItem "031"
    Combo3.AddItem "041"
    Combo3.AddItem "051"
    Combo3.AddItem "061"
    Combo3.AddItem "071"
    Combo3.AddItem "081"
    Combo3.AddItem "091"
End Sub

Private Sub CompoeCombo4()
    'Combo4.Clear
    If Combo3 >= 1 And Combo3 < 11 Then
        Combo4.AddItem "01"
        Combo4.AddItem "02"
        Combo4.AddItem "03"
        Combo4.AddItem "04"
        Combo4.AddItem "05"
        Combo4.AddItem "06"
        Combo4.AddItem "07"
        Combo4.AddItem "08"
        Combo4.AddItem "09"
        Combo4.AddItem "10"
    End If
    If Combo3 >= 11 And Combo3 < 21 Then
        Combo4.AddItem "11"
        Combo4.AddItem "12"
        Combo4.AddItem "13"
        Combo4.AddItem "14"
        Combo4.AddItem "15"
        Combo4.AddItem "16"
        Combo4.AddItem "17"
        Combo4.AddItem "18"
        Combo4.AddItem "19"
        Combo4.AddItem "20"
    End If
    If Combo3 >= 21 And Combo3 < 31 Then
        Combo4.AddItem "21"
        Combo4.AddItem "22"
        Combo4.AddItem "23"
        Combo4.AddItem "24"
        Combo4.AddItem "25"
        Combo4.AddItem "26"
        Combo4.AddItem "27"
        Combo4.AddItem "28"
        Combo4.AddItem "29"
        Combo4.AddItem "30"
    End If
    If Combo3 >= 31 And Combo3 < 41 Then
        Combo4.AddItem "31"
        Combo4.AddItem "32"
        Combo4.AddItem "33"
        Combo4.AddItem "34"
        Combo4.AddItem "35"
        Combo4.AddItem "36"
        Combo4.AddItem "37"
        Combo4.AddItem "38"
        Combo4.AddItem "39"
        Combo4.AddItem "40"
    End If
    If Combo3 >= 41 And Combo3 < 51 Then
        Combo4.AddItem "41"
        Combo4.AddItem "42"
        Combo4.AddItem "43"
        Combo4.AddItem "44"
        Combo4.AddItem "45"
        Combo4.AddItem "46"
        Combo4.AddItem "47"
        Combo4.AddItem "48"
        Combo4.AddItem "49"
        Combo4.AddItem "50"
    End If
    If Combo3 >= 51 And Combo3 < 61 Then
        Combo4.AddItem "51"
        Combo4.AddItem "52"
        Combo4.AddItem "53"
        Combo4.AddItem "54"
        Combo4.AddItem "55"
        Combo4.AddItem "56"
        Combo4.AddItem "57"
        Combo4.AddItem "58"
        Combo4.AddItem "59"
        Combo4.AddItem "60"
    End If
    If Combo3 >= 61 And Combo3 < 71 Then
        Combo4.AddItem "61"
        Combo4.AddItem "62"
        Combo4.AddItem "63"
        Combo4.AddItem "64"
        Combo4.AddItem "65"
        Combo4.AddItem "66"
        Combo4.AddItem "67"
        Combo4.AddItem "68"
        Combo4.AddItem "69"
        Combo4.AddItem "70"
    End If
    If Combo3 >= 71 And Combo3 < 81 Then
        Combo4.AddItem "71"
        Combo4.AddItem "72"
        Combo4.AddItem "73"
        Combo4.AddItem "74"
        Combo4.AddItem "75"
        Combo4.AddItem "76"
        Combo4.AddItem "77"
        Combo4.AddItem "78"
        Combo4.AddItem "79"
        Combo4.AddItem "80"
    End If
    If Combo3 >= 81 And Combo3 < 91 Then
        Combo4.AddItem "81"
        Combo4.AddItem "82"
        Combo4.AddItem "83"
        Combo4.AddItem "84"
        Combo4.AddItem "85"
        Combo4.AddItem "86"
        Combo4.AddItem "87"
        Combo4.AddItem "88"
        Combo4.AddItem "89"
        Combo4.AddItem "90"
    End If
End Sub

Private Sub cmdCadAtividade_Click(Index As Integer)
    Select Case Index
    Case 0
        vPonte1 = Format(Date, "dd/MM/YYYY") + " " + Format(Time, "hh:mm")
        IncluirLV ListView7, Text6, vPonte1, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
        LimpaControles Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
    Case 1
        AlteraLV ListView7, Text6, vPonte1, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
    Case 2
        LimpaControles Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
    Case 3
        ExcluirItemLV ListView7
    End Select
End Sub

Private Sub cmdCadEmail_Click(Index As Integer)
    Select Case Index
    Case 0
        vPonte1 = "CD"
        IncluirLV ListView1, txtCadParametro(1), vPonte1, txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
        LimpaControles txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
    Case 1
        AlteraLV ListView1, txtCadParametro(1), vPonte1, txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
    Case 2
        LimpaControles txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
    Case 3
        ExcluirItemLV ListView1
    Case 4
        ExcluirItemLV ListView2
    Case 5
        LimpaControles txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0)
    Case 6
        AlteraLV ListView2, txtCadParametro(0), vPonte1, txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0)
    Case 7
        vPonte1 = "RNC"
        IncluirLV ListView2, txtCadParametro(0), vPonte1, txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0)
        LimpaControles txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0), txtCadParametro(0)
    Case 8
        ExcluirItemLV ListView4
    Case 9
        LimpaControles txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2)
    Case 10
        AlteraLV ListView4, txtCadParametro(2), vPonte1, txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2)
    Case 11
        vPonte1 = "SI"
        IncluirLV ListView4, txtCadParametro(2), vPonte1, txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2)
        LimpaControles txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2), txtCadParametro(2)
    Case 12
        ExcluirItemLV ListView5
    Case 13
        LimpaControles txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3)
    Case 14
        AlteraLV ListView5, txtCadParametro(3), vPonte1, txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3)
    Case 15
        vPonte1 = "SRM"
        IncluirLV ListView5, txtCadParametro(3), vPonte1, txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3)
        LimpaControles txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3), txtCadParametro(3)
    Case 16
        ExcluirItemLV ListView6
    Case 17
        LimpaControles txtCadTipo(0), txtCadTipo(1), txtCadTipo(2), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0)
        txtCadTipo(0).Text = Format(GeraCodigoLV(ListView6), "00")
    Case 18
        AlteraLV ListView6, txtCadTipo(0), txtCadTipo(1), txtCadTipo(2), vPonte1, txtCorTipo, vPonte2, txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(1), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0)
        If vPonte1 = "S" Then
            Check10.Value = 1 ' Status do tipo de FCE
        Else
            Check10.Value = 0 ' Status do tipo de FCE
        End If
        txtCorTipo.BackColor = txtCorTipo.Text
    Case 19
        If Check10.Value = 1 Then
            vPonte1 = "S" ' Status do critério
        Else
            vPonte1 = "N" ' Status do critério
        End If
        vPonte2 = vCodcoligada
        IncluirLV ListView6, txtCadTipo(0), txtCadTipo(1), txtCadTipo(2), vPonte1, txtCorTipo, vPonte2, txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0)
        LimpaControles txtCadTipo(0), txtCadTipo(1), txtCadTipo(2), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0)
        txtCadTipo(0).Text = Format(GeraCodigoLV(ListView6), "00")
    End Select
End Sub

Private Sub Combo5_Click()
    vColectionIcons = Combo5.ListIndex
End Sub

Private Sub Form_Load()
    Set vPonte1 = Me.Controls.Add("VB.TextBox", "vPonte1")
    Set vPonte2 = Me.Controls.Add("VB.TextBox", "vPonte2")
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    SSTab4.Tab = 0
    CarregaParametros
    configControles
    listview_cabecalho
    Compoe_ListviewConf
    SkinLabel13 = Format(GeraCodigoMenu, "000000")
    chamaSQL "Select a.id,a.nome,a.descricao,ativo,cor,codcoligada from tbTipoFCE as a where a.ativo = 'S' "
    Compoe_Listview ListView6, Sqlp, "00"
    chamaSQL "SELECT DESCRICAO,DATAHORA FROM TBCONTROLEATIVIDADES"
    Compoe_Listview ListView7, Sqlp, ""
    txtCadTipo(0).Text = Format(GeraCodigoLV(ListView6), "00")
    CompoeTAB

    Me.ListView7.Sorted = True
    Me.ListView7.SortKey = 1
    Me.ListView7.SortOrder = lvwDescending

    carregarIconBotao

    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub carregarIconBotao()
    carregaImagemBotao cmdCadEmail(19), 19, 46 'Inserir
    carregaImagemBotao cmdCadEmail(18), 18, 32 'Editar
    carregaImagemBotao cmdCadEmail(17), 18, 31 'Novo
    carregaImagemBotao cmdCadEmail(16), 16, 33 'Excluir
    carregaImagemBotao cmdCadEmail(0), 0, 46 'Inserir
    carregaImagemBotao cmdCadEmail(1), 1, 32 'Editar
    carregaImagemBotao cmdCadEmail(2), 2, 31 'Novo
    carregaImagemBotao cmdCadEmail(3), 3, 33 'Excluir
    carregaImagemBotao cmdCadEmail(7), 7, 46 'Inserir
    carregaImagemBotao cmdCadEmail(6), 6, 32 'Editar
    carregaImagemBotao cmdCadEmail(5), 5, 31 'Novo
    carregaImagemBotao cmdCadEmail(4), 4, 33 'Excluir
    carregaImagemBotao cmdCadEmail(11), 15, 46 'Inserir
    carregaImagemBotao cmdCadEmail(10), 14, 32 'Editar
    carregaImagemBotao cmdCadEmail(9), 13, 31 'Novo
    carregaImagemBotao cmdCadEmail(8), 12, 33 'Excluir
    carregaImagemBotao cmdCadEmail(15), 15, 46 'Inserir
    carregaImagemBotao cmdCadEmail(14), 14, 32 'Editar
    carregaImagemBotao cmdCadEmail(13), 13, 31 'Novo
    carregaImagemBotao cmdCadEmail(12), 12, 33 'Excluir
    
    carregaImagemBotao cmdCadAtividade(0), 0, 46 'Inserir
    carregaImagemBotao cmdCadAtividade(1), 1, 32 'Editar
    carregaImagemBotao cmdCadAtividade(2), 2, 31 'Novo
    carregaImagemBotao cmdCadAtividade(3), 3, 33 'Excluir
    
    carregaImagemBotao chameleonButton4, 4, 32 'Editar
    carregaImagemBotao chameleonButton5, 5, 33 'Excluir

    carregaImagemBotao cmdCadastro(15), 15, 32 'Novo
    carregaImagemBotao cmdCadastro(13), 13, 33 'Excluir
    carregaImagemBotao cmdCadastro(16), 16, 48 'Seta Direita
    carregaImagemBotao cmdCadastro(12), 12, 47 'Add User
    carregaImagemBotao cmdCadastro(2), 2, 46 'Inserir
    carregaImagemBotao cmdCadastro(3), 3, 32 'Editar
    carregaImagemBotao cmdCadastro(5), 5, 31 'Novo
    carregaImagemBotao cmdCadastro(4), 4, 33 'Excluir
    
    carregaImagemBotao cmdCadastro(0), 0, 45 'Salvar
    carregaImagemBotao cmdCadastro(1), 1, 34 'Sair
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Email", ListView1.Width / 1.5
    ListView1.ColumnHeaders.Add , , "Módulo", ListView1.Width / 5

    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Email", ListView2.Width / 1.5
    ListView2.ColumnHeaders.Add , , "Módulo", ListView2.Width / 5

    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "Email", ListView4.Width / 1.5
    ListView4.ColumnHeaders.Add , , "Módulo", ListView4.Width / 5

    ListView5.ColumnHeaders.Clear
    ListView5.ColumnHeaders.Add , , "Email", ListView5.Width / 1.5
    ListView5.ColumnHeaders.Add , , "Módulo", ListView5.Width / 5


    ListView6.ColumnHeaders.Add , , "ID", ListView6.Width / 14
    ListView6.ColumnHeaders.Add , , "Nome", ListView6.Width / 2.8
    ListView6.ColumnHeaders.Add , , "Descrição", ListView6.Width / 2.8
    ListView6.ColumnHeaders.Add , , "Ativo", ListView6.Width / 10000
    ListView6.ColumnHeaders.Add , , "Cor", ListView6.Width / 10
    ListView6.ColumnHeaders.Add , , "Coligada", ListView6.Width / 10000
    

    ListView7.ColumnHeaders.Clear
    ListView7.ColumnHeaders.Add , , "Descrição", ListView7.Width / 1.5
    ListView7.ColumnHeaders.Add , , "Data/Hora", ListView7.Width / 4


'    ListView2.ColumnHeaders.Clear
'    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 12
'    ListView2.ColumnHeaders.Add , , "Tipo", ListView2.Width / 10
'    ListView2.ColumnHeaders.Add , , "Ocorrência1", ListView2.Width / 8
'    ListView2.ColumnHeaders.Add , , "Ocorrência2", ListView2.Width / 8
'    ListView2.ColumnHeaders.Add , , "Pontos", ListView2.Width / 12
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "Código", ListView3.Width / 11
    ListView3.ColumnHeaders.Add , , "Empresa/Coligada", ListView3.Width / 1.5
    ListView3.ColumnHeaders.Add , , "Status", ListView3.Width / 11
    ListView3.ColumnHeaders.Add , , "Endereço", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Bairro", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Cidade", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "UF", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "CEP", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Email", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Site", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Telefone", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Fax", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "CNPJ", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "IE", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Caminho", ListView3.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView6.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview
    ListView7.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

'Private Sub LimpaControlesABS()
'    Dim X As Integer
'    cboABS(0).Text = ""
'    cboABS(1).Text = ""
'    cboABS(2).Text = ""
'    txtABS = ""
'    If ListView2.ListItems.Count > 0 Then
'        Label21.Caption = Format(GeraCodigo(ListView2), "00")
'    Else
'        Label21.Caption = Format(Val(Label21) + 1, "00")
'    End If
'End Sub

Private Sub LimpaControlesAvaliacao()
'    Dim X As Integer
'    txtCadParametro(2) = ""
'    optCadParametro(0).Value = True
'    If ListView1.ListItems.Count > 0 Then
'        Label14.Caption = Format(GeraCodigo(ListView1), "00")
'    Else
'        Label14.Caption = Format(Val(Label14) + 1, "00")
'    End If
'    SkinLabel18.Caption = "-"
End Sub

Private Sub LimpaControlesColigada()
    Dim X As Integer
    For X = 0 To txtDadosEmpresa.Count - 1
        txtDadosEmpresa(X) = ""
    Next
    If ListView3.ListItems.Count > 0 Then
        txtDadosEmpresa(11).Text = Format(GeraCodigo(ListView3), "00")
    Else
        txtDadosEmpresa(11).Text = Format(Val(txtDadosEmpresa(11).Text) + 1, "00")
    End If
    aicAlphaImage1.ClearImage
    Label53 = "-"
    txtDadosEmpresa(0).SetFocus
End Sub

Private Sub Compoe_ListviewConf()
On Error GoTo Err
    Dim rsAD As New ADODB.Recordset
    Dim sqlAD As String
    Dim rsABS As New ADODB.Recordset
    Dim sqlABS As String
    Dim rsColigadas As New ADODB.Recordset
    Dim sqlColigadas As String
    
    Dim ItemLst As ListItem
    Dim X As Integer
    
    RestauraLV ListView1, "CD"
    RestauraLV ListView2, "RNC"
    RestauraLV ListView4, "SI"
    RestauraLV ListView5, "SRM"

    ' Compoe Listview3
    sqlColigadas = "Select * from tbDadosEmpresa Order by codcoligada"
'    sqlColigadas = "Select * from tbDadosEmpresa where codcoligada = '" & vCodColigada & "' Order by codcoligada"
    rsColigadas.Open sqlColigadas, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsColigadas.EOF
        Set ItemLst = ListView3.ListItems.Add(, , Format(rsColigadas.Fields(13), "00"))
        ItemLst.SubItems(1) = "" & rsColigadas.Fields(0)
        ItemLst.SubItems(2) = "" & rsColigadas.Fields(14)
        
        If rsColigadas.Fields(14) = "N" Then 'Ativo
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(2) = ""
            ItemLst.ListSubItems.Item(2).ReportIcon = "OK"
        End If
        ItemLst.SubItems(3) = "" & rsColigadas.Fields(1)
        ItemLst.SubItems(4) = "" & rsColigadas.Fields(2)
        ItemLst.SubItems(5) = "" & rsColigadas.Fields(3)
        ItemLst.SubItems(6) = "" & rsColigadas.Fields(4)
        ItemLst.SubItems(7) = "" & rsColigadas.Fields(5)
        ItemLst.SubItems(8) = "" & rsColigadas.Fields(6)
        ItemLst.SubItems(9) = "" & rsColigadas.Fields(7)
        ItemLst.SubItems(10) = "" & rsColigadas.Fields(8)
        ItemLst.SubItems(11) = "" & rsColigadas.Fields(9)
        ItemLst.SubItems(12) = "" & rsColigadas.Fields(10)
        ItemLst.SubItems(13) = "" & rsColigadas.Fields(11)
        ItemLst.SubItems(14) = "" & rsColigadas.Fields(12)
        rsColigadas.MoveNext
        X = X + 1
    Wend
    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 0
    Me.ListView3.SortOrder = lvwDescending
    rsColigadas.Close
    Set rsColigadas = Nothing
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

'Private Sub MarcaDesmarca(LV As ListView)
'    Dim Y As Integer, X As Integer
'
'    Y = LV.ListItems.Count
'    For X = 1 To Y
'        LV.ListItems(X).Selected = True
'        If LV.ListItems.Item(X).Checked = True Then
'            LV.ListItems.Item(X).Checked = False
'        Else
'            LV.ListItems.Item(X).Checked = True
'        End If
'    Next
'End Sub

Private Function GeraCodigo(LV As Listview)
    Dim X As Integer
    X = 1
    LV.SortOrder = lvwDescending
    LV.ListItems.Item(X).Selected = True
    GeraCodigo = LV.ListItems.Item(X) + 1
    LV.SortOrder = lvwAscending
    Exit Function
End Function

'Private Sub IncluirAvaliacao()
'    Dim ItemLst As ListItem
'    Dim X As Integer, Y As Integer
'    Y = ListView1.ListItems.Count
'    If SkinLabel18 = "-" Then
'        mobjMsg.Abrir "Selecione um modelo para a ADP", Ok, critico, "Atenção"
'        Exit Sub
'    End If
'    If Y > 0 Then
'        For X = 1 To Y
'            ListView1.ListItems.Item(X).Selected = True
'            If ListView1.ListItems.Item(X) = Me.Label14.Caption Then
'                Label14.Caption = ListView1.ListItems.Item(X)
'                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadParametro(2).Text
'                If optCadParametro(0).Value = True Then
'                    ListView1.SelectedItem.ListSubItems.Item(2) = "Experiência"
'                Else
'                    ListView1.SelectedItem.ListSubItems.Item(2) = "Periódico"
'                End If
'                ListView1.SelectedItem.ListSubItems.Item(3) = SkinLabel18
'                Y = ListView1.ListItems.Count
'                Me.ListView1.Sorted = True
'                Me.ListView1.SortKey = 0
'                Me.ListView1.SortOrder = lvwAscending
''                LimpaControlesAvaliacao
'                Exit Sub
'            End If
'        Next
'        Set ItemLst = ListView1.ListItems.Add(, , Label14)
'        Y = ListView1.ListItems.Count
'    Else
'        Set ItemLst = ListView1.ListItems.Add(, , Label14)
'        Y = ListView1.ListItems.Count
'        Me.ListView1.Sorted = True
'        Me.ListView1.SortKey = 0
'        Me.ListView1.SortOrder = lvwDescending
'    End If
'    ItemLst.SubItems(1) = txtCadParametro(2).Text
'    If optCadParametro(0).Value = True Then
'        ItemLst.SubItems(2) = "Experiência"
'    Else
'        ItemLst.SubItems(2) = "Periódico"
'    End If
'    ItemLst.SubItems(3) = SkinLabel18
'    Me.ListView1.SortOrder = lvwAscending
'    txtCadParametro(2).SetFocus
'    'LimpaControlesAvaliacao
'End Sub

'Private Sub IncluirABS()
'    If ValidaABS = False Then Exit Sub
'    Dim ItemLst As ListItem
'    Dim X As Integer, Y As Integer
'    Y = ListView2.ListItems.Count
'    If Y > 0 Then
'        For X = 1 To Y
'            ListView2.ListItems.Item(X).Selected = True
'            If ListView2.ListItems.Item(X) = Me.Label21.Caption Then
'                Label21.Caption = ListView2.ListItems.Item(X)
'                ListView2.SelectedItem.ListSubItems.Item(1) = cboABS(0).Text
'                ListView2.SelectedItem.ListSubItems.Item(2) = cboABS(1).Text
'                ListView2.SelectedItem.ListSubItems.Item(3) = cboABS(2).Text
'                ListView2.SelectedItem.ListSubItems.Item(4) = txtABS.Text
'                Y = ListView2.ListItems.Count
'                Me.ListView2.Sorted = True
'                Me.ListView2.SortKey = 0
'                Me.ListView2.SortOrder = lvwAscending
''                LimpaControlesABS
'                Exit Sub
'            End If
'        Next
'        Set ItemLst = ListView2.ListItems.Add(, , Label21)
'        Y = ListView2.ListItems.Count
'    Else
'        Set ItemLst = ListView2.ListItems.Add(, , Label21)
'        Y = ListView2.ListItems.Count
'        Me.ListView2.Sorted = True
'        Me.ListView2.SortKey = 0
'        Me.ListView2.SortOrder = lvwDescending
'    End If
'    ItemLst.SubItems(1) = cboABS(0).Text
'    ItemLst.SubItems(2) = cboABS(1).Text
'    ItemLst.SubItems(3) = cboABS(2).Text
'    ItemLst.SubItems(4) = txtABS.Text
'    Me.ListView2.SortOrder = lvwAscending
'    cboABS(0).SetFocus
''    LimpaControlesABS
'End Sub

Private Sub IncluirColigada()
    If ValidaDadosColigada = False Then Exit Sub
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView3.ListItems.Count
    If txtDadosEmpresa(11) = "" Then
        If ListView3.ListItems.Count > 0 Then
            txtDadosEmpresa(11).Text = Format(GeraCodigo(ListView3), "00")
        Else
            txtDadosEmpresa(11).Text = Format(Val(txtDadosEmpresa(11).Text) + 1, "00")
        End If
    End If
    
    If Y > 0 Then
        For X = 1 To Y
            ListView3.ListItems.Item(X).Selected = True
            If ListView3.ListItems.Item(X) = Me.txtDadosEmpresa(11).Text Then
                Me.txtDadosEmpresa(11).Text = ListView3.ListItems.Item(X)
                ListView3.SelectedItem.ListSubItems.Item(1) = txtDadosEmpresa(0).Text
                ListView3.SelectedItem.ListSubItems.Item(2) = ""
                ListView3.SelectedItem.ListSubItems.Item(2).ReportIcon = "OK"
                ListView3.SelectedItem.ListSubItems.Item(3) = txtDadosEmpresa(1)
                ListView3.SelectedItem.ListSubItems.Item(4) = txtDadosEmpresa(2)
                ListView3.SelectedItem.ListSubItems.Item(5) = txtDadosEmpresa(3)
                ListView3.SelectedItem.ListSubItems.Item(6) = cboDadosEmpresa
                ListView3.SelectedItem.ListSubItems.Item(7) = txtDadosEmpresa(4)
                ListView3.SelectedItem.ListSubItems.Item(8) = txtDadosEmpresa(5)
                ListView3.SelectedItem.ListSubItems.Item(9) = txtDadosEmpresa(6)
                ListView3.SelectedItem.ListSubItems.Item(10) = txtDadosEmpresa(7)
                ListView3.SelectedItem.ListSubItems.Item(11) = txtDadosEmpresa(8)
                ListView3.SelectedItem.ListSubItems.Item(12) = txtDadosEmpresa(9)
                ListView3.SelectedItem.ListSubItems.Item(13) = txtDadosEmpresa(10)
                ListView3.SelectedItem.ListSubItems.Item(14) = Label53
                
                Y = ListView3.ListItems.Count
                Me.ListView3.Sorted = True
                Me.ListView3.SortKey = 0
                Me.ListView3.SortOrder = lvwAscending
                LimpaControlesColigada
                Exit Sub
            End If
        Next
        Set ItemLst = ListView3.ListItems.Add(, , txtDadosEmpresa(11).Text)
        Y = ListView3.ListItems.Count
    Else
        Set ItemLst = ListView3.ListItems.Add(, , txtDadosEmpresa(11).Text)
        Y = ListView3.ListItems.Count
        Me.ListView3.Sorted = True
        Me.ListView3.SortKey = 0
        Me.ListView3.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = txtDadosEmpresa(0).Text
    ItemLst.SubItems(2) = ""
    ItemLst.ListSubItems.Item(2).ReportIcon = "OK"
    ItemLst.SubItems(3) = txtDadosEmpresa(1)
    ItemLst.SubItems(4) = txtDadosEmpresa(2)
    ItemLst.SubItems(5) = txtDadosEmpresa(3)
    ItemLst.SubItems(6) = cboDadosEmpresa
    ItemLst.SubItems(7) = txtDadosEmpresa(4)
    ItemLst.SubItems(8) = txtDadosEmpresa(5)
    ItemLst.SubItems(9) = txtDadosEmpresa(6)
    ItemLst.SubItems(10) = txtDadosEmpresa(7)
    ItemLst.SubItems(11) = txtDadosEmpresa(8)
    ItemLst.SubItems(12) = txtDadosEmpresa(9)
    ItemLst.SubItems(13) = txtDadosEmpresa(10)
    ItemLst.SubItems(14) = Label53
    
    Me.ListView3.SortOrder = lvwAscending
    txtDadosEmpresa(0).SetFocus
    LimpaControlesColigada
End Sub

'Private Function ValidaABS()
'    ValidaABS = False
'    Dim X As Integer
'    For X = 0 To 1
'        If cboABS(X).Text = "" Then
'            mobjMsg.Abrir "Favor informar o campo " & Me.cboABS(X).Tag, Ok, informacao, "Atenção"
'            Me.cboABS(X).SetFocus
'            Exit Function
'        End If
'
'    Next
'    If txtABS.Text = "" Then
'        mobjMsg.Abrir "Favor informar o campo " & Me.txtABS.Tag, Ok, informacao, "Atenção"
'        Me.txtABS.SetFocus
'        Exit Function
'    End If
'    ValidaABS = True
'End Function

'Private Sub AlteraAvaliacao()
'    'Dim rsAlteraAV As New ADODB.Recordset
'    'Dim sqlAlteraAV As String
'
'    Dim Y As Integer, X As Integer
'    Y = ListView1.ListItems.Count
'    For X = 1 To Y
'        If ListView1.ListItems.Item(X).Selected = True Then
'            Exit For
'        End If
'    Next
'    Me.Label14.Caption = ListView1.ListItems.Item(X)
'    Me.txtCadParametro(2).Text = ListView1.SelectedItem.ListSubItems.Item(1)
'    If ListView1.SelectedItem.ListSubItems.Item(2) = "Experiência" Then
'        Me.optCadParametro(0).Value = True
'    Else
'        Me.optCadParametro(1).Value = True
'    End If
'
'    Me.SkinLabel18.Caption = ListView1.SelectedItem.ListSubItems.Item(3)
'
'    'sqlAlteraAV = "Select * from tbModeloADP where codmodelo = '" & Val(ListView1.ListItems.Item(X)) & "'"
'    'rsAlteraAV.Open sqlAlteraAV, cnBanco, adOpenKeyset, adLockReadOnly
'    'If Not rsAlteraAV.EOF Then
'    '    SkinLabel18 = rsAlteraAV.Fields(0)
'    'Else
'    '    SkinLabel18 = "-"
'    'End If
'    'rsAlteraAV.Close
'    'Set rsAlteraAV = Nothing
'End Sub

'Private Sub AlteraABS()
'    Dim Y As Integer, X As Integer
'    Y = ListView2.ListItems.Count
'    For X = 1 To Y
'        If ListView2.ListItems.Item(X).Selected = True Then
'            Exit For
'        End If
'    Next
'    Me.Label21.Caption = ListView2.ListItems.Item(X)
'    Me.cboABS(0).Text = ListView2.SelectedItem.ListSubItems.Item(1)
'    Me.cboABS(1).Text = ListView2.SelectedItem.ListSubItems.Item(2)
'    Me.cboABS(2).Text = ListView2.SelectedItem.ListSubItems.Item(3)
'    Me.txtABS.Text = ListView2.SelectedItem.ListSubItems.Item(4)
'End Sub

Private Sub AlteraColigada()
    Dim Y As Integer, X As Integer
    Y = ListView3.ListItems.Count
    For X = 1 To Y
        If ListView3.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtDadosEmpresa(11).Text = ListView3.ListItems.Item(X) 'Codigo
    Me.txtDadosEmpresa(0).Text = ListView3.SelectedItem.ListSubItems.Item(1) 'Razao social
    Me.txtDadosEmpresa(1).Text = ListView3.SelectedItem.ListSubItems.Item(3) 'Endereco
    Me.txtDadosEmpresa(2).Text = ListView3.SelectedItem.ListSubItems.Item(4) 'Bairro
    Me.txtDadosEmpresa(3).Text = ListView3.SelectedItem.ListSubItems.Item(5) 'Cidade
    Me.cboDadosEmpresa.Text = ListView3.SelectedItem.ListSubItems.Item(6) 'UF
    
    Me.txtDadosEmpresa(4).Text = ListView3.SelectedItem.ListSubItems.Item(7) 'CEP
    Me.txtDadosEmpresa(5).Text = ListView3.SelectedItem.ListSubItems.Item(8) 'Email
    Me.txtDadosEmpresa(6).Text = ListView3.SelectedItem.ListSubItems.Item(9) 'Site
    Me.txtDadosEmpresa(7).Text = ListView3.SelectedItem.ListSubItems.Item(10) 'Telefone
    Me.txtDadosEmpresa(8).Text = ListView3.SelectedItem.ListSubItems.Item(11) 'Fax
    Me.txtDadosEmpresa(9).Text = ListView3.SelectedItem.ListSubItems.Item(12) 'CNPJ
    Me.txtDadosEmpresa(10).Text = ListView3.SelectedItem.ListSubItems.Item(13) 'IE
    Me.Label53.Caption = ListView3.SelectedItem.ListSubItems.Item(14) 'Caminho da foto
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
End Sub

Private Sub CarregaParametros()
On Error GoTo TrataErro1
    Dim rsParametros As New ADODB.Recordset
    Dim sqlParametros As String
    Dim rsEmpresa As New ADODB.Recordset
    Dim sqlEmpresa As String
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String
    Dim rsIntegracao As New ADODB.Recordset
    Dim sqlIntegracao As String
    
    Dim rsRelogio As New ADODB.Recordset
    Dim sqlRelogio As String
    Dim rsFlexJr As New ADODB.Recordset
    Dim sqlFlexJr As String
    
    If Text1.Text = "" Then cmdCad(1).Enabled = False
    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsParametros.EOF Then
        If Not IsNull(rsParametros.Fields(0)) Then Text4 = rsParametros.Fields(0) 'Valor que irá incia relatórios de inspeção/expedição
        'txtCadParametro(1) = rsParametros.Fields(2) 'Aprovado com restrição
        If Not IsNull(rsParametros.Fields(11)) And rsParametros.Fields(11) <> 0 Then 'Afastamento
            'Check8.Value = 1
            'txtCadParametro(3) = rsParametros.Fields(11)
        Else
            'txtCadParametro(3).Enabled = False
            'Check9.Enabled = False
            'Check10.Enabled = False
        End If
        If rsParametros.Fields(1) = "S" Then 'Gera treinamento introdutorio
            'Check3.Value = 1
        Else
            'Check3.Value = 0
        End If
        
        If rsParametros.Fields(7) = "S" Then 'Avisos
            Check5.Value = 1
        Else
            Check5.Value = 0
        End If
        If rsParametros.Fields(10) = "S" Then 'Calcula experiência
            'Check7.Value = 1
        Else
            'Check7.Value = 0
        End If
        
        If rsParametros.Fields(8) = "S" Then 'Atualização automática
            Check6.Value = 1
            Text2.Text = rsParametros.Fields(9)
        Else
            Check6.Value = 0
        End If
        
        If rsParametros.Fields(5) = "S" Then 'Integração
            Check4.Value = 1
        Else
            Check4.Value = 0
        End If
        
        If rsParametros.Fields(4) = "S" Then 'Gerar treinamento obrigatorio
            'Check2.Value = 1
        Else
            'Check2.Value = 0
        End If
        
        If rsParametros.Fields(3) = "S" Then ' Ativa LOG
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        
        If Check8.Value = 1 Then
            'If rsParametros.Fields(12) = "S" Then 'Introdutorio após afastamento
            '    Check9.Value = 1
            'Else
            '    Check9.Value = 0
            'End If
            'If rsParametros.Fields(13) = "S" Then ' Obrigatorio após afastamento
            '    Check10.Value = 1
            'Else
            '    Check10.Value = 0
            'End If
        End If
        txtCadParametro(4) = rsParametros.Fields(14)
    End If
    rsParametros.Close
    Set rsParametros = Nothing
    
    sqlEmpresa = "Select * from tbDadosEmpresa where codcoligada = '" & vCodcoligada & "'"
    rsEmpresa.Open sqlEmpresa, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsEmpresa.EOF Then 'rsEmpresa.AddNew
        txtDadosEmpresa(11) = Format(rsEmpresa.Fields(13), "00")
        txtDadosEmpresa(0) = rsEmpresa.Fields(0)
        txtDadosEmpresa(1) = rsEmpresa.Fields(1)
        txtDadosEmpresa(2) = rsEmpresa.Fields(2)
        txtDadosEmpresa(3) = rsEmpresa.Fields(3)
        cboDadosEmpresa.Text = rsEmpresa.Fields(4)
        txtDadosEmpresa(4) = rsEmpresa.Fields(5)
        txtDadosEmpresa(5) = rsEmpresa.Fields(6)
        txtDadosEmpresa(6) = rsEmpresa.Fields(7)
        txtDadosEmpresa(7) = rsEmpresa.Fields(8)
        txtDadosEmpresa(8) = rsEmpresa.Fields(9)
        txtDadosEmpresa(9) = rsEmpresa.Fields(10)
        txtDadosEmpresa(10) = rsEmpresa.Fields(11)
    
        If rsEmpresa.Fields(12) <> "Null" Then
            'On Error GoTo TrataErro1
            Label53.Caption = rsEmpresa.Fields(12)
            aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
        End If
    End If
    
    sqlConfEmail = "Select * from tbConfEmail where codcoligada = '" & vCodcoligada & "'"
    rsConfEmail.Open sqlConfEmail, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsConfEmail.EOF Then 'rsEmpresa.AddNew
        txtEmail(0) = rsConfEmail.Fields(0)
        txtEmail(1) = rsConfEmail.Fields(1)
        txtEmail(2) = rsConfEmail.Fields(2)
    End If
    
'*********************
    sqlIntegracao = "Select * from tbIntegracao where codcoligada = '" & vCodcoligada & "'"
    rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsIntegracao.EOF Then
        If rsIntegracao.Fields(0) = 1 Then optIntegra(0).Value = True
        If rsIntegracao.Fields(1) = 1 Then chkIntegra(0).Value = True
        
        txtIntegra(0).Text = rsIntegracao.Fields(3)
        txtIntegra(1).Text = rsIntegracao.Fields(4)
        txtIntegra(2).Text = rsIntegracao.Fields(5)
        txtIntegra(3).Text = rsIntegracao.Fields(6)
    End If
'*********************
    
    sqlRelogio = "SELECT IDRELOGIO,IPRELOGIO,CPFRESPONSAVEL,PASSWORD,CAMINHO FROM TBCONFRELOGIO"
    rsRelogio.Open sqlRelogio, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsRelogio.EOF Then
        txtIntegra(4).Text = rsRelogio.Fields(0)
        txtIntegra(5).Text = rsRelogio.Fields(1)
        txtIntegra(6).Text = rsRelogio.Fields(2)
        txtIntegra(7).Text = rsRelogio.Fields(3)
        txtIntegra(10).Text = rsRelogio.Fields(4)
    End If
    
    sqlFlexJr = "SELECT USUARIO,PASSWORD,CAMINHO FROM TBCONFFLEXJR"
    rsFlexJr.Open sqlFlexJr, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsFlexJr.EOF Then
        Text5.Text = rsFlexJr.Fields(2)
        txtIntegra(8).Text = rsFlexJr.Fields(0)
        txtIntegra(9).Text = rsFlexJr.Fields(1)
    End If
    
    rsConfEmail.Close
    Set rsConfEmail = Nothing
    rsEmpresa.Close
    Set rsEmpresa = Nothing
    rsIntegracao.Close
    Set rsIntegracao = Nothing
    
    rsRelogio.Close
    Set rsRelogio = Nothing
    rsFlexJr.Close
    Set rsFlexJr = Nothing
    
    Exit Sub
TrataErro1:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    End If
    'Label59.Visible = True
    'Resume Next
End Sub

Private Sub GravaParametros()
On Error GoTo Err
    If ValidaCampo = False Then Exit Sub
    If Check6.Value = 1 And Text2 = "" Then
        mobjMsg.Abrir "Informe o caminho do executável: AtualizaZEUSH.exe", Ok, critico, "Atenção"
    End If
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsParametros As New ADODB.Recordset
    Dim sqlParametros As String
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String
    Dim rsConfAD As New ADODB.Recordset
    Dim sqlConfAD As String
    Dim rsConfABS As New ADODB.Recordset
    Dim sqlConfABS As String
    Dim rsEmpresa As New ADODB.Recordset
    Dim sqlEmpresa As String
    Dim rsIntegracao As New ADODB.Recordset
    Dim sqlIntegracao As String
    
    Dim rsRelogio As New ADODB.Recordset
    Dim sqlRelogio As String
    Dim rsFlexJr As New ADODB.Recordset
    Dim sqlFlexJr As String
    
10  cnBanco.BeginTrans

    
    sqlDeletar = "Delete from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco

    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockOptimistic
    rsParametros.AddNew
    If Text4 <> "" Then
        Dim rsIDRels As New ADODB.Recordset
        Dim sqlIDRels As String
        sqlIDRels = "Select * from tbrelinspexp"
        rsIDRels.Open sqlIDRels, cnBanco, adOpenKeyset, adLockReadOnly
        'If rsIDRels.RecordCount = 0 Then
            rsParametros.Fields(0) = Val(Text4)
        'Else
        '    rsParametros.Fields(0) = 0
        'End If
        rsIDRels.Close
        Set rsIDRels = Nothing
    End If
    rsParametros.Fields(2) = 0
    If Check8.Value = 1 Then
        rsParametros.Fields(11) = ""
    Else
        rsParametros.Fields(11) = ""
    End If
    
    If Check7.Value = 1 Then
        rsParametros.Fields(10) = "S"
    Else
        rsParametros.Fields(10) = "N"
    End If
    
    If Check6.Value = 1 Then
        rsParametros.Fields(8) = "S"
    Else
        rsParametros.Fields(8) = "N"
    End If
    If Check6.Value = 1 Then
        rsParametros.Fields(9) = Text2
        vCaminhoAtu = Text2
    Else
        rsParametros.Fields(9) = ""
    End If
    
    If Check5.Value = 1 Then
        rsParametros.Fields(7) = "S"
    Else
        rsParametros.Fields(7) = "N"
    End If
    
    If Check3.Value = 1 Then
        rsParametros.Fields(1) = "S"
    Else
        rsParametros.Fields(1) = "N"
    End If
    If Check2.Value = 1 Then
        rsParametros.Fields(4) = "S"
    Else
        rsParametros.Fields(4) = "N"
    End If
    If Check1.Value = 1 Then
        rsParametros.Fields(3) = "S"
    Else
        rsParametros.Fields(3) = "N"
    End If
    If Check8.Value = 1 Then
        If Check9.Value = 1 Then
            rsParametros.Fields(12) = "S"
        Else
            rsParametros.Fields(12) = "N"
        End If
'        If Check10.Value = 1 Then
'            rsParametros.Fields(13) = "S"
'        Else
'            rsParametros.Fields(13) = "N"
'        End If
    Else
        rsParametros.Fields(12) = "N"
        rsParametros.Fields(13) = "N"
    End If
    If txtCadParametro(4) = "" Or txtCadParametro(4) = 0 Then
        rsParametros.Fields(14) = 1
        vOpenTabs = 1
    Else
        rsParametros.Fields(14) = txtCadParametro(4)
        vOpenTabs = txtCadParametro(4)
    End If
    
    If Combo5.ListIndex = 0 Then
        rsParametros.Fields(15) = 1
    ElseIf Combo5.ListIndex = 1 Then
        rsParametros.Fields(15) = 2
    ElseIf Combo5.ListIndex = 2 Then
        rsParametros.Fields(15) = 3
    ElseIf Combo5.ListIndex = 3 Then
        rsParametros.Fields(15) = 4
    ElseIf Combo5.ListIndex = 4 Then
        rsParametros.Fields(15) = 5
    ElseIf Combo5.ListIndex = 5 Then
        rsParametros.Fields(15) = 6
    End If
    
    rsParametros.Fields(6) = vCodcoligada 'Codigo da coligada
    If Check4.Value = 1 Then
'*************************
        'GRAVA DADOS DE INTEGRAÇÃO
        
        vServerTotvs = txtIntegra(0).Text
        vBancoTotvs = txtIntegra(1).Text
        vUsuBancoTovs = txtIntegra(2).Text
        vSenhaBancoTotvs = txtIntegra(3).Text
    
        If testaParametros = False Then
            Check4.Value = 0
            mobjMsg.Abrir "Os dados informados para conexão não estão corretos", Ok, critico, "Conexão TOTVS"
            rsParametros.Fields(5) = "N"
        Else
            rsParametros.Fields(5) = "S"
            vIntegra = "S"
        End If
            
        sqlIntegracao = "Select * from tbIntegracao Where codcoligada = '" & vCodcoligada & "'"
        rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockOptimistic
        If rsIntegracao.EOF Then rsIntegracao.AddNew
        ' 1 = SQL Server / 2 = Oracle
        If optIntegra(0).Value = True Then rsIntegracao.Fields(0) = 1 Else rsIntegracao.Fields(0) = 2
        If chkIntegra(0).Value = True Then rsIntegracao.Fields(1) = 1
        'If chkIntegra(1).Value = True Then rsIntegracao.Fields(1) = 2
        'If chkIntegra(2).Value = True Then rsIntegracao.Fields(1) = 3
        rsIntegracao.Fields(2) = "1.1"
        rsIntegracao.Fields(3) = txtIntegra(0).Text
        rsIntegracao.Fields(4) = txtIntegra(1).Text
        rsIntegracao.Fields(5) = txtIntegra(2).Text
        rsIntegracao.Fields(6) = txtIntegra(3).Text
        rsIntegracao.Fields(7) = vCodcoligada 'Codigo da coligada
        rsIntegracao.Update
        rsIntegracao.Close
        Set rsIntegracao = Nothing
    Else
        rsParametros.Fields(5) = "N"
        vIntegra = "N"
    End If
'*************************
    rsParametros.Update
    rsParametros.Close
    IniciaRelsEm = 0
    vAprovadoRest = 0
    If Check7.Value = 1 Then
        vCalcExp = "S"
    Else
        vCalcExp = "N"
    End If
    
    If Check5.Value = 1 Then
        vAvisos = "S"
    Else
        vAvisos = "N"
    End If
    
    If Check3.Value = 1 Then
        GeraIntr = "S"
    Else
        GeraIntr = "N"
    End If
    
    If Check2.Value = 1 Then
        GeraObri = "S"
    Else
        GeraObri = "N"
    End If
    
    If Check1.Value = 1 Then
        GeraLog = "S"
    Else
        GeraLog = "N"
    End If
    
    If Check8.Value = 1 Then
        vAfastDias = ""
        If Check9.Value = 1 Then
            vAfastTreiInt = "S"
        Else
            vAfastTreiInt = "N"
        End If
'        If Check10.Value = 1 Then
'            vAfastTreiObr = "S"
'        Else
'            vAfastTreiObr = "N"
'        End If
    End If
    
    Set rsParametros = Nothing
    
    sqlEmpresa = "Delete from tbDadosEmpresa"
    rsEmpresa.Open sqlEmpresa, cnBanco
    
    sqlEmpresa = "Select * from tbDadosEmpresa"
    rsEmpresa.Open sqlEmpresa, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To ListView3.ListItems.Count
        ListView3.ListItems.Item(X).Selected = True
        rsEmpresa.AddNew
        rsEmpresa.Fields(0) = ListView3.SelectedItem.ListSubItems.Item(1) ' Nome
        rsEmpresa.Fields(1) = ListView3.SelectedItem.ListSubItems.Item(3) 'Endereco
        rsEmpresa.Fields(2) = ListView3.SelectedItem.ListSubItems.Item(4) 'Bairro
        rsEmpresa.Fields(3) = ListView3.SelectedItem.ListSubItems.Item(5) 'Cidade
        rsEmpresa.Fields(4) = ListView3.SelectedItem.ListSubItems.Item(6) 'UF
        rsEmpresa.Fields(5) = ListView3.SelectedItem.ListSubItems.Item(7) 'CEP
        rsEmpresa.Fields(6) = ListView3.SelectedItem.ListSubItems.Item(8) 'Email
        rsEmpresa.Fields(7) = ListView3.SelectedItem.ListSubItems.Item(9) 'Site
        rsEmpresa.Fields(8) = ListView3.SelectedItem.ListSubItems.Item(10) 'Telefone
        rsEmpresa.Fields(9) = ListView3.SelectedItem.ListSubItems.Item(11) 'Fax
        rsEmpresa.Fields(10) = ListView3.SelectedItem.ListSubItems.Item(12) 'CNPJ
        rsEmpresa.Fields(11) = ListView3.SelectedItem.ListSubItems.Item(13) 'IE
        rsEmpresa.Fields(12) = ListView3.SelectedItem.ListSubItems.Item(14) 'Logo
        rsEmpresa.Fields(13) = ListView3.ListItems.Item(X) ' codigo da coligada
        rsEmpresa.Fields(14) = ListView3.SelectedItem.ListSubItems.Item(2) 'Status
    Next
    
    

    'Dim rsFlexJr As New ADODB.Recordset
    'Dim sqlFlexJr As String
    
    sqlDeletar = "Delete from TBCONFRELOGIO"
    rsDeletar.Open sqlDeletar, cnBanco
    sqlRelogio = "Insert into TBCONFRELOGIO(IDRELOGIO,IPRELOGIO,CPFRESPONSAVEL,PASSWORD, CAMINHO) Values('" & txtIntegra(4) & "','" & txtIntegra(5) & "','" & txtIntegra(6) & "','" & txtIntegra(7) & "','" & txtIntegra(10) & "')"
    rsRelogio.Open sqlRelogio, cnBanco
    
    
    sqlDeletar = "Delete from TBCONFFLEXJR"
    rsDeletar.Open sqlDeletar, cnBanco
    sqlFlexJr = "Insert into TBCONFFLEXJR(USUARIO, PASSWORD, CAMINHO) Values('" & txtIntegra(8) & "','" & txtIntegra(9) & "','" & Text5 & "')"
    rsFlexJr.Open sqlFlexJr, cnBanco
    
    
    sqlDeletar = "Delete from tbConfEmail where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    sqlConfEmail = "Insert into tbConfEmail(smtp,usuario,senha,codcoligada) Values('" & txtEmail(0) & "','" & txtEmail(1) & "','" & txtEmail(2) & "','" & vCodcoligada & "')"
    rsConfEmail.Open sqlConfEmail, cnBanco
    
    vSMTP = txtEmail(0)
    vUsuEmail = txtEmail(1)
    vSenhaEmail = txtEmail(2)
    
    rsEmpresa.Update
    rsEmpresa.Close
    Set rsEmpresa = Nothing
    
    '-GRAVA EMAIL CD - COMUNICACAO DE DESVIO-----------
    GravaLV ListView1, "CD", sEmailCD
    '-GRAVA EMAIL RNC - REGISTRO DE NÃO CONFORMIDADE-----------
    GravaLV ListView2, "RNC", sEmailRNC
    '-GRAVA EMAIL SI - SOLICITAÇÃO DE INSPEÇÃO-----------
    GravaLV ListView4, "SI", sEmailSI
    '-GRAVA EMAIL SRM - SOLICITAÇÃO DE RETIRADA DE MATERIAL-----------
    GravaLV ListView5, "SRM", sEmailSRM
    
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ZEUS\" & "sLogoEmpresa", Label53 'Logo da empresa
    Set Reg = Nothing
    
    
    cnBanco.CommitTrans
    
    If ListView6.ListItems.Count > 0 Then
        limpaQualquerDado
        ordenaLVArray ListView6, "0", "1", "2", "3", "5", "4", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "tbTipoFCE", "", "I", txtCadTipo(0)
    End If
    
    If ListView7.ListItems.Count > 0 Then
        limpaQualquerDado
        ordenaLVArray ListView7, "0", "0", "1", "", "", "", "", "", "", "", "", "", "", "", "", ""
        GravaDadosLV "TBCONTROLEATIVIDADES", "", "I", Text6
        
        Me.ListView7.Sorted = True
        Me.ListView7.SortKey = 1
        Me.ListView7.SortOrder = lvwDescending
    End If
    
    carregaCoresTipoFCE
    mobjMsg.Abrir "Os dados de configuração do sistema foram salvos com sucesso", Ok, informacao, "ZEUS"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Sub
    End If
End Sub

Private Sub GravaLV(vLV As Listview, vModulo As String, vRecEmails As String)
On Error GoTo Err
'--------------------------------------------------------------------------------
'-GRAVA OS EMAILS Q SERÃO ENVIADOS PELO SISTEMA DE ACORDO COM O MODULO-----------
    Dim rsEmailSystem As New ADODB.Recordset
    Dim sqlEmailSystem As String
    Dim X As Integer
    sqlEmailSystem = "Delete from tbEnvioEmail where modulo = '" & vModulo & "'"
    rsEmailSystem.Open sqlEmailSystem, cnBanco
    
    sqlEmailSystem = "Select * from tbEnvioEmail"
    rsEmailSystem.Open sqlEmailSystem, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To vLV.ListItems.Count
        vLV.ListItems.Item(X).Selected = True
        rsEmailSystem.AddNew
        rsEmailSystem.Fields(0) = vModulo
        rsEmailSystem.Fields(1) = vLV.ListItems.Item(X) ' E-mail
        If vRecEmails = "" Then
            vRecEmails = vLV.ListItems.Item(X)
        Else
            vRecEmails = vRecEmails & ";" & vLV.ListItems.Item(X)
        End If
    Next
    rsEmailSystem.Update
    rsEmailSystem.Close
    Set rsEmailSystem = Nothing
'--------------------------------------------------
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

Private Sub RestauraLV(vLV As Listview, vModulo As String)
On Error GoTo Err
    Dim rsEmailSystem As New ADODB.Recordset
    Dim sqlEmailSystem As String
    
    sqlEmailSystem = "Select * from tbEnvioEmail where modulo = '" & vModulo & "'"
    rsEmailSystem.Open sqlEmailSystem, cnBanco, adOpenKeyset, adLockReadOnly
    X = 0
    While Not rsEmailSystem.EOF
        Set ItemLst = vLV.ListItems.Add(, , rsEmailSystem.Fields(1))
        ItemLst.SubItems(1) = rsEmailSystem.Fields(0)
        rsEmailSystem.MoveNext
        X = X + 1
    Wend
    vLV.Sorted = True
    vLV.SortKey = 0
    vLV.SortOrder = lvwDescending
    rsEmailSystem.Close
    Set rsEmailSystem = Nothing
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

Private Function testaParametros()
On Error GoTo Err
    testaParametros = False
    ConexaoTotvs
    If vIntegra = "S" Then testaParametros = True Else testaParametros = False
    cnBancoTotvs.Close
    Set cnBancoTotvs = Nothing
    Exit Function
Err:
    testaParametros = False
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If ListView3.ListItems.Count = 0 Then
        mobjMsg.Abrir "Nenhuma empresa/coligada cadastrada. Favor informar os dados da empresa/coligada", Ok, informacao, "ZEUS"
        SSTab1.Tab = 2
        Exit Function
    End If
    
    'If txtCadParametro(0) = "" Then
    '    mobjMsg.Abrir "Favor informar o campo " & Me.txtCadParametro(0).Tag, Ok, informacao, "Atenção"
    '    Me.txtCadParametro(0).SetFocus
    '    Exit Function
    'End If
    'If txtCadParametro(1) = "" Then
    '    mobjMsg.Abrir "Favor informar o campo " & Me.txtCadParametro(1).Tag, Ok, informacao, "Atenção"
    '    Me.txtCadParametro(1).SetFocus
    '    Exit Function
    'End If
    ValidaCampo = True
End Function

Private Function ValidaCampoTree()
    ValidaCampoTree = False
    If Check8.Value = 1 And Combo2.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Combo2.Tag, Ok, informacao, "Atenção"
        Combo2.SetFocus
        Exit Function
    End If
    If Check2.Value = 1 And Combo1.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Combo1.Tag, Ok, informacao, "Atenção"
        Combo1.SetFocus
        Exit Function
    End If
    If Check3.Value = 1 And Combo3.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Combo3.Tag, Ok, informacao, "Atenção"
        Combo3.SetFocus
        Exit Function
    End If
    If Check7.Value = 1 And Combo4.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Combo4.Tag, Ok, informacao, "Atenção"
        Combo4.SetFocus
        Exit Function
    End If
    If Check9.Value = 1 And Text7.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Text7.Tag, Ok, informacao, "Atenção"
        Text7.SetFocus
        Exit Function
    End If
    If Text3.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Text3.Tag, Ok, informacao, "Atenção"
        Text3.SetFocus
        Exit Function
    End If
    ValidaCampoTree = True
End Function

Private Function ValidaDadosColigada()
    ValidaDadosColigada = False
    If ListView3.ListItems.Count = 0 Then
        If txtDadosEmpresa(0) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtDadosEmpresa(0).Tag, Ok, informacao, "Atenção"
            Me.txtDadosEmpresa(0).SetFocus
            Exit Function
        End If
    End If
    ValidaDadosColigada = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub configControles()
    'If vInc = "N" Then
    '    cmdCadastro(12).UseGreyscale = True
    '    cmdCadastro(12).DragMode = 1
    '    cmdCadastro(12).SpecialEffect = cbEngraved
    'End If
    'If vExc = "N" Then
    '    cmdCadastro(13).UseGreyscale = True
    '    cmdCadastro(13).DragMode = 1
    '    cmdCadastro(13).SpecialEffect = cbEngraved
    'End If
    'If vSal = "N" Then
    '    cmdCadastro(0).UseGreyscale = True
    '    cmdCadastro(0).DragMode = 1
    '    cmdCadastro(0).SpecialEffect = cbEngraved
    'End If
End Sub

Private Sub ListView3_DblClick()
    AlteraColigada
    SSTab4.Tab = 0
End Sub

Private Sub ListView6_DblClick()
    AlteraLV ListView6, txtCadTipo(0), txtCadTipo(1), txtCadTipo(2), vPonte1, txtCorTipo, vPonte2, txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(1), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0), txtCadTipo(0)
    If vPonte1 = "S" Then
        Check10.Value = 1 ' Status do tipo de FCE
    Else
        Check10.Value = 0 ' Status do tipo de FCE
    End If
    txtCorTipo.BackColor = txtCorTipo.Text
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
     If Text1.Text <> "" Then cmdCadastro(14).Enabled = True Else cmdCadastro(14).Enabled = False
End Sub

Private Sub Text1_LostFocus()
     If Text1.Text <> "" Then cmdCadastro(14).Enabled = True Else cmdCadastro(14).Enabled = False
End Sub

'Private Sub ListView1_DblClick()
'    AlteraAvaliacao
'End Sub

'Private Sub ListView2_DblClick()
'    AlteraABS
'End Sub

Private Sub importaColaboradores()
    On Error Resume Next
    Dim X As Integer
    Dim F As Long
    Dim Linhas As Variant
    Dim i As Long
    Dim Tmp As String
    F = FreeFile
    Open Text1.Text For Input As #F

    Tmp = Input(LOF(F), F)
    Close #F

    Linhas = Split(Tmp, Chr(10))
    For i = 0 To UBound(Linhas)
        var = Split(Linhas(i), ";")
        For X = 0 To 17
            colheDados(X) = var(X)
        Next
        If ValidaDados = False Then
            mobjMsg.Abrir "Erro na linha: " & i + 1, Ok, critico, "Atenção"
            Exit Sub
        End If
        insertDados
    Next
    mobjMsg.Abrir "Dados importados com sucesso!", Ok, informacao, "ZEUS"
End Sub

Private Function ValidaDados()
    ValidaDados = False
    Dim Y As Integer
    For Y = 0 To 3
        If colheDados(Y) = "" Then
            mobjMsg.Abrir "Erro de consistência na fonte de dados", Ok, informacao, "ZEUS"
            Exit Function
        End If
    Next
    ValidaDados = True
End Function

Private Sub insertDados()
On Error GoTo Err
    Dim rsImportaColabs As New ADODB.Recordset
    Dim SqlImportaColabs As String
    
    SqlImportaColabs = "Select a.* from tbcolaboradores as a where a.codcoligada = '" & vCodcoligada & "' and a.cpf= '" & colheDados(0) & "' and codcolaborador = '" & colheDados(1) & "'"
    rsImportaColabs.Open SqlImportaColabs, cnBanco, adOpenKeyset, adLockReadOnly
    If rsImportaColabs.RecordCount = 0 Then
        rsImportaColabs.Close
        Set rsImportaColabs = Nothing

        SqlImportaColabs = "Insert into tbColaboradores(cpf,codcolaborador,datacadastro,nomecolaborador,datanascimento,sexo,estadocivil,nacionalidade,naturalidade,ufnaturalidade,email,ctpsnumero,ctpsserie,cnhnumero,cnhtipo,ativo,telefone,celular,tipo,codcoligada) Values('" & colheDados(0) & "','" & colheDados(1) & "','" & colheDados(2) & "','" & colheDados(3) & "','" & colheDados(4) & "','" & colheDados(5) & _
                           "','" & colheDados(6) & "','" & colheDados(7) & "','" & colheDados(8) & "','" & colheDados(9) & "','" & colheDados(10) & "','" & colheDados(11) & "','" & colheDados(12) & "','" & colheDados(13) & "','" & colheDados(14) & "','" & colheDados(15) & "','" & colheDados(16) & "','" & colheDados(17) & "','Colaborador','" & vCodcoligada & "')"
        rsImportaColabs.Open SqlImportaColabs, cnBanco
    Else
        rsImportaColabs.Close
        Set rsImportaColabs = Nothing
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

Private Sub Text4_KeyPress(KeyAscii As Integer)
    'aceitar somente números e "Back Space", "Enter", "virgula"
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TreeView1_Click()
    AlteraTreeview
    CompoeDados
End Sub

Private Sub TreeView1_DblClick()
    AlteraTreeview
    CompoeDados
End Sub

Private Function GeraCodigoMenu()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from tbMenuConf order by id Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodigoMenu = rsGeraCodigo.Fields(4) + 1
    Else
        GeraCodigoMenu = 1
    End If
    SkinLabel13 = GeraCodigoMenu
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

Private Sub txtCadParametro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 4
            'aceitar somente números e "Back Space", "Enter", "virgula"
            If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
                KeyAscii = 0
            End If
    End Select
End Sub
