VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmLM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LM - Lista de Materiais"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17565
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   17565
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Lista de Materiais"
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   6
      Left            =   1320
      TabIndex        =   58
      Tag             =   "Gravar e Sair"
      ToolTipText     =   "Gravar e Sair"
      Top             =   8880
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
      MICON           =   "frmLM.frx":0CCA
      PICN            =   "frmLM.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   5
      Left            =   720
      TabIndex        =   57
      Tag             =   "Exporta para o Excel"
      ToolTipText     =   "Exporta para o Excel"
      Top             =   8880
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
      MICON           =   "frmLM.frx":19C0
      PICN            =   "frmLM.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame20 
      Caption         =   "Clonar Desenho "
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
      Left            =   9960
      TabIndex        =   145
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   49
         Left            =   1200
         TabIndex        =   160
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   48
         Left            =   4200
         TabIndex        =   159
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   47
         Left            =   3840
         TabIndex        =   156
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   46
         Left            =   3480
         TabIndex        =   155
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   45
         Left            =   3120
         TabIndex        =   154
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   43
         Left            =   840
         TabIndex        =   153
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   42
         Left            =   480
         TabIndex        =   152
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   330
         Index           =   41
         Left            =   120
         TabIndex        =   151
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   3120
         TabIndex        =   146
         Tag             =   "Desenho"
         ToolTipText     =   "Desenho"
         Top             =   480
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "frmLM.frx":26B6
         TabIndex        =   158
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtcadastro 
         Enabled         =   0   'False
         Height          =   285
         Index           =   39
         Left            =   120
         TabIndex        =   147
         Tag             =   "Desenho"
         ToolTipText     =   "Desenho"
         Top             =   480
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmLM.frx":2718
         TabIndex        =   157
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clonar"
         Height          =   375
         Left            =   6120
         TabIndex        =   150
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   5520
         TabIndex        =   149
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   148
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da LM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   59
      Top             =   120
      Width           =   9735
      Begin VB.ComboBox cboCadastro 
         Height          =   345
         Index           =   0
         Left            =   7920
         TabIndex        =   161
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   345
         Left            =   3360
         TabIndex        =   0
         Top             =   480
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "frmLM.frx":2776
         TabIndex        =   86
         Top             =   240
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label32 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmLM.frx":27E2
         TabIndex        =   66
         Top             =   480
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel Label2 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "frmLM.frx":2840
         TabIndex        =   67
         Top             =   480
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1800
         TabIndex        =   65
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
         Format          =   162267137
         CurrentDate     =   40449
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "frmLM.frx":289C
         TabIndex        =   85
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "frmLM.frx":2910
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmLM.frx":2974
         TabIndex        =   83
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "frmLM.frx":29DA
         TabIndex        =   162
         Top             =   240
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   60
      Top             =   1200
      Width           =   17355
      _ExtentX        =   30612
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmLM.frx":2A44
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(2)=   "Frame11"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Desenhos de Conjunto"
      TabPicture(1)   =   "frmLM.frx":2A60
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Lista de Materiais"
      TabPicture(2)   =   "frmLM.frx":2A7C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame19"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).Control(2)=   "Frame9(0)"
      Tab(2).Control(3)=   "Frame14"
      Tab(2).Control(4)=   "txtLvw"
      Tab(2).Control(5)=   "Frame17"
      Tab(2).Control(6)=   "Frame10"
      Tab(2).Control(7)=   "Frame3"
      Tab(2).Control(8)=   "Frame7"
      Tab(2).Control(9)=   "chamCad(7)"
      Tab(2).Control(10)=   "chamCad(1)"
      Tab(2).Control(11)=   "chamCad(0)"
      Tab(2).Control(12)=   "ScriptControl1"
      Tab(2).Control(13)=   "ListView2"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Resumo"
      TabPicture(3)   =   "frmLM.frx":2A98
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lbltotl"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lbltotpm"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label45"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label44"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "ListView3"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame19 
         Caption         =   "Identificador do Conjunto"
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
         Left            =   -69480
         TabIndex        =   141
         Top             =   3960
         Width           =   2535
         Begin VB.ComboBox Combo1 
            Height          =   345
            Left            =   120
            TabIndex        =   142
            Text            =   "-"
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Desenhos de Conjunto "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   -62760
         TabIndex        =   130
         Top             =   480
         Width           =   4935
         Begin MSComctlLib.ListView ListView5 
            Height          =   3855
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   6800
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
      Begin VB.Frame Frame8 
         Caption         =   "Desenho de Conjunto "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74880
         TabIndex        =   127
         Top             =   480
         Width           =   17055
         Begin ZEUS.chameleonButton chamCad 
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   140
            Top             =   6240
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
            MICON           =   "frmLM.frx":2AB4
            PICN            =   "frmLM.frx":2AD0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame18 
            Caption         =   "Desenho do conjunto"
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
            Left            =   2160
            TabIndex        =   132
            Top             =   240
            Width           =   10215
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   22
               Left            =   1920
               TabIndex        =   22
               Tag             =   "Desenho"
               ToolTipText     =   "Desenho"
               Top             =   480
               Width           =   2535
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   32
               Left            =   4560
               TabIndex        =   23
               Tag             =   "Revisão"
               ToolTipText     =   "Revisão"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               Height          =   285
               Index           =   33
               Left            =   120
               TabIndex        =   21
               Tag             =   "Descrição do desenho"
               ToolTipText     =   "Descrição do desenho"
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton Command2 
               Caption         =   "..."
               Height          =   255
               Left            =   1440
               TabIndex        =   133
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   5280
               TabIndex        =   24
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   35
               Left            =   6600
               TabIndex        =   25
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   36
               Left            =   7680
               TabIndex        =   26
               Top             =   480
               Width           =   2415
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
               Height          =   255
               Left            =   7680
               OleObjectBlob   =   "frmLM.frx":37AA
               TabIndex        =   134
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
               Height          =   255
               Left            =   6600
               OleObjectBlob   =   "frmLM.frx":3812
               TabIndex        =   135
               Top             =   240
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
               Height          =   255
               Index           =   1
               Left            =   5280
               OleObjectBlob   =   "frmLM.frx":3880
               TabIndex        =   136
               Top             =   240
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
               Height          =   255
               Index           =   1
               Left            =   120
               OleObjectBlob   =   "frmLM.frx":38E8
               TabIndex        =   137
               Top             =   240
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
               Height          =   375
               Index           =   1
               Left            =   4560
               OleObjectBlob   =   "frmLM.frx":3956
               TabIndex        =   138
               Top             =   240
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
               Height          =   255
               Index           =   1
               Left            =   1920
               OleObjectBlob   =   "frmLM.frx":39BE
               TabIndex        =   139
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Conjunto"
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
            TabIndex        =   131
            Top             =   240
            Width           =   1935
            Begin VB.TextBox txtcadastro 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   44
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   855
            End
            Begin ZEUS.chameleonButton chameleonButton6 
               Height          =   615
               Left            =   1200
               TabIndex        =   20
               Tag             =   "Novo Conjunto"
               ToolTipText     =   "Novo Conjunto"
               Top             =   240
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
               MICON           =   "frmLM.frx":3A26
               PICN            =   "frmLM.frx":3A42
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
            Height          =   615
            Left            =   2160
            TabIndex        =   128
            Top             =   1320
            Width           =   1695
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmLM.frx":471C
               TabIndex        =   129
               Top             =   240
               Width           =   1455
            End
         End
         Begin MSComctlLib.ListView ListView4 
            Height          =   3975
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   16815
            _ExtentX        =   29660
            _ExtentY        =   7011
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
         Begin ZEUS.chameleonButton chameleonButton2 
            Height          =   615
            Left            =   720
            TabIndex        =   28
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
            MICON           =   "frmLM.frx":4778
            PICN            =   "frmLM.frx":4794
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ZEUS.chameleonButton chameleonButton1 
            Height          =   615
            Left            =   120
            TabIndex        =   27
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
            MICON           =   "frmLM.frx":546E
            PICN            =   "frmLM.frx":548A
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
         Height          =   975
         Index           =   0
         Left            =   -74880
         TabIndex        =   123
         Top             =   2880
         Width           =   5295
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   24
            Left            =   120
            TabIndex        =   47
            Tag             =   "Dimensão"
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Pintura:"
            Height          =   255
            Left            =   4200
            TabIndex        =   126
            Top             =   240
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   15
            Left            =   4200
            TabIndex        =   49
            Tag             =   "Parâmetro para cálculo da área de pintura"
            ToolTipText     =   "Parâmetro para cálculo da área de pintura"
            Top             =   480
            Width           =   495
         End
         Begin VB.OptionButton optCadastro 
            Caption         =   "Dimensão:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optCadastro 
            Caption         =   "Peso:"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   124
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   23
            Left            =   2130
            TabIndex        =   48
            Tag             =   "Peso"
            Top             =   495
            Width           =   1935
         End
      End
      Begin VB.Frame Frame14 
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
         Height          =   615
         Left            =   -72840
         TabIndex        =   120
         Top             =   3900
         Width           =   1695
         Begin ACTIVESKINLibCtl.SkinLabel Label36 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6164
            TabIndex        =   121
            Top             =   240
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label37 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmLM.frx":61CC
            TabIndex        =   122
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.TextBox txtLvw 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70320
         TabIndex        =   119
         Top             =   4320
         Width           =   855
      End
      Begin VB.Frame Frame17 
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
         Height          =   975
         Left            =   -69480
         TabIndex        =   114
         Top             =   2880
         Width           =   3255
         Begin ACTIVESKINLibCtl.SkinLabel lblTotPint 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "frmLM.frx":6232
            TabIndex        =   115
            Top             =   480
            Width           =   1410
         End
         Begin ACTIVESKINLibCtl.SkinLabel lblTotal 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":628C
            TabIndex        =   116
            Top             =   480
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":62E6
            TabIndex        =   117
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "frmLM.frx":6354
            TabIndex        =   118
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame10 
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
         Height          =   975
         Left            =   -66120
         TabIndex        =   109
         Top             =   2880
         Width           =   2895
         Begin ACTIVESKINLibCtl.SkinLabel Label39 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmLM.frx":63CC
            TabIndex        =   110
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel Label38 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6426
            TabIndex        =   111
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
            Height          =   255
            Left            =   1440
            OleObjectBlob   =   "frmLM.frx":6480
            TabIndex        =   112
            Top             =   240
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":64F4
            TabIndex        =   113
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informações do Desenho "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74880
         TabIndex        =   100
         Top             =   480
         Width           =   5295
         Begin VB.Frame Frame15 
            Caption         =   "Peso Posição"
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
            Left            =   2760
            TabIndex        =   144
            Tag             =   "Peso Total da Posição"
            ToolTipText     =   "Peso Total da Posição"
            Top             =   1440
            Width           =   1575
            Begin VB.TextBox txtcadastro 
               Height          =   285
               Index           =   38
               Left            =   120
               TabIndex        =   37
               Tag             =   "Peso Total da Posição"
               ToolTipText     =   "Peso Total da Posição"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   37
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   33
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   26
            Left            =   4440
            TabIndex        =   38
            Tag             =   "Item"
            ToolTipText     =   "Item"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   25
            Left            =   120
            TabIndex        =   36
            Tag             =   "Posição"
            ToolTipText     =   "Posição"
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   31
            Left            =   1680
            TabIndex        =   35
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   28
            Left            =   120
            TabIndex        =   30
            Tag             =   "Descrição do desenho"
            ToolTipText     =   "Descrição do desenho"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   27
            Left            =   4080
            TabIndex        =   32
            Tag             =   "Revisão"
            ToolTipText     =   "Revisão"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   20
            Left            =   1440
            TabIndex        =   31
            Tag             =   "Desenho"
            ToolTipText     =   "Desenho"
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            TabIndex        =   101
            Top             =   480
            Visible         =   0   'False
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6562
            TabIndex        =   102
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
            Height          =   255
            Left            =   4440
            OleObjectBlob   =   "frmLM.frx":65CA
            TabIndex        =   103
            Top             =   1440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":662C
            TabIndex        =   104
            Top             =   1440
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "frmLM.frx":66A0
            TabIndex        =   105
            Top             =   840
            Width           =   2175
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Index           =   0
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":672E
            TabIndex        =   106
            Top             =   240
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Index           =   0
            Left            =   4080
            OleObjectBlob   =   "frmLM.frx":679C
            TabIndex        =   107
            Top             =   240
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Index           =   0
            Left            =   1440
            OleObjectBlob   =   "frmLM.frx":6804
            TabIndex        =   108
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
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
         Height          =   2295
         Left            =   -69480
         TabIndex        =   91
         Top             =   480
         Width           =   6615
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   2040
            TabIndex        =   40
            Top             =   480
            Width           =   3975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "frmLM.frx":686C
            TabIndex        =   143
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   21
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   29
            Left            =   4455
            TabIndex        =   45
            Tag             =   "Quantidade Unitária"
            ToolTipText     =   "Quantidade Unitária"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   16
            Left            =   5280
            TabIndex        =   44
            Tag             =   "Quant. CJ"
            ToolTipText     =   "Quant. CJ"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Height          =   285
            Index           =   19
            Left            =   120
            TabIndex        =   39
            Tag             =   "Código do Material"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   6120
            TabIndex        =   46
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   93
            Top             =   1440
            Width           =   6375
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
               Height          =   495
               Left            =   105
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   94
               Top             =   165
               Width           =   5775
            End
         End
         Begin VB.CommandButton chameleonButton5 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   41
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   3840
            TabIndex        =   43
            Top             =   1080
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   3840
            OleObjectBlob   =   "frmLM.frx":68D4
            TabIndex        =   92
            Top             =   840
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   4440
            OleObjectBlob   =   "frmLM.frx":6936
            TabIndex        =   95
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   5280
            OleObjectBlob   =   "frmLM.frx":69A4
            TabIndex        =   96
            Top             =   840
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   2040
            OleObjectBlob   =   "frmLM.frx":6A16
            TabIndex        =   97
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6A82
            TabIndex        =   98
            Top             =   240
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6AE8
            TabIndex        =   99
            Top             =   840
            Width           =   735
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   62
         Top             =   480
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   2880
            TabIndex        =   11
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   10
            Top             =   2880
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   3240
            TabIndex        =   9
            Top             =   2280
            Width           =   2295
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   8
            Top             =   2280
            Width           =   2175
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   7
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   2880
            TabIndex        =   6
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   4560
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   2
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Tag             =   "Código do Cliente"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmLM.frx":6B52
            TabIndex        =   78
            Top             =   2640
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6BB4
            TabIndex        =   77
            Top             =   2640
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   3240
            OleObjectBlob   =   "frmLM.frx":6C18
            TabIndex        =   76
            Top             =   2040
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   960
            OleObjectBlob   =   "frmLM.frx":6C78
            TabIndex        =   75
            Top             =   2040
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6CE2
            TabIndex        =   74
            Top             =   2040
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   2880
            OleObjectBlob   =   "frmLM.frx":6D48
            TabIndex        =   73
            Top             =   1440
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6DAE
            TabIndex        =   72
            Top             =   1440
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   4560
            OleObjectBlob   =   "frmLM.frx":6E14
            TabIndex        =   71
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6E74
            TabIndex        =   70
            Top             =   840
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmLM.frx":6EDE
            TabIndex        =   69
            Top             =   240
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":6F40
            TabIndex        =   68
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Observações da LM "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   64
         Top             =   5520
         Width           =   5655
         Begin VB.TextBox Text3 
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   5415
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
         Height          =   1515
         Left            =   -74880
         TabIndex        =   63
         Top             =   3960
         Width           =   5655
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   12
            Tag             =   "Código do Contato"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   1200
            TabIndex        =   13
            Top             =   480
            Width           =   4335
         End
         Begin VB.TextBox txtcadastro 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   1920
            TabIndex        =   15
            Top             =   1080
            Width           =   3615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   1920
            OleObjectBlob   =   "frmLM.frx":6FA6
            TabIndex        =   82
            Top             =   840
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":700A
            TabIndex        =   81
            Top             =   840
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmLM.frx":7074
            TabIndex        =   80
            Top             =   240
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmLM.frx":70D6
            TabIndex        =   79
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "FO's - Fichas de Orçamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -69120
         TabIndex        =   61
         Top             =   480
         Width           =   11295
         Begin ZEUS.chameleonButton chamCad 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   6480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   "Importar..."
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
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
            MICON           =   "frmLM.frx":713C
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
            Height          =   6135
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   10821
            LabelEdit       =   1
            MultiSelect     =   -1  'True
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   6015
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   10610
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   7
         Left            =   -73560
         TabIndex        =   52
         Tag             =   "Gerar resumo"
         ToolTipText     =   "Gerar resumo"
         Top             =   3960
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
         MICON           =   "frmLM.frx":7158
         PICN            =   "frmLM.frx":7174
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   1
         Left            =   -74160
         TabIndex        =   51
         Tag             =   "Excluir registro"
         ToolTipText     =   "Excluir registro"
         Top             =   3960
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
         MICON           =   "frmLM.frx":7E4E
         PICN            =   "frmLM.frx":7E6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ZEUS.chameleonButton chamCad 
         Height          =   615
         Index           =   0
         Left            =   -74760
         TabIndex        =   50
         Tag             =   "Inserir registro"
         ToolTipText     =   "Inserir registro"
         Top             =   3960
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
         MICON           =   "frmLM.frx":8B44
         PICN            =   "frmLM.frx":8B60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSScriptControlCtl.ScriptControl ScriptControl1 
         Left            =   -71040
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   54
         Tag             =   "Duplo clique para editar"
         ToolTipText     =   "Duplo clique para editar"
         Top             =   4800
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   4683
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483646
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
      Begin VB.Label Label44 
         Caption         =   "Total Peso Área de Pintura:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   6960
         Width           =   2535
      End
      Begin VB.Label Label45 
         Caption         =   "Total Peso Líquido:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Label lbltotpm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   88
         Top             =   6960
         Width           =   1455
      End
      Begin VB.Label lbltotl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   87
         Top             =   6600
         Width           =   1455
      End
   End
   Begin ZEUS.chameleonButton chamCad 
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   56
      Tag             =   "Gravar dados"
      ToolTipText     =   "Gravar dados"
      Top             =   8880
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
      MICON           =   "frmLM.frx":983A
      PICN            =   "frmLM.frx":9856
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   2280
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmLM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private rsLocal As New ADODB.Recordset
Private rsMaterial As New ADODB.Recordset
Private const0(9) As Double
Private vAr0(9) As Double
Private X As Integer, Y As Integer
Private Conta As Integer, ContaItensFO As Integer
Private Formula As String
Private ForPint As String
Private SqlM As String
Private SomaTotal As Double
Private SomaPint As Double
Private QuantCJ As Double
Private PesoTotal As Double

Private CaminhoArquivo As String
Private NomeArquivo As String
Private pathArq As String
Private Plan As Object 'Aplicação Excel

Private Sub chamCad_Click(Index As Integer)
    Select Case Index
    Case 0
        txtcadastro_KeyDown 24, 13, 1
        txtcadastro(18).Enabled = True
    Case 1
        ExcluirItem
        SomaListview
        LimpaControles
        Label36.Caption = "Inclusão"
        txtcadastro(26).SetFocus
    Case 2
        GravarDadosConjunto
    Case 3
        VerFOSel
    Case 4
        If GravarDados = True Then
            mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
        Else
            If ListView2.ListItems.Count <> 0 Then
                mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ATENÇÃO!!!!!!!!!!!!!!!!!!!!!"
            End If
        End If
    '    txtcadastro(0).SetFocus
    Case 5
        SalvaXLS
    '    Label10 = ""
    Case 6
        mobjMsg.Abrir "Deseja sair da tela de cadastro?", YesNo, pergunta, "ZEUS"
        If Tp = 1 Then
            'CancelaSN = 1
            'GravarDados
            If GravarDados = True Then
                mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
            Else
                If ListView2.ListItems.Count <> 0 Then
                    mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ATENÇÃO!!!!!!!!!!!!!!!!!!!!!"
                End If
            End If
            excluiTabelaExcLM
            Unload Me
        End If
    Case 7
        GerarResumo
        SSTab1.Tab = 2
    '    optCadastro(3).Value = True
    '    Check3.Value = 1
        mobjMsg.Abrir "Resumo gerado com sucesso", Ok, informacao, "Atenção"
    '-------------------
    End Select
End Sub

Private Sub chameleonButton1_Click()
    IncluirDesConj
End Sub

Private Sub IncluirDesConj()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView4.ListItems.Count
    'Edição
    If Y > 0 Then
        For X = 1 To Y
            If ListView4.ListItems.Item(X) = Me.Text1 Then
                Me.txtcadastro(44) = ListView4.ListItems.Item(X) 'Código do conjunto
                ListView4.SelectedItem.ListSubItems.Item(1) = Format(Me.SkinLabel37, "00") 'Codigo sequencial
                ListView4.SelectedItem.ListSubItems.Item(2) = txtcadastro(33) 'Código do Desenho
                ListView4.SelectedItem.ListSubItems.Item(3) = Me.txtcadastro(22).Text 'Desenho
                ListView4.SelectedItem.ListSubItems.Item(4) = Me.txtcadastro(32).Text 'Revisão
                ListView4.SelectedItem.ListSubItems.Item(5) = Me.txtcadastro(34).Text 'Projeto
                ListView4.SelectedItem.ListSubItems.Item(6) = Me.txtcadastro(35).Text 'Quantidade
                ListView4.SelectedItem.ListSubItems.Item(7) = Me.txtcadastro(36).Text 'Posição
                ListView4.SelectedItem.ListSubItems.Item(8) = Val(Me.Label2.Caption) 'LM
                LimpaControlesDC
                Y = ListView4.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView4.ListItems.Add(, , txtcadastro(44))
        Y = ListView4.ListItems.Count
        Me.ListView4.Sorted = True
        Me.ListView4.SortKey = 0
        Me.ListView4.SortOrder = lvwDescending
    'Inclusão
    Else
        Set ItemLst = ListView4.ListItems.Add(, , txtcadastro(44))
        Y = ListView4.ListItems.Count
    End If
    ItemLst.SubItems(1) = Format(Me.SkinLabel37, "00") 'Codigo sequencial
    ItemLst.SubItems(2) = txtcadastro(33) 'Código do Desenho
    ItemLst.SubItems(3) = Me.txtcadastro(22).Text 'Desenho
    ItemLst.SubItems(4) = Me.txtcadastro(32).Text 'Revisão
    ItemLst.SubItems(5) = Me.txtcadastro(34).Text 'Projeto
    ItemLst.SubItems(6) = Me.txtcadastro(35).Text 'Quantidade
    ItemLst.SubItems(7) = Me.txtcadastro(36).Text 'Posição
    ItemLst.SubItems(8) = Val(Me.Label2.Caption) 'LM
    txtcadastro(22).SetFocus
    LimpaControlesDC
End Sub

Private Sub chameleonButton2_Click()
    ExcluirItemLV ListView4
End Sub

Private Sub chameleonButton5_Click()
    ChamaGridProduto
    CarregaDados (19)
    txtcadastro(29).SetFocus
End Sub

Private Sub ChamaGridProduto()
On Error GoTo Err
    Dim F As New frmPesqger2
    Sqlp = "Select a.codigoprd,a.nomefantasia from " & vBancoTotvs & ".dbo.TPRD as a inner join tbmateriais as b on a.IDPRD = b.idprd where a.CODIGOPRD like '%%' and a.codigoprd like '%" & txtcadastro(19) & "%' order by a.nomefantasia"
    procnom = "nomefantasia"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Materiais"
    Pesquisa = frmLM.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        rsLocal.MoveFirst
        rsLocal.Find "CODIGOPRD=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If Pesquisa = "Lista de Materiais" Then Pesquisa = ""
            txtcadastro(19) = Pesquisa
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

Private Sub chameleonButton6_Click()
    If ListView4.ListItems.Count > 0 Then
        txtcadastro(44).Text = Format(GeraCodigoConj(ListView4), "00")
    Else
        txtcadastro(44).Text = Format(Val(txtcadastro(44)) + 1, "00")
    End If
End Sub

Private Sub chameleonButton7_Click()
    ChamaGridMat
End Sub

Private Sub chamCad_MouseOver(Index As Integer)
    Legenda = chamCad(Index).ToolTipText
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub chamCad_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Combo1_Click()
    CarregaDesConjunto ListView5
    SomaQtdCJ
End Sub

Private Sub Combo1_GotFocus()
    CarregaDesConjunto ListView5
End Sub

Private Sub Command1_Click()
    ChamaGridDesenho 20, 27, 28, 37, 48
    txtcadastro(31).SetFocus
End Sub

Private Sub ChamaGridDesenho(posA As String, posB As String, posC As String, posD As String, posE As String)
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    'Sqlp = "Select a.desenho,a.revisao,a.codprojeto,a.iddesenho,b.projeto,b.fce from tbDesenhos as a inner join tbProjetos as b on a.codprojeto = b.codprojeto where b.fce = '" & Label32 & "' order by b.fce,b.projeto,a.desenho,a.revisao"
    Sqlp = "Select a.desenho,a.revisao,a.codprojeto,a.iddesenho,b.projeto,b.fce from tbDesenhos as a inner join tbProjetos as b on a.codprojeto = b.codprojeto order by b.fce,b.projeto,a.desenho,a.revisao"
    procnom = "desenho"
    campo = 0
    Campo1 = 1
    campo2 = 4
    campo3 = 5
    Campo4 = 3
    Load F
    F.Caption = "Pesquisa de Desenho"
    Pesquisa = frmLM.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockReadOnly
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "iddesenho=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(posA).Text = rsLocal.Fields(0) 'Desenho
            txtcadastro(posB).Text = rsLocal.Fields(1) 'Revisao
            txtcadastro(posC).Text = rsLocal.Fields(3) 'IDDesenho
            'Label32.Caption = rsLocal.Fields(5) 'FCE
            txtcadastro(posE).Text = rsLocal.Fields(5) 'FCE
            txtcadastro(posD).Text = rsLocal.Fields(4) 'Projeto
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

Private Sub achaIDDesenho()
    On Error GoTo Err
    Dim rsIDDesenho As New ADODB.Recordset
    Dim SqlIDDesenho As String
    
    SqlIDDesenho = "Select * from tbDesenhos where iddesenho = '" & txtcadastro(28) & "' order by desenho"
    rsIDDesenho.Open SqlIDDesenho, cnBanco, adOpenKeyset, adLockOptimistic
    If rsIDDesenho.EOF Then
        txtcadastro(27).Text = txtcadastro(27)
        mobjMsg.Abrir "Desenho não cadastrado", Ok, critico, "Atenção"
    Else
        txtcadastro(27).Text = rsIDDesenho.Fields(4)
        achaFCEProj rsIDDesenho.Fields(2), rsIDDesenho.Fields(0)
    End If
    rsIDDesenho.Close
    Set rsIDDesenho = Nothing
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

Private Sub achaDesenho()
    On Error GoTo Err
    Dim rsDesenho As New ADODB.Recordset
    Dim SqlDesenho As String
    Dim X As Integer
    
    SqlDesenho = "Select * from tbDesenhos where desenho = '" & txtcadastro(20) & "' order by desenho"
    rsDesenho.Open SqlDesenho, cnBanco, adOpenKeyset, adLockOptimistic
    If rsDesenho.EOF Then
        txtcadastro(20).Text = txtcadastro(20)
        mobjMsg.Abrir "Desenho não cadastrado", Ok, critico, "Atenção"
    Else
        txtcadastro(20).Text = rsDesenho.Fields(3)
    End If
    rsDesenho.Close
    Set rsDesenho = Nothing
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

Private Sub achaRevisao()
    On Error GoTo Err
    Dim rsRevisao As New ADODB.Recordset
    Dim SqlRevisao As String
    
    SqlRevisao = "Select * from tbDesenhos where desenho = '" & txtcadastro(20) & "' and revisao = '" & txtcadastro(27) & "' order by desenho"
    rsRevisao.Open SqlRevisao, cnBanco, adOpenKeyset, adLockOptimistic
    If rsRevisao.EOF Then
        txtcadastro(27).Text = txtcadastro(27)
        mobjMsg.Abrir "Desenho não cadastrado", Ok, critico, "Atenção"
    Else
        txtcadastro(27).Text = rsRevisao.Fields(4)
        achaFCEProj rsRevisao.Fields(2), rsRevisao.Fields(0)
    End If
    rsRevisao.Close
    Set rsRevisao = Nothing
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

Private Sub achaFCEProj(vCodProj As Integer, vIDDesenho)
On Error GoTo Err
    Dim rsProjeto As New ADODB.Recordset
    Dim SqlProjeto As String
    SqlProjeto = "Select * from tbProjetos where codprojeto = '" & vCodProj & "'"
    rsProjeto.Open SqlProjeto, cnBanco, adOpenKeyset, adLockOptimistic
    Label32.Caption = rsProjeto.Fields(1) 'fce
    txtcadastro(37).Text = rsProjeto.Fields(2) 'projeto
    txtcadastro(28).Text = vIDDesenho 'iddesenho
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

Private Sub Command2_Click()
    ChamaGridDesenho 22, 32, 33, 34, 48
End Sub

Private Sub Command3_Click()
    ChamaGridDesenho 39, 41, 42, 43, 49
    'If txtCadastro(49).Text <> Label32 Then
    '    txtCadastro(40).Enabled = False
    '    Command4.Enabled = False
    'Else
        txtcadastro(40).Enabled = True
        Command4.Enabled = True
    'End If
    'txtcadastro(39).SetFocus
End Sub

Private Sub Command4_Click()
    ChamaGridDesenho 40, 45, 46, 47, 48
    'txtcadastro(40).SetFocus
End Sub

Private Sub Command5_Click()
    clonarDesenho
End Sub

Private Sub clonarDesenho()
On Error GoTo Err
'    If Val(Label32) <> Val(txtCadastro(49)) Then
        Dim rsLisview As New ADODB.Recordset
        Dim ItemLst As ListItem
        Dim sql As String
        Dim pegaSequencia As Integer
        PesoTotal = 0
    
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 15
        Me.ListView2.SortOrder = lvwAscending
    
        pegaSequencia = ListView2.ListItems.Count
    
        sql = "select a.fce, a.codlm, a.codseq, c.desenho, c.revisao, b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat, d.posicao, d.item, a.quantcj, a.quantunit, a.dimensoes, a.pesounit, a.area, b.CODTB2FAT, a.codfo, a.observação, d.descposicao, a.matncadast, b.NOMEFANTASIA, " & _
              "b.CODUNDCONTROLE, e.DESCRICAO, c.descricao, a.calcpor, a.idconjunto, c.iddesenho,d.pesoposicao from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
              "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT " & _
              "where a.fce = '" & Val(txtcadastro(49)) & "' and a.codigodes = '" & Val(txtcadastro(42)) & "' order by a.codseq"
        rsLisview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
        
        If rsLisview.RecordCount > 0 Then
            While Not rsLisview.EOF
                
                Set ItemLst = ListView2.ListItems.Add(, , txtcadastro(40))
                ItemLst.SubItems(1) = txtcadastro(45) 'Revisao
                ItemLst.SubItems(2) = txtcadastro(46) 'idDesenho
'                Set ItemLst = ListView2.ListItems.Add(, , rsLisview.Fields(3)) 'Desenho
'                ItemLst.SubItems(1) = "" & rsLisview.Fields(4) ' Revisao
'                ItemLst.SubItems(2) = "" & rsLisview.Fields(24) 'ID Desenho
                
                ItemLst.SubItems(3) = "" & rsLisview.Fields(6) 'Posição
                ItemLst.SubItems(4) = "" & rsLisview.Fields(16)  'Descricao Posição/Marca
                ItemLst.SubItems(5) = "" & rsLisview.Fields(7) 'Item
                ItemLst.SubItems(6) = "" & Format(rsLisview.Fields(5), "000000") 'Código
                If Val(rsLisview.Fields(5)) <> 0 Then
                    ItemLst.SubItems(7) = "" & rsLisview.Fields(18) 'Descrição Material
                Else
                    ItemLst.SubItems(7) = "" & rsLisview.Fields(17) 'Descrição Material
                End If
                If rsLisview.Fields(13) <> 0 Then ItemLst.SubItems(8) = "" & Format(rsLisview.Fields(13), "000000") & "-" & rsLisview.Fields(20) Else ItemLst.SubItems(8) = "-" 'Código do tipo+ descrição do material - Ex.: 000001 - ASTM A36 ' Tipo Material (codigo+descrição)
                ItemLst.SubItems(9) = "" & rsLisview.Fields(10) 'Dimensão
                ItemLst.SubItems(10) = "" & rsLisview.Fields(9) 'Q. Unit
                ItemLst.SubItems(11) = "" & Format(rsLisview.Fields(11), "#,##0.000;(#,##0.000)") 'Peso Unit/Qtd
                ItemLst.SubItems(12) = "" & rsLisview.Fields(8) 'Q CJ
                If ItemLst.SubItems(8) <> "-" Then ItemLst.SubItems(13) = ItemLst.SubItems(7) & ItemLst.SubItems(8) Else ItemLst.SubItems(13) = ItemLst.SubItems(7) 'codigo+material
                PesoTotal = Format(rsLisview.Fields(8) * rsLisview.Fields(9) * rsLisview.Fields(11), "#,##0.000;(#,##0.000")
                ItemLst.SubItems(14) = "" & Format(PesoTotal, "#,##0.000;(#,##0.000)") 'Peso total
                ItemLst.SubItems(15) = Format(GeraCodigoCloneLV(ListView2), "0000") 'sequencia
                If Val(rsLisview.Fields(5)) <> 0 Then
                    ItemLst.SubItems(16) = "" & rsLisview.Fields(19) 'UN
                Else
                    ItemLst.SubItems(16) = "pç" 'UN
                End If
                ItemLst.SubItems(17) = "" & Format(rsLisview.Fields(12), "#,##0.000;(#,##0.000)") 'Área de Pintura
                If rsLisview.Fields(14) <> "Null" Then ItemLst.SubItems(20) = Format(rsLisview.Fields(14), "000000") Else ItemLst.SubItems(20) = "-" 'FO
                ItemLst.SubItems(18) = "" & rsLisview.Fields(15) 'Observação
                ItemLst.SubItems(22) = "" & rsLisview.Fields(22) 'Calculador por
                If Not IsNull(rsLisview.Fields(23)) Then ItemLst.SubItems(23) = "" & rsLisview.Fields(23) Else ItemLst.SubItems(23) = "-" 'ID Conjunto
            
                If IsNull(rsLisview.Fields(25)) Then ItemLst.SubItems(24) = "0" Else ItemLst.SubItems(24) = rsLisview.Fields(25)  'Peso Posicao
                ItemLst.ListSubItems(18).Bold = True
                ItemLst.ForeColor = &H404080
                ItemLst.ListSubItems(1).ForeColor = &H404080
                ItemLst.ListSubItems(2).ForeColor = &H404080
                ItemLst.ListSubItems(3).ForeColor = &H404080
                ItemLst.ListSubItems(4).ForeColor = &H404080
                ItemLst.ListSubItems(5).ForeColor = &H404080
                Dim rsConstantes As New ADODB.Recordset
                Dim SqlConstantes As String
                SqlConstantes = "Select * from tbconstantes as a where a.idprd = '" & Val(Mid$(rsLisview.Fields(5), 15, 6)) & "'order by a.idprd"
                rsConstantes.Open SqlConstantes, cnBanco, adOpenKeyset, adLockOptimistic
                If rsConstantes.RecordCount > 0 Then
                    If Val(rsLisview.Fields(5)) <> 0 Then ItemLst.SubItems(19) = rsConstantes.Fields(1) 'Peso Especifico
                End If
                If Val(ItemLst.SubItems(6)) <> 0 Then
                
                    If Mid$(ItemLst.SubItems(7), 1, 5) <> "CHAPA" And Mid$(ItemLst.SubItems(7), 1, 5) <> "GRADE" And Mid$(ItemLst.SubItems(7), 1, 4) <> "TELA" And Val(ItemLst.SubItems(6)) <> 0 Then
                        If ItemLst.SubItems(9) <> "" And ItemLst.SubItems(9) <> "-" Then
                            ItemLst.SubItems(21) = ItemLst.SubItems(9) * ItemLst.SubItems(10) * ItemLst.SubItems(12) 'Comprimento
                        Else
                            ItemLst.SubItems(21) = ItemLst.SubItems(14) / ItemLst.SubItems(19) * 1000 'Comprimento
                        End If
                    Else
                        ItemLst.SubItems(21) = 0 'Comprimento
                    End If
                Else
                    ItemLst.SubItems(21) = 0 'Comprimento
                End If
                rsConstantes.Close
                Set rsConstantes = Nothing
            
                rsLisview.MoveNext
            Wend
        End If
    'ABAIXO ROTINA QUE CLONA DESENHOS QUE ESTAO DENTRO DA PROPRIA LM
'    Else
    
'    ' Pega os dados da FO selecionada no Listview1
'        Dim Y As Integer, X As Integer, Contador As Integer
'        Dim fce As String
'        'Dim ItemLst As ListItem
'        Y = ListView2.ListItems.Count
'        Contador = 0
'
'        ListView2.Sorted = True
'        ListView2.SortKey = 15
'        ListView2.SortOrder = lvwAscending
'
'        For X = 1 To Y
'            ListView2.ListItems(X).Selected = True
'            If ListView2.SelectedItem.ListSubItems.Item(2) = txtCadastro(42).Text Then
'                Set ItemLst = ListView2.ListItems.Add(, , txtCadastro(40))
'                ItemLst.SubItems(1) = txtCadastro(45) 'Revisao
'                ItemLst.SubItems(2) = txtCadastro(46) 'idDesenho
'                ItemLst.SubItems(3) = ListView2.SelectedItem.ListSubItems.Item(3) 'posição
'                ItemLst.SubItems(4) = ListView2.SelectedItem.ListSubItems.Item(4) 'Descrição Posição/Marca
'                ItemLst.SubItems(5) = ListView2.SelectedItem.ListSubItems.Item(5) 'item
'                ItemLst.SubItems(6) = ListView2.SelectedItem.ListSubItems.Item(6) 'codigo do material
'                ItemLst.SubItems(7) = ListView2.SelectedItem.ListSubItems.Item(7) 'descrição do material
'                ItemLst.SubItems(8) = ListView2.SelectedItem.ListSubItems.Item(8) 'codigo+descrição do tipo de material
'                ItemLst.SubItems(9) = ListView2.SelectedItem.ListSubItems.Item(9) 'dimensão
'                ItemLst.SubItems(10) = ListView2.SelectedItem.ListSubItems.Item(10) 'Qtd. Unit.
'
'                ItemLst.SubItems(11) = ListView2.SelectedItem.ListSubItems.Item(11) 'Peso Unit./Qtd.
'                ItemLst.SubItems(12) = ListView2.SelectedItem.ListSubItems.Item(12) 'Qtd. CJ
'                ItemLst.SubItems(13) = ListView2.SelectedItem.ListSubItems.Item(13) 'Código+material
'                ItemLst.SubItems(14) = ListView2.SelectedItem.ListSubItems.Item(14) 'peso total
'                ItemLst.SubItems(16) = ListView2.SelectedItem.ListSubItems.Item(16) 'UN
'                ItemLst.SubItems(17) = ListView2.SelectedItem.ListSubItems.Item(17) 'Área de pintura
'                ItemLst.SubItems(18) = ListView2.SelectedItem.ListSubItems.Item(18) 'observação
'                ItemLst.SubItems(19) = ListView2.SelectedItem.ListSubItems.Item(18) 'peso específico
'                ItemLst.SubItems(20) = ListView2.SelectedItem.ListSubItems.Item(20) 'FO
'
'                ItemLst.SubItems(21) = ListView2.SelectedItem.ListSubItems.Item(21) 'Comprimento
'                ItemLst.SubItems(22) = ListView2.SelectedItem.ListSubItems.Item(22) 'Cálculo por
'                ItemLst.SubItems(23) = ListView2.SelectedItem.ListSubItems.Item(23) 'Conjunto
'                ItemLst.SubItems(24) = ListView2.SelectedItem.ListSubItems.Item(24) 'Peso posicao
'                ItemLst.SubItems(15) = Format(GeraCodigoCloneLV(ListView2), "0000") 'sequencia
'            End If
'        Next
'    End If
    mobjMsg.Abrir "Clonagem realizada com sucesso", Ok, informacao, "Zeus"
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

Public Function GeraCodigoCloneLV(LV As Listview)
On Error GoTo Err
    If LV.ListItems.Count > 0 Then
'        Dim X As Integer
'        X = 1
        LV.Sorted = True
        LV.SortKey = 15
        LV.SortOrder = lvwDescending
        LV.ListItems.Item(1).Selected = True
        GeraCodigoCloneLV = LV.SelectedItem.ListSubItems.Item(15) + 1
        LV.SortOrder = lvwAscending
        Exit Function
    Else
        GeraCodigoCloneLV = 1
    End If
Err:
    GeraCodigoCloneLV = 1
    Resume Next
End Function

Private Sub Form_Activate()
    excluiTabelaExcLM
    criaTabelaExcLM
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GravarDados = True Then
        mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
        excluiTabelaExcLM
    Else
        If ListView2.ListItems.Count <> 0 Then
            mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ATENÇÃO!!!!!!!!!!!!!!!!!!!!!"
            Cancel = 1
        End If
    End If
End Sub

Private Sub ListView2_Click()
    AlterarItem1
End Sub

Private Sub ListView2_DblClick()
    AlterarItem1
End Sub

Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    AlterarItem1
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
    AlterarItem1
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then txtcadastro(20).SetFocus
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    Principal.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    If varGlobal = "-" Or varGlobal = "" Then
        mobjMsg.Abrir "Nenhuma FCE selecionada", Ok, critico, "ZEUS"
        Unload Me
        Exit Sub
    End If
    SSTab1.Tab = 0
    SomaTotal = 0
    SomaPint = 0
    listview_cabecalho
    CompoeCombo2 cboCadastro(0), "tbTipoFCE", "id", "nome"
    CompoeControles
    optCadastro_Click 1
    LimpaControlesDC
    'Initialize edit box
    txtLvw = ""
    txtLvw.Visible = False
    txtLvw.Tag = False 'is ListView2 dirty, not used in this example
    
    If ListView4.ListItems.Count > 0 Then
        txtcadastro(44).Text = Format(GeraCodigoConj(ListView4), "00")
    Else
        txtcadastro(44).Text = Format(Val(txtcadastro(44)) + 1, "00")
    End If
    CompoeCombo2 Combo1, "tbDesConjunto", Label32.Caption, "idConjunto"
    
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    'On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub listview_cabecalho()
    ListView2.ColumnHeaders.Add , , "Desenho", ListView2.Width / 10 'gravado
    ListView2.ColumnHeaders.Add , , "Rev", ListView2.Width / 26 'gravado
    ListView2.ColumnHeaders.Add , , "ID Desenho", ListView2.Width / 15 'gravado
    ListView2.ColumnHeaders.Add , , "Posição", ListView2.Width / 17 'gravado
    ListView2.ColumnHeaders.Add , , "Descrição Posição/Marca", ListView2.Width / 8 'gravado
    ListView2.ColumnHeaders.Add , , "Item", ListView2.Width / 21 'gravado
    ListView2.ColumnHeaders.Add , , "Código", ListView2.Width / 9 'gravado
    ListView2.ColumnHeaders.Add , , "Descrição", ListView2.Width / 4 'gravado
    ListView2.ColumnHeaders.Add , , "Material", ListView2.Width / 6 'gravado
    ListView2.ColumnHeaders.Add , , "Dimensão", ListView2.Width / 10 'gravado
    ListView2.ColumnHeaders.Add , , "Q.Unit", ListView2.Width / 16 'gravado
    ListView2.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView2.Width / 7.6 'gravado
    ListView2.ColumnHeaders.Add , , "Q.CJ", ListView2.Width / 19.5 'gravado
    ListView2.ColumnHeaders.Add , , "codigo+material", ListView2.Width / 10000 'gravado
    ListView2.ColumnHeaders.Add , , "Peso Total", ListView2.Width / 7 'calculado
    ListView2.ColumnHeaders.Add , , "sequencia", ListView2.Width / 10000 'gravado
    ListView2.ColumnHeaders.Add , , "Un", ListView2.Width / 28 'gravado
    ListView2.ColumnHeaders.Add , , "Área Pint.", ListView2.Width / 10 'calculado
    ListView2.ColumnHeaders.Add , , "Observação", ListView2.Width / 7 'gravado
    ListView2.ColumnHeaders.Add , , "Peso Especifico", ListView2.Width / 10000 'gravado
    ListView2.ColumnHeaders.Add , , "FO", ListView2.Width / 16 'gravado
    ListView2.ColumnHeaders.Add , , "Comprimento", ListView2.Width / 10000 'calculado
    ListView2.ColumnHeaders.Add , , "Calculo por", ListView2.Width / 10000 'gravado
    ListView2.ColumnHeaders.Add , , "Conjunto", ListView2.Width / 10000 'gravado
    ListView2.ColumnHeaders.Add , , "Peso Posição", ListView2.Width / 10 'gravado
    
    ListView3.ColumnHeaders.Add , , "Item", ListView3.Width / 16
    ListView3.ColumnHeaders.Add , , "Código", ListView3.Width / 16
    ListView3.ColumnHeaders.Add , , "Descrição", ListView3.Width / 5
    ListView3.ColumnHeaders.Add , , "Material", ListView3.Width / 6
    ListView3.ColumnHeaders.Add , , "Un", ListView3.Width / 32
    ListView3.ColumnHeaders.Add , , "Peso Unit/Qtd.", ListView3.Width / 7.6
    ListView3.ColumnHeaders.Add , , "Área Pint.", ListView3.Width / 10
    ListView3.ColumnHeaders.Add , , "Comprimento", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Peso Especifico", ListView3.Width / 10000
    ListView3.ColumnHeaders.Add , , "Observação", ListView3.Width / 5
    
    ListView4.ColumnHeaders.Add , , "Conjunto", ListView4.Width / 16
    ListView4.ColumnHeaders.Add , , "ID", ListView4.Width / 16
    ListView4.ColumnHeaders.Add , , "ID Desenho", ListView4.Width / 10000
    ListView4.ColumnHeaders.Add , , "Desenho", ListView4.Width / 8
    ListView4.ColumnHeaders.Add , , "Rev", ListView4.Width / 20
    ListView4.ColumnHeaders.Add , , "Projeto", ListView4.Width / 4
    ListView4.ColumnHeaders.Add , , "Qtd.", ListView4.Width / 18
    ListView4.ColumnHeaders.Add , , "Posição", ListView4.Width / 12
    ListView4.ColumnHeaders.Add , , "LM", ListView4.Width / 10000
    
    ListView5.ColumnHeaders.Add , , "ID", ListView5.Width / 10
    ListView5.ColumnHeaders.Add , , "ID Desenho", ListView5.Width / 10000
    ListView5.ColumnHeaders.Add , , "Desenho", ListView5.Width / 6
    ListView5.ColumnHeaders.Add , , "Rev", ListView5.Width / 9
    ListView5.ColumnHeaders.Add , , "Projeto", ListView5.Width / 4
    ListView5.ColumnHeaders.Add , , "Qtd.", ListView5.Width / 10
    ListView5.ColumnHeaders.Add , , "Posição", ListView5.Width / 6
    
    Me.ListView2.ColumnHeaders(10).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(14).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(17).Alignment = lvwColumnRight
    
    Me.ListView3.ColumnHeaders(6).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(7).Alignment = lvwColumnRight
    Me.ListView3.ColumnHeaders(9).Alignment = lvwColumnRight
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "FO", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 1.3
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport
    ListView3.View = lvwReport
    ListView4.View = lvwReport
    ListView5.View = lvwReport
End Sub

Private Sub LimpaControlesDC()
    Dim X As Integer
    txtcadastro(22) = ""
    txtcadastro(32) = ""
    txtcadastro(33) = ""
    txtcadastro(34) = ""
    txtcadastro(35) = ""
    txtcadastro(36) = ""
    
    
    If ListView4.ListItems.Count > 0 Then
        SkinLabel37.Caption = Format(GeraCodigo(ListView4), "00")
    Else
        SkinLabel37.Caption = Format(Val(SkinLabel37) + 1, "00")
    End If
End Sub

Private Function GeraCodigo(LV As Listview)
    Dim X As Integer, Y As Integer
    X = 1
    LV.SortOrder = lvwAscending
    Y = LV.ListItems.Count
    If Y = 0 Then
        GeraCodigo = 0
    Else
        GeraCodigo = LV.ListItems.Item(X)
        For X = 1 To Y
            LV.ListItems.Item(X).Selected = True
            If LV.ListItems.Item(X).Selected = True Then
                If GeraCodigo <> LV.SelectedItem.ListSubItems.Item(1) Then
                    GeraCodigo = LV.SelectedItem.ListSubItems.Item(1)
                End If
            End If
        Next
        GeraCodigo = GeraCodigo + 1
    End If
    Exit Function
End Function

Private Function GeraCodigoConj(LV As Listview)
    Dim X As Integer, Y As Integer
    X = 1
    LV.SortOrder = lvwAscending
    Y = LV.ListItems.Count
    If Y = 0 Then
        GeraCodigoConj = 0
    Else
        GeraCodigoConj = LV.ListItems.Item(X)
        For X = 1 To Y
            LV.ListItems.Item(X).Selected = True
            If LV.ListItems.Item(X).Selected = True Then
                If GeraCodigoConj <> LV.ListItems.Item(X) Then
                    GeraCodigoConj = LV.ListItems.Item(X)
                End If
            End If
        Next
        GeraCodigoConj = GeraCodigoConj + 1
    End If
    Exit Function
End Function

Private Sub ExcluirItem()
On Error GoTo Err
    Dim rsExcItemLM As New ADODB.Recordset
    Dim sqlExcItemLM As String
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y = 0 Then Exit Sub
    
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            ListView2.ListItems.Item(X).Checked = True
        End If
    Next
    
    sqlExcItemLM = "Select * from tbExcluidosLM" & vTime & ""
    rsExcItemLM.Open sqlExcItemLM, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To Y
        If X > Y Then Exit For
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Selected = True
            rsExcItemLM.AddNew
            rsExcItemLM(0) = Val(Label32) 'fce
            rsExcItemLM(1) = Val(Label2) 'codigo LM
            rsExcItemLM(2) = Val(ListView2.SelectedItem.ListSubItems.Item(15)) 'Sequência
            ListView2.ListItems.Remove (X)
            X = X - 1
            Y = ListView2.ListItems.Count
        End If
    Next
    rsExcItemLM.Update
    rsExcItemLM.Close
    Set rsExcItemLM = Nothing
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Ocorreu um erro, Selecione um item antes de excluir", Ok, critico, "Atenção"
        Exit Sub
    End If
End Sub

Private Function IncluirItem()
On Error GoTo Err
    IncluirItem = True
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, ProxSeq As Integer
    If ValidaCampo2 = False Then Exit Function
    'Variavel do sistema para calculo da area de pintura, consta na formula de pintura
    If optCadastro(0).Value = True Then
        'A formula abaixo considera o peso de conjunto na multiplicação
        PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(16) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
        'A formula abaixo NAO considera o peso de conjunto na multiplicação
        'PesoTotal = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
    Else
        PesoTotal = Format(txtcadastro(23) * txtcadastro(16) * txtcadastro(29), "#,##0.000;(#,##0.000)")
'        PesoTotal = Format(txtcadastro(23) * txtcadastro(29), "#,##0.000;(#,##0.000)")
    End If
    Y = ListView2.ListItems.Count
    
    ' IF DE ALTERAÇÃO
    If Label36.Caption = "Alteração" Then
        ListView2.ListItems(Val(Label37)).Selected = True
        ListView2.ListItems(Val(Label37)).EnsureVisible
        'Listview2.SelectedItem.ListSubItems.Item (1)
        Label36.Caption = "Inclusão"
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            
            If Me.txtcadastro(24) <> ListView2.SelectedItem.ListSubItems.Item(9) Then
                ListView2.SelectedItem.ListSubItems.Item(17) = Format(ScriptControl1.Eval(ForPint) * txtcadastro(15), "#,##0.000;(#,##0.000)")
            End If
            If Me.txtcadastro(23) <> ListView2.SelectedItem.ListSubItems.Item(9) Then
                If Val(txtcadastro(19)) <> 0 Then ListView2.SelectedItem.ListSubItems.Item(17) = Format(ScriptControl1.Eval(ForPint) * txtcadastro(15), "#,##0.000;(#,##0.000)")
            End If
            ListView2.SelectedItem.ListSubItems.Item(14) = Format(Me.txtcadastro(16) * Me.txtcadastro(16) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
'            ListView2.SelectedItem.ListSubItems.Item(14) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
        End If
        
        If optCadastro(0).Value = True Then
            ListView2.SelectedItem.ListSubItems.Item(11) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)")
            ListView2.SelectedItem.ListSubItems.Item(14) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(16) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
'            ListView2.SelectedItem.ListSubItems.Item(14) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)")
            
            'Informa se o calculo de peso e area de pintura esta sendo baseado pela dimensao
            ListView2.SelectedItem.ListSubItems.Item(22) = "Dimensão"
        Else
            ListView2.SelectedItem.ListSubItems.Item(11) = Format(txtcadastro(23), "#,##0.000;(#,##0.000)")
            ListView2.SelectedItem.ListSubItems.Item(14) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            'Informa que os caculos de peso e area de pintura estão sendo calculado baseados no peso lançado
            ListView2.SelectedItem.ListSubItems.Item(22) = "Peso"
        End If
        
        If Me.txtcadastro(23) <> ListView2.SelectedItem.ListSubItems.Item(11) Then
            If optCadastro(1).Value = True Then
                ListView2.SelectedItem.ListSubItems.Item(11) = Format(txtcadastro(23), "#,##0.000;(#,##0.000)")
                ListView2.SelectedItem.ListSubItems.Item(14) = Format(PesoTotal, "#,##0.000;(#,##0.000)")
            End If
        End If
        ListView2.ListItems.Item(Val(Label37)) = Me.txtcadastro(20).Text ' Desenho
        ListView2.SelectedItem.ListSubItems.Item(1) = Me.txtcadastro(27).Text 'Revisao
        ListView2.SelectedItem.ListSubItems.Item(2) = Me.txtcadastro(28).Text 'Descricao desenho
        ListView2.SelectedItem.ListSubItems.Item(3) = Me.txtcadastro(25).Text 'Posição
        ListView2.SelectedItem.ListSubItems.Item(4) = Me.txtcadastro(31) 'Descrição Posição/Marca
        ListView2.SelectedItem.ListSubItems.Item(5) = Me.txtcadastro(26).Text 'Item
        ListView2.SelectedItem.ListSubItems.Item(6) = txtcadastro(19) & " - " & SkinLabel40 'Código
        ListView2.SelectedItem.ListSubItems.Item(7) = Me.txtcadastro(18).Text '
        ListView2.SelectedItem.ListSubItems.Item(8) = Me.txtcadastro(21).Text 'Me.txtcadastro(22).Text & "-" & Me.txtcadastro(21).Text
        ListView2.SelectedItem.ListSubItems.Item(9) = Me.txtcadastro(24).Text
        ListView2.SelectedItem.ListSubItems.Item(10) = Me.txtcadastro(29).Text
        ListView2.SelectedItem.ListSubItems.Item(12) = Me.txtcadastro(16).Text
        ListView2.SelectedItem.ListSubItems.Item(13) = Me.txtcadastro(19).Text & Me.txtcadastro(21).Text
        ListView2.SelectedItem.ListSubItems.Item(16) = Me.txtcadastro(17).Text
        ListView2.SelectedItem.ListSubItems.Item(19) = Me.txtcadastro(30) 'Peso especifico
        
        'A DESATIVAÇÃO DO BLOCO ABAIXO ESTA EM ESTUDO ----------------------------
        
        'If Mid$(ListView2.SelectedItem.ListSubItems.Item(7), 1, 5) <> "GRADE" And Val(ListView2.SelectedItem.ListSubItems.Item(6)) <> 0 Then
        
        'If Mid$(ListView2.SelectedItem.ListSubItems.Item(7), 1, 5) <> "CHAPA" And Mid$(ListView2.SelectedItem.ListSubItems.Item(7), 1, 5) <> "GRADE" And Val(ListView2.SelectedItem.ListSubItems.Item(6)) <> 0 Then
        '    If ListView2.SelectedItem.ListSubItems.Item(9) <> "" Then
        '        ListView2.SelectedItem.ListSubItems.Item(21) = ListView2.SelectedItem.ListSubItems.Item(9) * ListView2.SelectedItem.ListSubItems.Item(10) * ListView2.SelectedItem.ListSubItems.Item(12)
        '    Else
        '        ListView2.SelectedItem.ListSubItems.Item(21) = ListView2.SelectedItem.ListSubItems.Item(14) / ListView2.SelectedItem.ListSubItems.Item(19) * 1000
        '    End If
        'Else
            ListView2.SelectedItem.ListSubItems.Item(21) = 0
        'End If
        
        'A DESATIVAÇÃO DO BLOCO ACIMA ESTA EM ESTUDO ----------------------------
        
        
        
        ListView2.SelectedItem.ListSubItems.Item(14) = Format(ListView2.SelectedItem.ListSubItems.Item(10) * ListView2.SelectedItem.ListSubItems.Item(12) * ListView2.SelectedItem.ListSubItems.Item(11), "#,##0.000;(#,##0.000)")
'        ListView2.SelectedItem.ListSubItems.Item(14) = Format(ListView2.SelectedItem.ListSubItems.Item(10) * ListView2.SelectedItem.ListSubItems.Item(11), "#,##0.000;(#,##0.000)")
        ListView2.SelectedItem.ListSubItems.Item(23) = Combo1.Text
        ListView2.SelectedItem.ListSubItems.Item(24) = Me.txtcadastro(38).Text 'Peso Posição

        ListView2.SetFocus
    ' IF DE INCLUSAO
    Else
        'Ordena Listview pela sequencia de cadastramento antes de gravar
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 15
        Me.ListView2.SortOrder = lvwAscending
        '------
        If ListView2.ListItems.Count > 0 Then
            ListView2.ListItems(ListView2.ListItems.Count).Selected = True
            ListView2.ListItems(ListView2.ListItems.Count).EnsureVisible
            ProxSeq = Val(ListView2.SelectedItem.ListSubItems.Item(15)) + 1
        Else
            ProxSeq = 1
        End If
        Set ItemLst = ListView2.ListItems.Add(, , txtcadastro(20))
        Label36.Caption = "Inclusão"
        ItemLst.SubItems(1) = Me.txtcadastro(27).Text 'RevisaoLabel38
        ItemLst.SubItems(2) = Me.txtcadastro(28) 'Descricao Desenho
        ItemLst.SubItems(3) = Me.txtcadastro(25) 'posição
        ItemLst.SubItems(4) = Me.txtcadastro(31) 'Descrição Posição/Marca
        ItemLst.SubItems(5) = Me.txtcadastro(26) 'item
        ItemLst.SubItems(6) = txtcadastro(19) & " - " & SkinLabel40 'codigo do material
        ItemLst.SubItems(7) = Me.txtcadastro(18).Text 'descrição do material
        ItemLst.SubItems(8) = Me.txtcadastro(21).Text 'codigo+descrição do tipo de material
        ItemLst.SubItems(9) = Me.txtcadastro(24).Text 'dimensão
        ItemLst.SubItems(10) = Me.txtcadastro(29) 'Qtd. Unit.
        If optCadastro(0).Value = True Then
            ItemLst.SubItems(11) = Format(ScriptControl1.Eval(Formula), "#,##0.000;(#,##0.000)") 'peso unit/qtd
            ItemLst.SubItems(14) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(16) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)") 'peso total
'            ItemLst.SubItems(14) = Format(ScriptControl1.Eval(Formula) * Me.txtcadastro(29), "#,##0.000;(#,##0.000)") 'peso total
            'Informa se o calculo de peso e area de pintura esta sendo baseado pela dimensao
            ItemLst.SubItems(22) = "Dimensão"
        Else
            ItemLst.SubItems(11) = Format(txtcadastro(23), "#,##0.000;(#,##0.000)") 'peso unit/qtd
            ItemLst.SubItems(14) = Format(PesoTotal, "#,##0.000;(#,##0.000)") 'peso total
            'Informa que os caculos de peso e area de pintura estão sendo calculado baseados no peso lançado
            ItemLst.SubItems(22) = "Peso"
        End If
        ItemLst.SubItems(12) = Me.txtcadastro(16).Text 'Quantidade conjunto
        ItemLst.SubItems(13) = Me.txtcadastro(18).Text & Me.txtcadastro(21).Text 'Descrição material+descricao tipo material
        ItemLst.SubItems(15) = Format(ProxSeq, "0000") 'Sequencia numerica do Listiview
        ItemLst.SubItems(16) = Me.txtcadastro(17).Text 'unidade
        If Check1.Value = 1 Then
            'Variavel q contem a formula para calcular a área de pintura
            'O Replace esta sendo aplicado aki pq so agora q foi encontrado o PesoTotal
            ForPint = Replace(ForPint, "pesototal", PesoTotal)
            ForPint = Replace(ForPint, ",", ".")
            ItemLst.SubItems(17) = Format(ScriptControl1.Eval(ForPint) * txtcadastro(15), "#,##0.000;(#,##0.000)")
        End If
        ItemLst.SubItems(20) = "-" 'FO
        ItemLst.SubItems(19) = Me.txtcadastro(30) 'Peso especifico
        ListView2.ListItems(ListView2.ListItems.Count).Selected = True
        ItemLst.SubItems(18) = "-" 'Observação
        
        If Mid$(ItemLst.SubItems(7), 1, 5) <> "CHAPA" And Mid$(ItemLst.SubItems(7), 1, 5) <> "GRADE" And Mid$(ItemLst.SubItems(7), 1, 5) <> "BARRA" And Mid$(ItemLst.SubItems(7), 1, 4) <> "TELA" And Val(ItemLst.SubItems(6)) <> 0 Then
            If ItemLst.SubItems(9) <> "" Then
                ItemLst.SubItems(21) = ItemLst.SubItems(9) * ItemLst.SubItems(10) * ItemLst.SubItems(12)
            Else
                ItemLst.SubItems(21) = ItemLst.SubItems(14) / ItemLst.SubItems(19) * 1000
            End If
        Else
            ItemLst.SubItems(21) = 0
        End If
        ItemLst.SubItems(23) = Combo1.Text
        ItemLst.SubItems(24) = Me.txtcadastro(38).Text 'Peso Posição
        
        ListView2.ListItems(ListView2.ListItems.Count).EnsureVisible
            
        'Deixar a coluna de OBSERVAÇÃO em negrito e vermelho
        ItemLst.ListSubItems(18).Bold = True
'        ItemLst.ListSubItems(18).ForeColor = vbRed
        ItemLst.ForeColor = &H404080
        ItemLst.ListSubItems(1).ForeColor = &H404080
        ItemLst.ListSubItems(2).ForeColor = &H404080
        ItemLst.ListSubItems(3).ForeColor = &H404080
        ItemLst.ListSubItems(4).ForeColor = &H404080
        ItemLst.ListSubItems(5).ForeColor = &H404080
        ItemLst.SubItems(14) = Format(ItemLst.SubItems(10) * ItemLst.SubItems(12) * ItemLst.SubItems(11), "#,##0.000;(#,##0.000)")
    End If
    Me.ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    If txtcadastro(17) <> "pç" And txtcadastro(17) <> "PÇ" Then
        If optCadastro(0).Value = True Then
            SomaTotal = SomaTotal + ScriptControl1.Eval(Formula) * Me.txtcadastro(16)
'            SomaTotal = SomaTotal + ScriptControl1.Eval(Formula)
        Else
            SomaTotal = Format(SomaTotal + txtcadastro(23) * Me.txtcadastro(16), "#,##0.000;(#,##0.000")
'            SomaTotal = Format(SomaTotal + txtcadastro(23), "#,##0.000;(#,##0.000")
        End If
        
        If Check1.Value = 1 Then SomaPint = SomaPint + ScriptControl1.Eval(ForPint) * Me.txtcadastro(15)
    End If
    txtcadastro(19) = ""
    txtcadastro(18) = ""
    txtcadastro(23) = ""
    txtcadastro(24) = ""
    'Text1.Text = ""
    
    Conta = 0
    LimpaControles
    
    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Ocorreu um erro ao tentar inserir o item. Provavelmente um erro na sitaxe de uma das fórmulas", Ok, critico, "Atenção"
        IncluirItem = False
        ListView2.ListItems.Remove (Val(ItemLst.SubItems(15)))
        Exit Function
    End If
End Function

Private Sub AlterarItem1()
    If ListView2.ListItems.Count < 1 Then Exit Sub
    Label36.Caption = "Alteração"
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Label37 = X
    Me.txtcadastro(20) = ListView2.ListItems.Item(X)
    Me.txtcadastro(27) = ListView2.SelectedItem.ListSubItems.Item(1)
    'Me.txtcadastro(22) = Mid$(ListView2.SelectedItem.ListSubItems.Item(8), 1, 6)
    Me.txtcadastro(19) = Mid$(ListView2.SelectedItem.ListSubItems.Item(6), 1, 11)
    Me.txtcadastro(17) = ListView2.SelectedItem.ListSubItems.Item(16)
    Me.txtcadastro(24) = ListView2.SelectedItem.ListSubItems.Item(9)
    Me.txtcadastro(16) = ListView2.SelectedItem.ListSubItems.Item(12)
    Me.txtcadastro(28) = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.txtcadastro(25) = ListView2.SelectedItem.ListSubItems.Item(3)
    Me.txtcadastro(31) = ListView2.SelectedItem.ListSubItems.Item(4)
    Me.txtcadastro(26) = ListView2.SelectedItem.ListSubItems.Item(5)
    Me.txtcadastro(29) = ListView2.SelectedItem.ListSubItems.Item(10)
    Me.txtcadastro(38) = ListView2.SelectedItem.ListSubItems.Item(24)
    
    'Label38 = ListView2.SelectedItem.ListSubItems.Item(14)
    'Label39 = ListView2.SelectedItem.ListSubItems.Item(11)
    
    Combo1.Text = ListView2.SelectedItem.ListSubItems.Item(23)
    If ListView2.SelectedItem.ListSubItems.Item(22) <> "Dimensão" Then
        Me.optCadastro(1).Value = True
        optCadastro_Click (1)
        Me.txtcadastro(23) = ListView2.SelectedItem.ListSubItems.Item(11)
        Check1.Value = 0
    End If
    If ListView2.SelectedItem.ListSubItems.Item(22) = "Dimensão" Then
        Me.optCadastro(0).Value = True
        optCadastro_Click (0)
        txtcadastro(19).BackColor = &H80000005
    End If
    CarregaDesConjunto ListView5
    If Val(txtcadastro(19)) = 0 Then
        optCadastro_Click (1)
    End If
    txtcadastro_KeyDown 22, 13, 2
    txtcadastro_KeyDown 19, 13, 19
    If txtcadastro(19) = "000000" Then
        txtcadastro(23).BackColor = &HC0C0FF
        Me.txtcadastro(17) = ListView2.SelectedItem.ListSubItems.Item(16)
        Me.txtcadastro(18) = ListView2.SelectedItem.ListSubItems.Item(7)
    End If
    achaIDDesenho
    If ListView2.SelectedItem.ListSubItems.Item(17) = "0,000" Or ListView2.SelectedItem.ListSubItems.Item(17) = "" Then Check1.Value = 0 Else Check1.Value = 1
    SomaPesoSelecionado
    ListView2.SetFocus
End Sub

Private Sub GravarDadosConjunto()
On Error GoTo Err
    Dim rsDeleta As New ADODB.Recordset
    Dim rsGravaDesConjunto As New ADODB.Recordset
    
    Dim SqlDesConjunto As String
    
10  SqlDesConjunto = "Delete from tbDesConjunto where tbDesConjunto.codlm = '" & Val(Label2) & "'"
    rsDeleta.Open SqlDesConjunto, cnBanco
    Y = ListView4.ListItems.Count
    cnBanco.BeginTrans
    
    SqlDesConjunto = "select * from tbDesConjunto"
    rsGravaDesConjunto.Open SqlDesConjunto, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To Y
        ListView4.ListItems.Item(X).Selected = True
        If Val(ListView4.SelectedItem.ListSubItems.Item(8)) = Val(Label2.Caption) Then
            rsGravaDesConjunto.AddNew
            rsGravaDesConjunto.Fields(0) = ListView4.ListItems.Item(X) 'Código do Conjunto
            rsGravaDesConjunto.Fields(1) = Val(ListView4.SelectedItem.ListSubItems.Item(8)) 'Código da Lista de Materiais - LM
            rsGravaDesConjunto.Fields(2) = Val(ListView4.SelectedItem.ListSubItems.Item(1)) 'Identificador sequencial dos desenhos de conjunto
            rsGravaDesConjunto.Fields(3) = Val(ListView4.SelectedItem.ListSubItems.Item(2)) 'Identificador do desenho
            rsGravaDesConjunto.Fields(4) = ListView4.SelectedItem.ListSubItems.Item(6) 'Quantidade
            rsGravaDesConjunto.Fields(5) = ListView4.SelectedItem.ListSubItems.Item(7) 'Posição
        End If
        'X = X + 1
    Next
    If Not rsGravaDesConjunto.EOF Then rsGravaDesConjunto.Update
    
    cnBanco.CommitTrans
    rsGravaDesConjunto.Close
    Set rsGravaDesConjunto = Nothing
    
    mobjMsg.Abrir "Dados dos conjuntos gravados com sucesso", Ok, informacao, "Zeus"
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

Private Function GravarDados()
On Error GoTo Err

    Dim CodigoPos As Long
    Dim rsDeleta As New ADODB.Recordset
    Dim rsGravaItemLM As New ADODB.Recordset
    Dim rsGravaLM As New ADODB.Recordset
    Dim rsGravaResumo As New ADODB.Recordset

    Dim sqlExc As String
    Dim sql As String
    Dim Y As Integer, X As Integer

10  GravarDados = True

    cnBanco.BeginTrans
    
    
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        
        sql = "Select * from tbitemlm where fce = '" & Val(Label32) & "' and codlm = '" & Val(Label2) & "' and codseq = '" & Val(ListView2.SelectedItem.ListSubItems.Item(15)) & "' order by fce, codlm, codseq"
        rsGravaItemLM.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaItemLM.RecordCount = 0 Then
            rsGravaItemLM.AddNew
            rsGravaItemLM(0) = Label32 'fce
            rsGravaItemLM(1) = Label2 'codigo LM
            rsGravaItemLM(2) = Val(ListView2.SelectedItem.ListSubItems.Item(15)) 'Sequência
        End If
        rsGravaItemLM(3) = Val(ListView2.SelectedItem.ListSubItems.Item(2)) ' ID Desenho

        Dim rsPosicaoItem As New ADODB.Recordset
        Dim SqlPosicaoItem As String

        SqlPosicaoItem = "Select * from tbposicoes where tbposicoes.codigodes=" & " '" & Val(ListView2.SelectedItem.ListSubItems.Item(2)) & "'" & _
        "and tbposicoes.posicao=" & " '" & ListView2.SelectedItem.ListSubItems.Item(3) & "'" & "and tbposicoes.item=" & " '" & ListView2.SelectedItem.ListSubItems.Item(5) & "'"
        rsPosicaoItem.Open SqlPosicaoItem, cnBanco, adOpenKeyset, adLockOptimistic

        If Not rsPosicaoItem.EOF Then 'Se o desenho/posicao/item existir na tabela tbposicoes
            CodigoPos = rsPosicaoItem(0)
            rsPosicaoItem(1) = Val(ListView2.SelectedItem.ListSubItems.Item(2)) 'codigo do Desenho
            rsPosicaoItem(2) = ListView2.SelectedItem.ListSubItems.Item(3) 'posicao
            rsPosicaoItem(3) = ListView2.SelectedItem.ListSubItems.Item(4) 'descricao posicao
            rsPosicaoItem(4) = ListView2.SelectedItem.ListSubItems.Item(5) 'item
            rsPosicaoItem(7) = ListView2.SelectedItem.ListSubItems.Item(24) 'Peso Posição
            rsGravaItemLM(4) = CodigoPos
        Else 'Se o desenho/revisao não existir na tabela tbdesenho
            rsPosicaoItem.AddNew
            CodigoPos = GeraCodPos
            rsPosicaoItem(0) = CodigoPos
            rsPosicaoItem(1) = Val(ListView2.SelectedItem.ListSubItems.Item(2)) 'codigo do Desenho
            rsPosicaoItem(2) = ListView2.SelectedItem.ListSubItems.Item(3) 'posicao
            rsPosicaoItem(3) = ListView2.SelectedItem.ListSubItems.Item(4) 'descricao posicao
            rsPosicaoItem(4) = ListView2.SelectedItem.ListSubItems.Item(5) 'item
            rsPosicaoItem(7) = ListView2.SelectedItem.ListSubItems.Item(24) 'Peso Posição
            rsGravaItemLM(4) = CodigoPos 'codigo do Posicao na tabela tbitemlm
        End If
        If Not rsPosicaoItem.EOF Then rsPosicaoItem.Update
        rsPosicaoItem.Close
''-------
        rsGravaItemLM(5) = Val(Mid$(ListView2.SelectedItem.ListSubItems.Item(6), 15, 6)) 'codmat
        'rsGravaItemLM(6) = ListView2.SelectedItem.ListSubItems.Item(3) 'Posição
        'rsGravaItemLM(7) = ListView2.SelectedItem.ListSubItems.Item(4) 'Item
        rsGravaItemLM(6) = ListView2.SelectedItem.ListSubItems.Item(12) 'Quantidade CJ
        rsGravaItemLM(8) = ListView2.SelectedItem.ListSubItems.Item(9) 'Dimensoes
        rsGravaItemLM(7) = ListView2.SelectedItem.ListSubItems.Item(10) 'Quant unit
        rsGravaItemLM(9) = Format(ListView2.SelectedItem.ListSubItems.Item(11), "#,##0.000;(#,##0.000)") 'Pesounit
        If ListView2.SelectedItem.ListSubItems.Item(17) <> "" Then rsGravaItemLM(10) = Format(ListView2.SelectedItem.ListSubItems.Item(17), "#,##0.000;(#,##0.000)") Else rsGravaItemLM(10) = 0 'Area
        rsGravaItemLM(11) = Val(Mid$(ListView2.SelectedItem.ListSubItems.Item(8), 1, 6)) 'Tipomat
        If ListView2.SelectedItem.ListSubItems.Item(20) <> "-" Then rsGravaItemLM(12) = ListView2.SelectedItem.ListSubItems.Item(20) 'codfo
        rsGravaItemLM(13) = ListView2.SelectedItem.ListSubItems.Item(18) 'observacao
        If Val(ListView2.SelectedItem.ListSubItems.Item(6)) = 0 Then rsGravaItemLM(14) = ListView2.SelectedItem.ListSubItems.Item(7) 'material n cadastrado
        rsGravaItemLM(15) = ListView2.SelectedItem.ListSubItems.Item(22) 'Calculado por
        If ListView2.SelectedItem.ListSubItems.Item(23) <> "-" Then rsGravaItemLM(16) = ListView2.SelectedItem.ListSubItems.Item(23) 'Conjunto

        If Not rsGravaItemLM.EOF Then rsGravaItemLM.Update
        rsGravaItemLM.Close
    Next
    Set rsGravaItemLM = Nothing


   'VERIFICA SE A TABELA EXISTE. SE EXISTE ENTRA NA CONDIÇÃO ABAIXO
    sqlExc = "SELECT * FROM SYSOBJECTS WHERE XTYPE = 'U' AND NAME = '" & "tbExcluidosLM" & vTime & "'"
    rsDeleta.Open sqlExc, cnBanco, adOpenKeyset, adLockReadOnly
    If rsDeleta.RecordCount > 0 Then
        rsDeleta.Close
        Set rsDeleta = Nothing
        sqlExc = "select * from tbExcluidosLM" & vTime & " order by codseq"
        rsDeleta.Open sqlExc, cnBanco, adOpenKeyset, adLockReadOnly
    
        Dim rsExcluiItem As New ADODB.Recordset
        Dim sqlExcluirItem As String
        While Not rsDeleta.EOF
            sqlExcluirItem = "delete from tbitemlm where fce = '" & rsDeleta.Fields(0) & "' and codlm = '" & rsDeleta.Fields(1) & "' and codseq = '" & rsDeleta.Fields(2) & "'"
            rsExcluiItem.Open sqlExcluirItem, cnBanco
            rsDeleta.MoveNext
        Wend
        rsDeleta.Close
        Set rsDeleta = Nothing
    End If
    
'--------------------------
'DELETA ITENS TABELA
'--------------------------



    sql = "Select * from tblm where tblm.fce = '" & Val(Label32) & "'" & _
    "and tblm.codlm=" & " '" & Val(Label2) & "'"
    rsGravaLM.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsGravaLM.EOF Then
        rsGravaLM.AddNew
    End If
    rsGravaLM(0) = Val(Label32)
    rsGravaLM(1) = Val(Label2)
    rsGravaLM(2) = Format(DTPicker1, "dd/mm/yyyy")
    rsGravaLM(3) = Text7
    rsGravaLM(5) = Text3
    rsGravaLM(6) = "S"

    If cboCadastro(0).Text <> "-" Then
        'rsGravaLM(7) = cboCadastro(0).ItemData(cboCadastro(0).ListIndex) 'ID Tipo FCE
        rsGravaLM(8) = cboCadastro(0).Text 'Descricao Tipo FCE
    End If

    If Not rsGravaLM.EOF Then rsGravaLM.Update

    cnBanco.CommitTrans

'    rsGravaItemLM.Close
    rsGravaLM.Close

    Exit Function
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        GoTo 10
    Else
        GravarDados = False
        mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
        cnBanco.RollbackTrans
        Exit Function
    End If
End Function

Private Sub CompoeControles()
On Error GoTo Err
    Dim llng_Contador As Long
    Dim SqlTreeview As String
    Dim Y As Integer, X As Integer, i As Integer
    
    Dim rsFCE As New ADODB.Recordset
    Dim rsClientes As New ADODB.Recordset
    Dim rsContatos As New ADODB.Recordset
    Dim sqlFCE As String
    Dim sqlClientes As String
    Dim sqlContatos As String

    sqlFCE = "select tbfo.codclifor, tbfo.codcontato from tbfo where tbFO.FCE = '" & Val(Mid$(varGlobal, 1, 4)) & "'"
    rsFCE.Open sqlFCE, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsFCE.RecordCount > 0 Then
        txtcadastro(0) = rsFCE.Fields(0)
        If Not IsNull(rsFCE.Fields(1)) Then txtcadastro(11) = rsFCE.Fields(1)
        Label32 = Mid$(varGlobal, 1, 4)
    End If
    If Pesquisa = "novo" Then
        Label2 = Format(GeraCodLM, "000000") & ""
    Else
        Label2 = Mid$(varGlobal, 5, 6)
    End If
    DTPicker1 = Date
    rsFCE.Close
    Set rsFCE = Nothing
    CarregaCorpoLM
    CarregaCli
    CarregaContato
    ContFOSel
    If Pesquisa = "editar" Then RestauraItens
    CarregaDesConjunto ListView4
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
        mobjMsg.Abrir "Cliente não cadastrado", Ok, critico, "Atenção"
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

Private Sub CarregaCorpoLM()
On Error GoTo Err
    Dim rsCorpoLM As New ADODB.Recordset
    Dim SqlCorpoCli As String
    
    'SqlCorpoCli = "Select * from tblm where tblm.fce = '" & Val(Label32) & "'" & _
    '"and tblm.codlm=" & " '" & Val(Label2) & "'"
    
    SqlCorpoCli = ""
    SqlCorpoCli = SqlCorpoCli & "SELECT " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.FCE, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.CODLM, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.DATAABERTURA, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.DESCRICAO, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.LD, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.OBSERVACAO, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " A.ATIVO, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " CASE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "     WHEN A.TIPOFCEDESC IS NOT NULL THEN " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "         A.TIPOFCEDESC " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "     ELSE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "         CASE WHEN FILTRO.TIPO IS NULL OR FILTRO.TIPO = '' THEN '-' ELSE SUBSTRING(FILTRO.TIPO,1,LEN(FILTRO.TIPO)-1) END " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " END AS TIPO_FCE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "FROM TBLM AS A " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "LEFT JOIN " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " ( " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " SELECT  FCE, " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "     COALESCE( " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "         (SELECT CAST(TIPOFCEDESC AS VARCHAR(10)) + '/' AS [text()] " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "          FROM TBPEDIDOS AS O " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "          WHERE O.FCE  = C.FCE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "          GROUP BY TIPOFCEDESC " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "          FOR XML PATH(''), TYPE).value('.[1]', 'VARCHAR(MAX)'), '') AS [TIPO] " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " FROM TBPEDIDOS AS C " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " GROUP BY FCE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & " ) AS FILTRO ON A.FCE = FILTRO.FCE " & vbCrLf
    SqlCorpoCli = SqlCorpoCli & "WHERE A.FCE = '" & Val(Label32) & "' AND A.CODLM = " & " '" & Val(Label2) & "'"
    
    rsCorpoLM.Open SqlCorpoCli, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsCorpoLM.EOF Then
        If rsCorpoLM.Fields(2) <> "Null" Then DTPicker1 = rsCorpoLM.Fields(2)
        If rsCorpoLM.Fields(3) <> "Null" Then Text7.Text = rsCorpoLM.Fields(3)
        If rsCorpoLM.Fields(5) <> "Null" Then Text3.Text = rsCorpoLM.Fields(5)
        cboCadastro(0) = rsCorpoLM.Fields(7)
    End If
    rsCorpoLM.Close
    Set rsCorpoLM = Nothing
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

Private Sub CarregaDesConjunto(LV As Listview)
On Error GoTo Err
    Dim ItemLst As ListItem
    Dim rsDesConjunto As New ADODB.Recordset
    Dim SqlDesConjunto As String
    If LV.Name = "ListView4" Then
        SqlDesConjunto = "select a.idConjunto,a.codlm,a.idseq,a.iddesenho,b.desenho,b.revisao,c.projeto,a.quantidade,a.posicao from tbDesConjunto as a inner join tbDesenhos as b on a.iddesenho = b.iddesenho inner join tbProjetos as c on b.codprojeto = c.codprojeto where c.fce = '" & Label32 & "'"
    Else
        SqlDesConjunto = "select a.idConjunto,a.codlm,a.idseq,a.iddesenho,b.desenho,b.revisao,c.projeto,a.quantidade,a.posicao from tbDesConjunto as a inner join tbDesenhos as b on a.iddesenho = b.iddesenho inner join tbProjetos as c on b.codprojeto = c.codprojeto where c.fce = '" & Label32 & "' and a.idconjunto = '" & Val(Combo1) & "' order by a.idconjunto"
    End If
    rsDesConjunto.Open SqlDesConjunto, cnBanco, adOpenKeyset, adLockOptimistic
    LV.ListItems.Clear
    While Not rsDesConjunto.EOF
        If LV.Name = "ListView4" Then
            Set ItemLst = LV.ListItems.Add(, , Format(rsDesConjunto(0), "00")) 'Conjunto
            ItemLst.SubItems(1) = "" & Format(rsDesConjunto.Fields(2), "00") 'Sequencia
            ItemLst.SubItems(2) = "" & rsDesConjunto.Fields(3) 'Identificador desenho
            ItemLst.SubItems(3) = "" & rsDesConjunto.Fields(4) 'Desenho
            ItemLst.SubItems(4) = "" & rsDesConjunto.Fields(5) 'Revisao
            ItemLst.SubItems(5) = "" & rsDesConjunto.Fields(6) 'Projeto
            ItemLst.SubItems(6) = "" & rsDesConjunto.Fields(7) 'Quantidade
            ItemLst.SubItems(7) = "" & rsDesConjunto.Fields(8) 'Posicao
            ItemLst.SubItems(8) = "" & rsDesConjunto.Fields(1) 'LM
        Else
            Set ItemLst = LV.ListItems.Add(, , Format(rsDesConjunto.Fields(2), "00")) 'Sequencia
            ItemLst.SubItems(1) = "" & rsDesConjunto.Fields(3) 'Identificador desenho
            ItemLst.SubItems(2) = "" & rsDesConjunto.Fields(4) 'Desenho
            ItemLst.SubItems(3) = "" & rsDesConjunto.Fields(5) 'Revisao
            ItemLst.SubItems(4) = "" & rsDesConjunto.Fields(6) 'Projeto
            ItemLst.SubItems(5) = "" & rsDesConjunto.Fields(7) 'Quantidade
            ItemLst.SubItems(6) = "" & rsDesConjunto.Fields(8) 'Posicao
        End If
        rsDesConjunto.MoveNext
    Wend
    LV.Sorted = True
    LV.SortKey = 0
    LV.SortOrder = lvwAscending
    rsDesConjunto.Close
    Set rsDesConjunto = Nothing
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
    Dim ItemLst As ListItem
    Dim rsLV As New ADODB.Recordset
    Dim SqlLV As String
    Dim Y As Integer, codfornec As Integer
    Dim numFCE As String
    Y = vListViewPrincipal.ListItems.Count
    
    For X = 1 To Y
        If vListViewPrincipal.ListItems.Item(X).Selected = True Then
            numFCE = Mid(varGlobal, 1, 6) 'frmPesqFCE.ListView1.ListItems.Item(X)
            Exit For
        End If
    Next
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

Private Sub VerFOSel()
' Pega os dados da FO selecionada no Listview1
    Dim Y As Integer, X As Integer, Contador As Integer
    Dim fce As String
    Y = ListView1.ListItems.Count
    fce = ""
    Contador = 0
    ContaItensFO = 1
    For X = 1 To Y
        ListView1.ListItems(X).Selected = True
        If ListView1.ListItems.Item(X).Selected = True Then
            If ListView1.ListItems.Item(X).Checked = True Then
                varGlobal = ListView1.ListItems.Item(X)
                If CheckFOImport = False Then
                    mobjMsg.Abrir "Essa FO ja foi carregada anteriormente", Ok, critico, "Atenção"
                    Exit Sub
                End If
                CarregaLM
                Contador = Contador + 1
            End If
        End If
    Next
    If Contador = 0 Then
        mobjMsg.Abrir "Nenhuma FO Selecionada", Ok, critico, "Atenção"
        Exit Sub
    End If
    mobjMsg.Abrir "Importação realizada com sucesso", Ok, informacao, "Zeus"
End Sub

Private Sub CarregaLM()
On Error GoTo Err
    Dim rsFO As New ADODB.Recordset
    Dim sqlFO As String
    Dim rsLisview As New ADODB.Recordset
    Dim ItemLst As ListItem
    Dim sql As String
    
    If ListView2.ListItems.Count > 0 Then
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 15
        Me.ListView2.SortOrder = lvwAscending
        ContaItensFO = ListView2.SelectedItem.ListSubItems.Item(15) + 1
    End If
    
    sql = "select tblistamaterial.codfo, tblistamaterial.codseq, tblistamaterial.desenho, tblistamaterial.codmat, tblistamaterial.quantcj,tblistamaterial.dimensoes, tblistamaterial.pesounit, tblistamaterial.area, tbmateriais.descricao, tbmateriais.unidade, tblistamaterial.TipoMat, tblistamaterial.revisao, tbtipomat.descricao[DescTipoMat], tblistamaterial.observacao  from tblistamaterial left join tbmateriais on tblistamaterial.codmat = tbmateriais.codmaterial left join tbtipomat on tblistamaterial.TipoMat=tbtipomat.codigo where tblistamaterial.codfo = '" & Val(varGlobal) & "'order by codseq"
    rsLisview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsLisview.RecordCount > 0 Then
        While Not rsLisview.EOF
            Set ItemLst = ListView2.ListItems.Add(, , rsLisview.Fields(2)) 'Nº do desenho
            ItemLst.SubItems(1) = "" & rsLisview.Fields(11) 'revisao
            ItemLst.SubItems(2) = "-" 'Descrição desenho
            ItemLst.SubItems(3) = "-" 'Posição
            ItemLst.SubItems(5) = "-"  'Item
            ItemLst.SubItems(6) = "" & Format(rsLisview.Fields(3), "000000") 'Código do material
            ItemLst.SubItems(7) = "" & rsLisview.Fields(8) 'Descrição do material
            If rsLisview.Fields(10) <> 0 Then ItemLst.SubItems(8) = "" & Format(rsLisview.Fields(10), "000000") & "-" & rsLisview.Fields(12) Else ItemLst.SubItems(8) = "-" 'Código do tipo+ descrição do material - Ex.: 000001 - ASTM A36
            ItemLst.SubItems(9) = "" & rsLisview.Fields(5) 'Dimensoes do material
            ItemLst.SubItems(10) = 1 'Quantidade unitaria
            ItemLst.SubItems(11) = "" & Format(rsLisview.Fields(6), "#,##0.000;(#,##0.000)") 'Peso Unit/qtd
            ItemLst.SubItems(12) = "" & rsLisview.Fields(4) 'quantidade do CJ
            If ItemLst.SubItems(8) <> "-" Then ItemLst.SubItems(13) = ItemLst.SubItems(7) & ItemLst.SubItems(8) Else ItemLst.SubItems(13) = ItemLst.SubItems(7) 'codigo+material
'            PesoTotal = Format(rsLisview.Fields(4) * rsLisview.Fields(6), "#,##0.000;(#,##0.000")
            PesoTotal = Format(rsLisview.Fields(6), "#,##0.000;(#,##0.000")
            
            ItemLst.SubItems(14) = "" & Format(PesoTotal, "#,##0.000;(#,##0.000)") 'Peso total
            
            ItemLst.SubItems(16) = "" & rsLisview.Fields(9) 'Unidade de medida
            If rsLisview.Fields(3) = 0 Then
                ItemLst.SubItems(7) = "" & rsLisview.Fields(13)
                ItemLst.SubItems(16) = "PÇ"
            End If
            ItemLst.SubItems(15) = Format(ContaItensFO, "0000") 'Sequência
            
            ItemLst.SubItems(17) = "" & Format(rsLisview.Fields(7), "#,##0.000;(#,##0.000)") 'Área de pintura
            ItemLst.SubItems(20) = Format(varGlobal, "000000")
            
            Dim rsConstantes As New ADODB.Recordset
            Dim SqlConstantes As String
            SqlConstantes = "Select * from tbconstantes where tbconstantes.codmaterial= '" & Val(rsLisview.Fields(6)) & "'order by tbconstantes.codigo"
            rsConstantes.Open SqlConstantes, cnBanco, adOpenKeyset, adLockOptimistic
            If ItemLst.SubItems(19) <> "" Then ItemLst.SubItems(19) = rsConstantes.Fields(1)
            
            ItemLst.SubItems(18) = "-" 'Observação
            
            'MsgBox "Descricao: " & Mid$(ItemLst.SubItems(7), 1, 5) & "- Código: " & Format(rsLisview.Fields(3), "000000") & "- Dimensão" & ItemLst.SubItems(9)
            
            If Mid$(ItemLst.SubItems(7), 1, 5) <> "CHAPA" And rsLisview.Fields(3) <> 0 Then
                If ItemLst.SubItems(9) <> "" Then
                    ItemLst.SubItems(21) = ItemLst.SubItems(9) * ItemLst.SubItems(10) * ItemLst.SubItems(12) 'Comprimento
                Else
                    ItemLst.SubItems(21) = ItemLst.SubItems(14) / ItemLst.SubItems(19) * 1000 ' Comprimento
                End If
            Else
                ItemLst.SubItems(21) = 0
            End If
            
            If rsLisview.Fields(5) = "" Or rsLisview.Fields(5) = "-" Then
                ItemLst.SubItems(22) = "Peso"
            Else
                ItemLst.SubItems(22) = "Dimensão"
            End If
            
            rsConstantes.Close
            Set rsConstantes = Nothing
            
            SomaTotal = SomaTotal + PesoTotal
            SomaPint = SomaPint + rsLisview.Fields(7)
            PesoTotal = 0
            
            'Deixar a coluna de OBSERVAÇÃO em negrito e vermelho
            ItemLst.ListSubItems(18).Bold = True
            'ItemLst.ListSubItems(18).ForeColor = vbRed
            ItemLst.ForeColor = &H404080
            ItemLst.ListSubItems(1).ForeColor = &H404080
            ItemLst.ListSubItems(2).ForeColor = &H404080
            ItemLst.ListSubItems(3).ForeColor = &H404080
            ItemLst.ListSubItems(4).ForeColor = &H404080
            ItemLst.ListSubItems(5).ForeColor = &H404080
            
            'vai para o proximo registro
            ContaItensFO = ContaItensFO + 1
            rsLisview.MoveNext
        Wend
    End If
    lblTotal = Format(SomaTotal, "#,##0.000;(#,##0.000)")
    lblTotPint = Format(SomaPint, "#,##0.000;(#,##0.000)")
    Me.ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    rsLisview.Close
    ListView2.Refresh
    Set rsLisview = Nothing
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

Private Function CheckFOImport()
    Dim F As Integer, G As Integer
    Dim fce As String
    CheckFOImport = True
    F = ListView2.ListItems.Count
    For G = 1 To F
        ListView2.ListItems(G).Selected = True
        If ListView2.SelectedItem.ListSubItems.Item(17) = varGlobal Then
            CheckFOImport = False
            Exit Function
        End If
    Next
End Function

Private Function GeraCodLM()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tblm where tblm.fce = '" & Val(Label32) & "' order by codlm Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsGeraCodigo.RecordCount > 0 Then
        GeraCodLM = rsGeraCodigo.Fields(1) + 1
    Else
        QualForm = "novalm"
        GeraCodLM = NovoCodigo
    End If
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

Private Function GeraCodDes()
On Error GoTo Err
    Dim rsGeraCodigoDes As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbdesenho order by codigodes Desc"
    rsGeraCodigoDes.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsGeraCodigoDes.RecordCount > 0 Then
        GeraCodDes = rsGeraCodigoDes.Fields(0) + 1
    Else
        GeraCodDes = 1
    End If
    rsGeraCodigoDes.Close
    Set rsGeraCodigoDes = Nothing
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

Private Function GeraCodPos()
On Error GoTo Err
    Dim rsGeraCodigoPos As New ADODB.Recordset
    Dim SqlGera As String
    SqlGera = "Select top 1 * from tbposicoes order by codigopos Desc"
    rsGeraCodigoPos.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    
    If rsGeraCodigoPos.RecordCount > 0 Then
        GeraCodPos = rsGeraCodigoPos.Fields(0) + 1
    Else
        GeraCodPos = 1
    End If
    
    rsGeraCodigoPos.Close
    Set rsGeraCodigoPos = Nothing
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

Private Sub optCadastro_Click(Index As Integer)
    If optCadastro(0).Value = True Then
        txtcadastro(24).Enabled = True
        txtcadastro(23).Enabled = False
        txtcadastro(23).Text = ""
        txtcadastro(23).BackColor = &H80000004
        txtcadastro(24).BackColor = &H80000005
    End If
    If optCadastro(1).Value = True Then
        txtcadastro(23).Enabled = True
        txtcadastro(24).Enabled = False
        txtcadastro(24).Text = ""
        txtcadastro(24).BackColor = &H80000004
        txtcadastro(23).BackColor = &H80000005
    End If
End Sub

Private Sub txtcadastro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'On Error GoTo Error
    If Index = 19 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtcadastro(19) = "" Then ChamaGridProduto
            
            txtcadastro(19).Text = Format(txtcadastro(19).Text, "00-00-00000")
            txtcadastro(19).Text = Replace(txtcadastro(19).Text, "-", ".")
            
            CarregaDados (Index)
            'txtcadastro(29).SetFocus
        End If
    ElseIf Index = 24 Then 'Or Index = 34 Or Index = 35 Then
    If KeyCode = &H8 Then
            txtcadastro(24) = ""
            Formula = ""
            ForPint = ""
            Conta = 0
            CarregaDados (0)
            'txtcadastro(24).SetFocus
        End If
        If Conta > 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                If Index = 24 Then
                    If Label36 = "Inclusão" Then
                        CapVar
                    Else
                        If Val(txtcadastro(19)) <> 0 Then SeparaDados
                    End If
                End If
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
                
                'If ValidaCampo2 = False Then Exit Sub

                If Index = 24 Then
                    If IncluirItem() = False Then GoTo Error
                    'DESABILITANDO LINHAS DE SALVAMENTO AUTOMATICO
                    'If GravarDados = True Then
                    '    mobjMsg.Abrir "Dados Salvos com sucesso!", Ok, informacao, "ZEUS"
                    'Else
                    '    mobjMsg.Abrir "Erro ao gravar dados", Ok, critico, "ATENÇÃO!!!!!!!!!!!!!!!!!!!!!"
                    'End If
                End If
                txtcadastro(38).SetFocus
            End If
            If KeyCode = &H6D Then
                If Index = 24 Then CapVar
                'If Index = 34 Or Index = 35 Then CapVar2
            End If
        ElseIf Conta = 0 Then
            If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
                vAr0(1) = Val(txtcadastro(Index))
                Text2.Text = vAr0(1)
                Y = X
                X = Len(txtcadastro(Index))
                Conta = Conta + 1
                Text2.Text = Formula
                txtcadastro_KeyDown Index, 13, 1
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
    ElseIf Index = 16 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'If Val(txtcadastro(19)) <> 0 Then
            'txtcadastro(29).SetFocus
        End If
    ElseIf Index = 18 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'If Val(txtcadastro(19)) <> 0 Then
            chameleonButton5.SetFocus
        End If
    ElseIf Index = 20 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            achaDesenho
            'txtcadastro(27).SetFocus
        End If
    ElseIf Index = 25 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'txtcadastro(26).SetFocus
        End If
    ElseIf Index = 26 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'txtcadastro(19).SetFocus
        End If
    ElseIf Index = 27 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            achaRevisao
'            txtcadastro(28).SetFocus
        End If
    ElseIf Index = 28 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
'            txtcadastro(22).SetFocus
        End If
    ElseIf Index = 31 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'txtcadastro(25).SetFocus
        End If
    ElseIf Index = 23 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            chamCad_Click (0)
            LimpaControles
            'txtcadastro(26).SetFocus
        End If
    ElseIf Index = 22 Then
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            'If txtCadastro(9) <> "" Then
                Pesquisa = 1
        End If
    End If
    Exit Sub
Error:
    'MsgBox "aki"
    Exit Sub
End Sub

Private Sub CarregaDados(Index)
On Error GoTo Err
    Dim X As Integer
    If Index <> 70 Then ' indice do codigo de material do resumo
        If Val(txtcadastro(19)) = 0 Then
            If Pesquisa = "" Then
                Exit Sub
            End If
            
            Check1.Enabled = False
            txtcadastro(15).Enabled = False
            optCadastro(1).Value = True
            Check1.Value = 0
            
            Exit Sub
        Else
            'txtcadastro(18).Enabled = False
            txtcadastro(17).Enabled = True
            'txtcadastro(21).Enabled = True
            txtcadastro(24).Enabled = True
            txtcadastro(18).BackColor = &H80000005
            txtcadastro(17).BackColor = &H80000005
            txtcadastro(21).BackColor = &H80000005
            txtcadastro(16).BackColor = &H80000005
            txtcadastro(19).BackColor = &H80000005
            txtcadastro(23).BackColor = &H80000005
            txtcadastro(29).BackColor = &H80000005
            chameleonButton5.Enabled = True
            'Text1.FontBold = False
            If optCadastro(0).Value = True Then optCadastro_Click (0) Else optCadastro_Click (1)
            Check1.Enabled = True
            Check1.Value = 1
            txtcadastro(15).Enabled = True
        End If
    End If
    
    If Index = 19 Or Index = 0 Then SqlM = "Select a.CODIGOPRD,a.NOMEFANTASIA,b.formula,b.constpint,c.valconst,a.CODUNDCONTROLE,b.forpint,b.observacao,d.DESCRICAO,a.idprd from " & vBancoTotvs & ".dbo.tprd as a left join tbMateriais as b on b.idprd = a.idprd left Join tbconstantes as c on b.idprd = c.idprd left join " & vBancoTotvs & ".dbo.TTB2 as d on a.CODTB2FAT = d.CODTB2FAT where a.CODIGOPRD = '" & txtcadastro(19) & "' and d.CODCOLIGADA = 1 and b.formula is not null order by c.idseq"
    If Index = 70 Then SqlM = "Select a.CODIGOPRD,a.NOMEFANTASIA,b.formula,b.constpint,c.valconst,a.CODUNDCONTROLE,b.forpint,b.observacao,d.DESCRICAO,a.idprd from " & vBancoTotvs & ".dbo.tprd as a left join tbMateriais as b on b.idprd = a.idprd left Join tbconstantes as c on b.idprd = c.idprd left join " & vBancoTotvs & ".dbo.TTB2 as d on a.CODTB2FAT = d.CODTB2FAT where a.CODIGOPRD = '" & txtcadastro(70) & "' and d.CODCOLIGADA = 1 and b.formula is not null  order by c.idseq"
    rsMaterial.Open SqlM, cnBanco, adOpenKeyset, adLockReadOnly
    'If Not rsMaterial.EOF Then rsMaterial.MoveFirst
    
    'If Index = 19 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtcadastro(19)) & "'"
    'If Index = 70 Then rsMaterial.Find "codmaterial=" & "'" & Val(Me.txtcadastro(70)) & "'"
    
    If rsMaterial.EOF Then
        If Index = 19 Then txtcadastro(0).Text = Format(txtcadastro(19), "000000") & ""
        If Index = 70 Then txtcadastro(0).Text = Format(txtcadastro(70), "000000") & ""
        mobjMsg.Abrir "Código de material não cadastrado", Ok, critico, "Atenção"
    Else
        If Index = 70 Then
            txtcadastro(70).Text = rsMaterial.Fields(0)
            txtcadastro(71).Text = rsMaterial.Fields(1)
            Formula = rsMaterial.Fields(2)
            ForPint = rsMaterial.Fields(6)
            txtcadastro(37).Text = rsMaterial.Fields(7)
        End If
        
        If Index = 19 Or Index = 0 Then
            txtcadastro(19).Text = rsMaterial.Fields(0)
            txtcadastro(18).Text = rsMaterial.Fields(1)
            If Not IsNull(rsMaterial(4)) Then txtcadastro(30).Text = rsMaterial.Fields(4)
            If Not IsNull(rsMaterial(2)) Then
                Formula = rsMaterial.Fields(2)
            Else
                mobjMsg.Abrir "Produto selecionado não possui FÓRMULA cadastrada", , critico
                optCadastro(1).Value = True
            End If
            SkinLabel40 = rsMaterial.Fields(9)
            If Not IsNull(rsMaterial(6)) Then ForPint = rsMaterial.Fields(6)
            txtcadastro(17) = rsMaterial(5)
            If Not IsNull(rsMaterial(3)) Then txtcadastro(15) = rsMaterial(3)
            If Not IsNull(rsMaterial(8)) Then txtcadastro(21) = rsMaterial(8)
            If Not IsNull(rsMaterial(7)) Then Text1 = rsMaterial(7)
            txtcadastro(17).Enabled = False
        End If
        For X = 1 To rsMaterial.RecordCount
            If Not IsNull(rsMaterial(4)) Then const0(X) = rsMaterial.Fields(4)
            rsMaterial.MoveNext
        Next
        'If Index = 19 Then chameleonButton5.SetFocus
    End If
    'rsConstantes.Close
    'Set rsConstantes = Nothing
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

Private Sub ChamaGridMat()
On Error GoTo Err
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTipoMat order by descricao"
    procnom = "descricao"
    campo = 1
    Campo1 = 0
    Pesquisa = frmLM.Tag
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "descricao=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtcadastro(22).Text = Format(rsLocal.Fields(0), "000000")
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

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ColumnSort ListView2, ColumnHeader
End Sub

Public Sub ColumnSort(ListViewControl As Listview, Column As ColumnHeader)
    With ListView2
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
'    If ListView4.ListItems.Count = 0 Then
'        mobjMsg.Abrir "Nenhum lançamento encontrado na LM", Ok, informacao, "Atenção"
'        ListView4.SetFocus
'        Exit Function
'    End If
    ValidaCampo = True
End Function

Private Function ValidaCampo2()
    ValidaCampo2 = False
    If Formula = "" Then
        mobjMsg.Abrir "Produto selecionado não possui FÓRMULA cadastrada. Não poder ser adicionado na lista", Ok, critico, "Atenção"
        Exit Function
    End If
    If txtcadastro(16).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(16).Tag, Ok, critico, "Atenção"
        Me.Combo1.SetFocus
        Exit Function
    End If
    If txtcadastro(19).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(19).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(19).SetFocus
        Exit Function
    End If
    If txtcadastro(20).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(20).Tag, vbInformation, "Atenção"
        Me.txtcadastro(20).SetFocus
        Exit Function
    End If
    'If txtcadastro(22).Text = "" Then
    '    Msgbox "Favor informar o campo " & Me.txtcadastro(22).Tag, vbInformation, "Atenção"
    '    Me.txtcadastro(22).SetFocus
    '    Exit Function
    'End If
    If optCadastro(0).Value = True Then
        If txtcadastro(24).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(24).Tag, Ok, critico, "Atenção"
            Me.txtcadastro(24).SetFocus
            Exit Function
        End If
    Else
        If txtcadastro(23).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(23).Tag, Ok, critico, "Atenção"
            Me.txtcadastro(23).SetFocus
            Exit Function
        End If
    End If
    
    If txtcadastro(25).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(25).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(25).SetFocus
        Exit Function
    End If
    If txtcadastro(26).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(26).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(26).SetFocus
        Exit Function
    End If
    If txtcadastro(27).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(27).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(27).SetFocus
        Exit Function
    End If
    If txtcadastro(29).Text = "" Or txtcadastro(29).Text = "0" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(29).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(29).SetFocus
        Exit Function
    End If
    If txtcadastro(28).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(28).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(28).SetFocus
        Exit Function
    End If
    If optCadastro(0).Value = True And txtcadastro(24).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(24).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(24).SetFocus
        Exit Function
    End If
    If optCadastro(1).Value = True And txtcadastro(23).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtcadastro(23).Tag, Ok, critico, "Atenção"
        Me.txtcadastro(23).SetFocus
        Exit Function
    End If
    ValidaCampo2 = True
End Function

Private Sub CapVar()
    vAr0(Conta + 1) = Val(Mid$(txtcadastro(24), X + 2, Len(txtcadastro(24)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtcadastro(24))
    Conta = Conta + 1
End Sub

Private Sub CapVar2()
    vAr0(Conta + 1) = Val(Mid$(txtcadastro(34), X + 2, Len(txtcadastro(34)) - X))
    Text2.Text = vAr0(Conta + 1)
    Y = X
    X = Len(txtcadastro(34))
    Conta = Conta + 1
End Sub

Private Sub SeparaDados()
    Dim RECEBE As String
    Dim Contador As Integer
    Contador = 0
    For X = 1 To Len(txtcadastro(24))
        If Mid(txtcadastro(24), X, 1) = "-" Then
            vAr0(Contador + 1) = RECEBE
            Contador = Contador + 1
            RECEBE = ""
        Else
            RECEBE = RECEBE & Mid(txtcadastro(24), X, 1)
        End If
    Next
    vAr0(Contador + 1) = Val(RECEBE)
End Sub

Private Sub LimpaControles()
    txtcadastro(15) = ""
    txtcadastro(17) = ""
    txtcadastro(18) = ""
    txtcadastro(19) = ""
    txtcadastro(21) = ""
    txtcadastro(23) = ""
    txtcadastro(24) = ""
    txtcadastro(29) = ""
    Formula = ""
    ForPint = ""
    Conta = 0
    txtcadastro(26).SetFocus
End Sub

Private Sub SomaListview()
    Dim SomaT As Currency, SomaP As Currency
    Dim i As Integer
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).SubItems(16) <> "pç" And ListView2.ListItems(i).SubItems(16) <> "PÇ" Then
            If ListView2.ListItems(i).SubItems(14) <> "" Then SomaT = SomaT + CCur(ListView2.ListItems(i).SubItems(14)) 'coluna de valores
            If ListView2.ListItems(i).SubItems(17) <> "" Then SomaP = SomaP + CCur(ListView2.ListItems(i).SubItems(17)) 'coluna de valores
        End If
    Next
    lblTotal.Caption = Format(SomaT, "#,##0.000;(#,##0.000)") 'Format(SomaTotal, "#,##0.000000000;(#,##0.000000000)")
    lblTotPint.Caption = Format(SomaP, "#,##0.000;(#,##0.000)")
End Sub

Private Sub SomaPesoSelecionado()
    Dim SomaPeso As Currency
    Dim SomaQtd As Currency
    Dim i As Integer
    SomaPeso = 0
    SomaQtd = 0
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems.Item(i).Selected = True Then
            ListView2.ListItems.Item(i).Selected = True
            SomaPeso = SomaPeso + CCur(ListView2.SelectedItem.ListSubItems.Item(14))
            SomaQtd = SomaQtd + CCur(ListView2.SelectedItem.ListSubItems.Item(11))
        End If
    Next
    'If ListView5.ListItems.Count = 0 Then SomaPeso = 1
    'If ListView5.ListItems.Count = 0 Then SomaQtd = 1
    Label38.Caption = Format(SomaPeso, "#,##0.000;(#,##0.000)")
    Label39.Caption = Format(SomaQtd, "#,##0.000;(#,##0.000)")
End Sub

Private Sub SomaQtdCJ()
    Dim SomaCJ As Integer
    Dim i As Integer
    SomaCJ = 1
    For i = 1 To ListView5.ListItems.Count
        ListView5.ListItems.Item(i).Selected = True
        SomaCJ = SomaCJ * Val(ListView5.SelectedItem.ListSubItems.Item(5))
    Next
    If ListView5.ListItems.Count = 0 Then SomaCJ = 1
    txtcadastro(16).Text = SomaCJ
End Sub

Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView3.SelectedItem Is Nothing Then
        ListView3.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 10 To 10
            leftPos = ListView3.Left + ListView3.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView3.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView3.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView3.SelectedItem.Text
                    Else
                        .Text = ListView3.SelectedItem.SubItems(i - 1)
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

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fired when the listitem is already selected, for this reason can't used mousedown event
'so we know which row is clicked, for the column, we need to translate the x to listview coordinate
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView2.SelectedItem Is Nothing Then
        ListView2.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 1 To 6
            'ListView2.LabelEdit = lvwAutomatic
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
                    .SetFocus
                End With
                Exit For
            End If
        Next i

'..............................
        ListView2.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        
        For i = 19 To 19
            'ListView2.LabelEdit = lvwAutomatic
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
                    .SetFocus
                End With
                Exit For
            End If
        Next i
'..............................

        For i = 25 To 25
            'ListView2.LabelEdit = lvwAutomatic
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
                    .SetFocus
                End With
                Exit For
            End If
        Next i

'..............................
    
    End If
End If
End Sub

Function GetLvwDeltaX() As Single
'returns deltaX, the scroll distance in pixels relative to ListView2.left, how much we scroll right
'si.npage propotional to both the width of the scroll box and ListView2.width
'si.npos is the scrolling position, which is propotional to deltaX

Dim si As SCROLLINFO, maxScrollPos As Long
Dim lvwCol As ColumnHeader, actualLvwWidth As Single
   
    If SSTab1.Tab = 1 Then
        Set lvwCol = ListView2.ColumnHeaders(ListView2.ColumnHeaders.Count)
        actualLvwWidth = lvwCol.Left + lvwCol.Width
    
        'PrintLvwColInfo
        si.cbSize = 28 '7 long vars x 4 bytes
        si.fMask = SIF_ALL
        GetScrollInfo ListView2.HWnd, SB_HORZ, si
        maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
        '58 is some constant to get things just right
        If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView2.Width + 58)
    ElseIf SSTab1.Tab = 2 Then
        Set lvwCol = ListView3.ColumnHeaders(ListView3.ColumnHeaders.Count)
        actualLvwWidth = lvwCol.Left + lvwCol.Width
    
        'PrintLvwColInfo
        si.cbSize = 28 '7 long vars x 4 bytes
        si.fMask = SIF_ALL
        GetScrollInfo ListView3.HWnd, SB_HORZ, si
        maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
        '58 is some constant to get things just right
        If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView3.Width + 58)
    End If
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
'called from ListView2 mouseup and subclass scroll events
'constants used are determined by trial & error, these are mainly the various widths and heights
'of edges in the classical windows. these constants may not be correct for other windows styles.
Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
If SSTab1.Tab = 2 Then
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
            If txtLeft < 11000 Then .Left = txtLeft + 50 Else .Left = txtLeft - 140
            .Top = txtTop
            .Width = txtWidth
            .Height = ListView2.SelectedItem.Height - 8
        End With
    End If
ElseIf SSTab1.Tab = 3 Then
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView3.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView3.Left + lvwCol.Left + 48 - dx + 80 'Determina inicio da coluna de Observação do Resumo da LM
        If txtLeft < ListView3.Left Then txtLeft = ListView3.Left + 48
    
        txtRightMax = ListView3.Left + ListView3.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView3.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView3.Left + ListView3.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
        'If txtRight > txtLeft Then txtWidth = txtRight - txtLeft Else txtLeft = -1000
    
        txtTopMin = ListView3.Top
        If Not ListView3.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView3.Top + ListView3.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView3.Top + ListView3.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
        With txtLvw '.move produces runtimez' error with -ve values
            .Left = txtLeft
            .Top = txtTop
            .Width = txtWidth
            .Height = ListView3.SelectedItem.Height - 8
        End With
    End If
End If
End Sub

Private Sub txtcadastro_KeyPress(Index As Integer, KeyAscii As Integer)
    'Para essa linha de comando existe um função dentro do módulo RotinaGeral
    'responsavel por desabilitar o BIP qdo precionada a tecla ENTER nos Texbox
    KeyAscii = Enter(KeyAscii)
    '-----------------
    If Index = 16 Or Index = 29 Or Index = 22 Or Index = 23 Then
        'aceitar somente números e "Back Space", "Enter", "virgula"
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    End If
    If Index = 24 Then
        'aceitar somente números e "Back Space", "Enter", "virgula"
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 And KeyAscii <> 45 Then
            KeyAscii = 0
        End If
    End If
    If Index = 27 Then
        If Len(txtcadastro(27)) > 1 Then
            KeyAscii = 8
        End If
    End If

    If Index = 38 Then
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtcadastro_LostFocus(Index As Integer)
    voltaCorText txtcadastro(Index)
    If Index = 18 Then
        If Val(txtcadastro(19)) <> 0 Then
            CarregaDados (19)
            'txtcadastro(18).Enabled = True
        End If
    End If
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView2 is edited
    Select Case KeyAscii
        Case 13 'enter key
            KeyAscii = 0
            txtLvw_LostFocus
        'other keys can be used for navigation
    End Select
End Sub

Private Sub txtLvw_LostFocus()
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        If SSTab1.Tab = 2 Then
            ListView2.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        ElseIf SSTab1.Tab = 3 Then
            ListView3.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        End If
        'add text entry to the last row
        'If ListView2.ListItems(ListView2.ListItems.Count) <> c_EntryTxt Then ListView2.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        If SSTab1.Tab = 2 Then
            ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
        ElseIf SSTab1.Tab = 3 Then
            ListView3.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
        End If
    End If
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView2's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView2.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub SalvaXLS()
'On Error GoTo testa_erro
    'If Text5.Text = "" Then
    '    Msgbox "Os dados do orçamento devem ser informados"
    '    Exit Sub
    'End If
    CaminhoArquivo = ""
    NomeArquivo = ""
    CaminhoArquivo = pathArq 'Mid$(frmConfiguracao.txtCaminho, 1, Len(frmConfiguracao.txtCaminho) - Len("contratoNOVO.mdb"))
    NomeArquivo = frmLM.Label32.Caption & " - " & frmLM.Label2.Caption & ".xls"
    
    cdg.Filter = "Planilha do Excel (*.xls)|*.xls"
    cdg.flags = cdlOFNHideReadOnly
    cdg.InitDir = CaminhoArquivo
    cdg.FileName = NomeArquivo
    pathArq = cdg.FileName
    cdg.ShowSave
    If Trim(pathArq) <> "" Then
        ExportaExcel
    End If
    Exit Sub
testa_erro:
    If Err.Number = 32755 Then
        mobjMsg.Abrir "Procedimento cancelado", Ok, critico, "Atenção"
    End If
End Sub

Private Sub ExportaExcel()
On Error GoTo Err
    Dim j As Integer, K As Integer, L As Integer
    'Dim Plan As Object 'Aplicação Excel
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")

    'PLANILHA DE LISTA DE MATERIAIS
    'CHAMA EXCEL / IMPRIME
    '**********************************************************************
    Plan.Workbooks.Open App.Path & "\LM - Padrao.xls"
    Plan.Visible = True
    Plan.UserControl = False

    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 15
    Me.ListView2.SortOrder = lvwAscending
    
    'PREENCHE CÉLULAS DESEJADAS
    '**********************************************************************
    Y = ListView2.ListItems.Count
    'linha1 = 27
    With Plan
        .Range("Y" & 1).Value = Label32 ' numero FCE
        .Range("Y" & 2).Value = Label2 ' numero LM
        .Range("D" & 3).Value = Text7 ' Descricao LM
        .Range("Y" & 3).Value = DTPicker1 ' Data LM
        
        .Range("Resumo!Q" & 1).Value = Label32 ' numero FCE
        .Range("Resumo!Q" & 2).Value = Label2 ' numero LM
        .Range("Resumo!C" & 3).Value = Text7 ' Descricao LM
        .Range("Resumo!Q" & 3).Value = DTPicker1 ' Data LM
        .Range("Resumo!B" & 4).Value = txtcadastro(1) ' Nome Cliente
        .Range("Resumo!E" & 4).Value = txtcadastro(7) ' Fone Cliente
        .Range("Resumo!G" & 4).Value = txtcadastro(9) ' Email Cliente
        .Range("Resumo!B" & 5).Value = txtcadastro(12) ' Nome Contato
        .Range("Resumo!E" & 5).Value = txtcadastro(13) ' Fone Contato
        .Range("Resumo!G" & 5).Value = txtcadastro(14) ' Email Contato
    End With
    
    'Dados dos Projetos e Desenhos de Fabricação
    Dim rsDesenhosFab As New ADODB.Recordset
    Dim SqlDesenhosFab As String
    
    'Dados se Desenhos de Conjunto
    Dim rsDesenhosCJ As New ADODB.Recordset
    Dim SqlDesenhosCJ As String
    
    
    j = 7
    X = 1
    Dim valor1 As Double, valor2 As Double, valor3 As Double, valor4 As Double, valor5 As Double, QtdTotCJ As Double
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        With Plan
            If j Mod 2 = 0 Then
                preencheVermelho j
            Else
                preencheBranco j
            End If
            'Dados Desenhos de Fabricação
            SqlDesenhosFab = "select b.projeto, b.descricao as [Desc Projeto],a.descricao as [Desc Desenho] from tbDesenhos as a inner join tbProjetos as b on a.codprojeto = b.codprojeto where a.iddesenho = '" & Val(ListView2.SelectedItem.ListSubItems.Item(2)) & "'"
            rsDesenhosFab.Open SqlDesenhosFab, cnBanco, adOpenKeyset, adLockReadOnly
            
            .Range("A" & j).Value = ListView2.SelectedItem.ListSubItems.Item(15) ' Item Sequêncial LM
            If rsDesenhosFab.RecordCount > 0 Then
                .Range("B" & j).Value = rsDesenhosFab.Fields(0) ' Projeto nº
                .Range("C" & j).Value = rsDesenhosFab.Fields(1) ' Projeto descrição
            End If
            contornoEVermelho j, "A"
            contornoDVermelho j, "C"
            
            SqlDesenhosCJ = "select a.idConjunto,a.codlm,a.idseq,b.desenho,a.quantidade,a.posicao,b.descricao from tbDesConjunto as a inner join tbdesenhos as b on a.iddesenho = b.iddesenho inner join tbProjetos as c on b.codprojeto = c.codprojeto where c.fce = '" & Label32.Caption & "' and a.idConjunto = '" & Val(ListView2.SelectedItem.ListSubItems.Item(23)) & "' order by idseq"
            rsDesenhosCJ.Open SqlDesenhosCJ, cnBanco, adOpenKeyset, adLockReadOnly
            'L = rsDesenhosCJ.ListItems.Count
            QtdTotCJ = 1
            While Not rsDesenhosCJ.EOF
                'Dados Desenhos de Conjunto
                'With Plan
                    .Range("D" & j).Value = rsDesenhosCJ.Fields(3) ' Desenho
                    .Range("E" & j).Value = rsDesenhosCJ.Fields(5) ' Item
                    .Range("F" & j).Value = rsDesenhosCJ.Fields(4) ' Quantidade
                    .Range("G" & j).Value = rsDesenhosCJ.Fields(6) ' Descrição
                    contornoEVermelho j, "D"
                    contornoDVermelho j, "G"
                    contornoEVermelho j, "H"
                    contornoDVermelho j, "K"
                    contornoEVermelho j, "L"
                    contornoDVermelho j, "Z"
                    QtdTotCJ = QtdTotCJ * rsDesenhosCJ.Fields(4)
                    rsDesenhosCJ.MoveNext
                    If Not rsDesenhosCJ.EOF Then j = j + 1
                    If j Mod 2 = 0 Then
                        preencheVermelho j
                    Else
                        preencheBranco j
                    End If
                'End With
            Wend
            rsDesenhosCJ.Close
            contornoEVermelho j, "D"
            contornoDVermelho j, "G"
            
            valor1 = ListView2.SelectedItem.ListSubItems.Item(10)
            valor2 = ListView2.SelectedItem.ListSubItems.Item(12)
            valor3 = ListView2.SelectedItem.ListSubItems.Item(14)
            If ListView2.SelectedItem.ListSubItems.Item(17) <> "" Then valor4 = ListView2.SelectedItem.ListSubItems.Item(17)
            valor5 = 0
            If Mid$(ListView2.SelectedItem.ListSubItems.Item(7), 1, 5) <> "CHAPA" And Mid$(ListView2.SelectedItem.ListSubItems.Item(7), 1, 5) <> "GRADE" And Val(ListView2.SelectedItem.ListSubItems.Item(6)) <> 0 Then
                If ListView2.SelectedItem.ListSubItems.Item(9) <> "" Then
                    valor5 = ListView2.SelectedItem.ListSubItems.Item(9) * valor1 * valor2
                Else
                    valor5 = ListView2.SelectedItem.ListSubItems.Item(14) / ListView2.SelectedItem.ListSubItems.Item(19)
                End If
            End If
            
            .Range("H" & j).Value = ListView2.ListItems.Item(X) ' Desenho
            .Range("I" & j).Value = ListView2.SelectedItem.ListSubItems.Item(5) ' Item
            .Range("J" & j).Value = ListView2.SelectedItem.ListSubItems.Item(10) ' Quantidade
            'If rsDesenhosFab.RecordCount > 0 Then
                '.Range("K" & j).Value = rsDesenhosFab.Fields(2) ' Desenho descrição
            'End If
            .Range("K" & j).Value = ListView2.SelectedItem.ListSubItems.Item(3) & " - " & ListView2.SelectedItem.ListSubItems.Item(4) 'rsDesenhosFab.Fields(2) ' Desenho descrição
            
            contornoEVermelho j, "H"
            contornoDVermelho j, "K"
'            .Range("L" & j).Value = QtdTotCJ * ListView2.SelectedItem.ListSubItems.Item(10) ' Quantidade Total CJ * Quantidade unit. Desenho
            .Range("L" & j).Value = ListView2.SelectedItem.ListSubItems.Item(12) * ListView2.SelectedItem.ListSubItems.Item(10) ' Quantidade Total CJ * Quantidade unit. Desenho
            .Range("M" & j).Value = ListView2.SelectedItem.ListSubItems.Item(7) ' Descrição produto
            .Range("N" & j).Value = ListView2.SelectedItem.ListSubItems.Item(9) ' Dimensão
            .Range("O" & j).Value = ListView2.SelectedItem.ListSubItems.Item(8) ' Tipo de Material
            .Range("R" & j).Value = ListView2.SelectedItem.ListSubItems.Item(11) 'Peso Unit./Qtd.
            .Range("S" & j).Value = ListView2.SelectedItem.ListSubItems.Item(14) 'Peso Total
            
            If InStr(1, ListView2.SelectedItem.ListSubItems.Item(9), "-") = 0 And ListView2.SelectedItem.ListSubItems.Item(9) <> "" Then
'                .Range("T" & j).Value = ListView2.SelectedItem.ListSubItems.Item(9) * QtdTotCJ 'Peso Total
                .Range("T" & j).Value = ListView2.SelectedItem.ListSubItems.Item(9) * ListView2.SelectedItem.ListSubItems.Item(12) 'Peso Total
            Else
            End If
            
            .Range("U" & j).Value = ListView2.SelectedItem.ListSubItems.Item(17) * 2 'Área de jato
            .Range("V" & j).Value = ListView2.SelectedItem.ListSubItems.Item(17) 'Área de pintura
            contornoEVermelho j, "L"
            contornoDVermelho j, "Z"
            rsDesenhosFab.Close
            j = j + 1
            tracejarVermelho j
        End With
    Next
    contornoBVermelho j
    Plan.Range("A1").Select

    'PLANILHA DE RESUMO
    'Daki pra baixo eh referente ao resumo de materiais
    
    Dim compara As String
    valor1 = 0
    valor2 = 0
    valor3 = 0
    valor4 = 0
    j = 9
    Y = ListView3.ListItems.Count
    For X = 1 To Y
        ListView3.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
            With Plan
                .Range("Resumo!A" & j).Value = Mid$(ListView3.SelectedItem.ListSubItems.Item(1), 1, 11) 'Código do Produto
                .Range("Resumo!B" & j).Value = ListView3.SelectedItem.ListSubItems.Item(2) 'Descrição do Produto
                .Range("Resumo!C" & j).Value = Mid$(ListView3.SelectedItem.ListSubItems.Item(3), 8, 20) 'Tipo de Material/Norma
                '.Range("Resumo!E" & j).Value = ListView3.SelectedItem.ListSubItems.Item(4) 'Unidade de Medida
                If ListView3.SelectedItem.ListSubItems.Item(4) <> "KG" And ListView3.SelectedItem.ListSubItems.Item(4) <> "kg" Then
                    .Range("Resumo!E" & j).Value = ListView3.SelectedItem.ListSubItems.Item(5) 'Peso Unitário total (p/ produto)
                Else
                    .Range("Resumo!D" & j).Value = ListView3.SelectedItem.ListSubItems.Item(5) 'Peso Unitário total (p/ produto)
                End If
                .Range("Resumo!H" & j).Value = ListView3.SelectedItem.ListSubItems.Item(6) 'Area de pintura total (p produto)
                If Val(ListView3.SelectedItem.ListSubItems.Item(7)) <> 0 Then .Range("Resumo!F" & j).Value = ListView3.SelectedItem.ListSubItems.Item(7) Else .Range("Resumo!F" & j).Value = "-" 'Comprimento
                '.Range("Resumo!I" & j).Value = ListView3.SelectedItem.ListSubItems.Item(8) 'Peso Específico
                .Range("Resumo!I" & j).Value = ListView3.SelectedItem.ListSubItems.Item(9) 'Observação
                j = j + 1
            End With
    Next
    
    Plan.Columns("A:R").EntireColumn.AutoFit 'Ajusta as colunas
    'FECHA REFERÊNCIA AOS OBJETOS
    '**********************************************************************
    Plan.ActiveWorkbook.SaveAs cdg.FileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    'Plan.Close
    Set Plan = Nothing
    
    mobjMsg.Abrir "Dados exportados com sucesso", Ok, informacao, "Atenção"
    Exit Sub
Err:
    If Err.Number = -2147467259 Then
        While reestabeleceConexao = False
        Wend
        Resume
    Else
        mobjMsg.Abrir "Ocorreu um erro, O MSOffice não esta instalado nesse computador!", Ok, critico, "Atenção"
        Exit Sub
    End If
End Sub

Private Sub preencheVermelho(posi As Integer)
    Plan.Range("A" & posi & ":Z" & posi).Select
    With Plan.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub preencheBranco(posi As Integer)
    Plan.Range("A" & posi & ":Z" & posi).Select
    With Plan.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub contornoDVermelho(vLin As Integer, vCol As String)
    Plan.Range(vCol & vLin).Select
    With Plan.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Private Sub contornoEVermelho(vLin As Integer, vCol As String)
    Plan.Range(vCol & vLin).Select
    With Plan.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Private Sub contornoBVermelho(vLin As Integer)
    Plan.Range("A" & vLin & ":Z" & vLin).Select
    With Plan.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub


Private Sub tracejarVermelho(vLin As Integer)
    Plan.Range("A" & vLin & ":Z" & vLin).Select
    With Plan.Selection.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ThemeColor = 6
        .TintAndShade = -0.249946592608417
        .Weight = xlThin
    End With
End Sub

Private Sub GerarResumo()
On Error GoTo Err
    Dim j As Integer
    Dim Plan As Object 'Aplicação Excel
    'INSTANCIA OBJETO EXCEL NA MEMÓRIA
    '**********************************************************************
    Set Plan = CreateObject("excel.application")
    
    Dim rsResumo As New ADODB.Recordset
    Dim SqlResumo As String
    Dim compara As String
    Dim ItemLst As ListItem
    Dim valor1 As Double, valor2 As Double, valor3 As Double, valor4 As Double, valor5 As Double, somaPL As Double, somaPM As Double
    valor1 = 0
    valor2 = 0
    valor3 = 0
    valor4 = 0
    valor5 = 0
    somaPL = 0
    somaPM = 0
    j = 3
    Y = ListView2.ListItems.Count
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 13
    Me.ListView2.SortOrder = lvwAscending
    ListView3.ListItems.Clear
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True 'Passar a selecao para o próximo item
        If ListView2.SelectedItem.ListSubItems.Item(13) <> compara Then
            With Plan
                j = j + 1
                valor1 = 0
                valor2 = 0
                valor3 = 0
                valor4 = 0
                valor5 = 0
'---------------- Rotina para somar comprimento
                If ListView2.SelectedItem.ListSubItems.Item(21) <> "" Then
                    valor5 = Format(ListView2.SelectedItem.ListSubItems.Item(21), "#,##0.000;(#,##0.000)")
                End If
'----------------
                
                If ListView2.SelectedItem.ListSubItems.Item(16) = "pç" Or ListView2.SelectedItem.ListSubItems.Item(16) = "PÇ" Or ListView2.SelectedItem.ListSubItems.Item(16) = "UN" Or ListView2.SelectedItem.ListSubItems.Item(16) = "un" Then
'                    valor3 = Format(ListView2.SelectedItem.ListSubItems.Item(10), "#,##0.000;(#,##0.000)") 'Quantidade
                    valor3 = Format(ListView2.SelectedItem.ListSubItems.Item(10) * ListView2.SelectedItem.ListSubItems.Item(12), "#,##0.000;(#,##0.000)") 'Quantidade
                Else
                    valor3 = Format(ListView2.SelectedItem.ListSubItems.Item(14), "#,##0.000;(#,##0.000)") 'Peso total
                End If
                
                'valor3 = Format(ListView2.SelectedItem.ListSubItems.Item(14), "#,##0.000;(#,##0.000)") 'Peso total
                If Format(ListView2.SelectedItem.ListSubItems.Item(17), "#,##0.000;(#,##0.000)") <> "" Then valor4 = Format(ListView2.SelectedItem.ListSubItems.Item(17), "#,##0.000;(#,##0.000)") 'Area de pintura
                Set ItemLst = ListView3.ListItems.Add(, , j - 3)
                ItemLst.SubItems(1) = ListView2.SelectedItem.ListSubItems.Item(6) 'Codigo
                ItemLst.SubItems(2) = ListView2.SelectedItem.ListSubItems.Item(7) 'Descricao
                ItemLst.SubItems(3) = ListView2.SelectedItem.ListSubItems.Item(8) 'Material
                ItemLst.SubItems(4) = ListView2.SelectedItem.ListSubItems.Item(16) 'UN
                ItemLst.SubItems(5) = Format(valor3, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(6) = Format(valor4, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(7) = Format(valor5, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(8) = ListView2.SelectedItem.ListSubItems.Item(19)
                ItemLst.SubItems(9) = "-"
             End With
        Else
            With Plan
                valor1 = ListView2.SelectedItem.ListSubItems.Item(10)
                valor2 = ListView2.SelectedItem.ListSubItems.Item(12)
                
                If ListView2.SelectedItem.ListSubItems.Item(16) = "pç" Or ListView2.SelectedItem.ListSubItems.Item(16) = "PÇ" Or ListView2.SelectedItem.ListSubItems.Item(16) = "UN" Or ListView2.SelectedItem.ListSubItems.Item(16) = "un" Then
'                    valor3 = Format(valor3 + ListView2.SelectedItem.ListSubItems.Item(10), "#,##0.000;(#,##0.000)") 'Quantidade
                    valor3 = Format(valor3 + (ListView2.SelectedItem.ListSubItems.Item(10) * ListView2.SelectedItem.ListSubItems.Item(12)), "#,##0.000;(#,##0.000)") 'Quantidade
                Else
                    valor3 = Format(valor3 + ListView2.SelectedItem.ListSubItems.Item(14), "#,##0.000;(#,##0.000)") 'Peso total
                End If
                
                'valor3 = Format(valor3 + ListView2.SelectedItem.ListSubItems.Item(14), "#,##0.000;(#,##0.000)") 'Peso total
                If ListView2.SelectedItem.ListSubItems.Item(17) <> "" Then valor4 = Format(valor4 + ListView2.SelectedItem.ListSubItems.Item(17), "#,##0.000;(#,##0.000)") 'Area de pintura
'---------------- Rotina para somar comprimento
                If ListView2.SelectedItem.ListSubItems.Item(21) <> "" Then
                    valor5 = Format(valor5 + ListView2.SelectedItem.ListSubItems.Item(21), "#,##0.000;(#,##0.000)")
                End If
'----------------
                ItemLst.SubItems(5) = Format(valor3, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(6) = Format(valor4, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(7) = Format(valor5, "#,##0.000;(#,##0.000)")
                ItemLst.SubItems(8) = ListView2.SelectedItem.ListSubItems.Item(19)
                ItemLst.SubItems(9) = "-"
            End With
        End If
                
        SqlResumo = "Select * from tbResumolm where tbResumolm.fce=" & " '" & Val(Label32) & "'" & _
        "and tbResumolm.codlm=" & " '" & Val(Label2) & "'" & "and tbResumolm.codmat = '" & Val(Mid$(ItemLst.SubItems(1), 15, 6)) & "'"
        rsResumo.Open SqlResumo, cnBanco, adOpenKeyset, adLockOptimistic
        
        If Not rsResumo.EOF Then
            ItemLst.SubItems(9) = rsResumo.Fields(4) 'Observação
        Else
            ItemLst.SubItems(9) = "-" 'Observação
        End If
        rsResumo.Close
        
        If ListView2.ListItems(X).SubItems(16) <> "pç" And ListView2.ListItems(X).SubItems(16) <> "PÇ" Then
            somaPL = somaPL + ListView2.SelectedItem.ListSubItems.Item(14)  'Peso total
            If ListView2.SelectedItem.ListSubItems.Item(17) <> "" Then somaPM = somaPM + ListView2.SelectedItem.ListSubItems.Item(17) 'Área de pintua
        End If
        compara = ListView2.SelectedItem.ListSubItems.Item(13) 'Código+material
        ItemLst.ListSubItems(9).Bold = True
    Next
    lbltotl = Format(somaPL, "#,##0.000;(#,##0.000)")
    lbltotpm = Format(somaPM, "#,##0.000;(#,##0.000)")
        
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 15
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

Private Sub RestauraItens()
    On Error GoTo Err
    Dim rsLisview As New ADODB.Recordset
    Dim ItemLst As ListItem
    Dim sql As String
    PesoTotal = 0
    sql = "select a.fce, a.codlm, a.codseq, c.desenho, c.revisao, b.CODIGOPRD + ' - ' + cast(a.codmat as varchar) as codmat, d.posicao, d.item, a.quantcj, a.quantunit, a.dimensoes, a.pesounit, a.area, b.CODTB2FAT, a.codfo, a.observação, d.descposicao, a.matncadast, b.NOMEFANTASIA, " & _
          "b.CODUNDCONTROLE, e.DESCRICAO, c.descricao, a.calcpor, a.idconjunto, c.iddesenho,d.pesoposicao from tbitemlm as a inner join " & vBancoTotvs & ".dbo.tprd as b on a.codmat = b.IDPRD inner join tbDesenhos as c on a.codigodes = c.iddesenho inner join tbPosicoes as d on a.codigopos = d.codigopos " & _
          "left join " & vBancoTotvs & ".dbo.TTB2 as e on b.CODTB2FAT = e.CODTB2FAT and e.codcoligada = 5" & _
          "where a.fce = '" & Val(Label32) & "' and a.codlm = '" & Val(Label2) & "' order by a.codseq"
    rsLisview.Open sql, cnBanco, adOpenKeyset, adLockOptimistic
    If rsLisview.RecordCount > 0 Then
        While Not rsLisview.EOF
            Set ItemLst = ListView2.ListItems.Add(, , rsLisview.Fields(3)) 'Desenho
            ItemLst.SubItems(1) = "" & rsLisview.Fields(4) ' Revisao
            ItemLst.SubItems(2) = "" & rsLisview.Fields(24) 'ID Desenho
            ItemLst.SubItems(3) = "" & rsLisview.Fields(6) 'Posição
            ItemLst.SubItems(4) = "" & rsLisview.Fields(16)  'Descricao Posição/Marca
            
            ItemLst.SubItems(5) = "" & rsLisview.Fields(7) 'Item
            ItemLst.SubItems(6) = "" & Format(rsLisview.Fields(5), "000000") 'Código
            If Val(rsLisview.Fields(5)) <> 0 Then
                ItemLst.SubItems(7) = "" & rsLisview.Fields(18) 'Descrição Material
            Else
                ItemLst.SubItems(7) = "" & rsLisview.Fields(17) 'Descrição Material
            End If
            
            If rsLisview.Fields(13) <> 0 Then ItemLst.SubItems(8) = "" & Format(rsLisview.Fields(13), "000000") & "-" & rsLisview.Fields(20) Else ItemLst.SubItems(8) = "-" 'Código do tipo+ descrição do material - Ex.: 000001 - ASTM A36 ' Tipo Material (codigo+descrição)
            ItemLst.SubItems(9) = "" & rsLisview.Fields(10) 'Dimensão
            ItemLst.SubItems(10) = "" & rsLisview.Fields(9) 'Q. Unit
            ItemLst.SubItems(11) = "" & Format(rsLisview.Fields(11), "#,##0.000;(#,##0.000)") 'Peso Unit/Qtd
            ItemLst.SubItems(12) = "" & rsLisview.Fields(8) 'Q CJ
            If ItemLst.SubItems(8) <> "-" Then ItemLst.SubItems(13) = ItemLst.SubItems(7) & ItemLst.SubItems(8) Else ItemLst.SubItems(13) = ItemLst.SubItems(7) 'codigo+material
            
            PesoTotal = Format(rsLisview.Fields(8) * rsLisview.Fields(9) * rsLisview.Fields(11), "#,##0.000;(#,##0.000")
'            PesoTotal = Format(rsLisview.Fields(9) * rsLisview.Fields(11), "#,##0.000;(#,##0.000")
            ItemLst.SubItems(14) = "" & Format(PesoTotal, "#,##0.000;(#,##0.000)") 'Peso total
            ItemLst.SubItems(15) = "" & Format(rsLisview.Fields(2), "0000") 'Sequência
            If Val(rsLisview.Fields(5)) <> 0 Then
                ItemLst.SubItems(16) = "" & rsLisview.Fields(19) 'UN
            Else
                ItemLst.SubItems(16) = "pç" 'UN
            End If
            ItemLst.SubItems(17) = "" & Format(rsLisview.Fields(12), "#,##0.000;(#,##0.000)") 'Área de Pintura
            If rsLisview.Fields(14) <> "Null" Then ItemLst.SubItems(20) = Format(rsLisview.Fields(14), "000000") Else ItemLst.SubItems(20) = "-" 'FO
            ItemLst.SubItems(18) = "" & rsLisview.Fields(15) 'Observação
            ItemLst.SubItems(22) = "" & rsLisview.Fields(22) 'Calculador por
            If Not IsNull(rsLisview.Fields(23)) Then ItemLst.SubItems(23) = "" & rsLisview.Fields(23) Else ItemLst.SubItems(23) = "-" 'ID Conjunto
            
            If IsNull(rsLisview.Fields(25)) Then ItemLst.SubItems(24) = "0" Else ItemLst.SubItems(24) = rsLisview.Fields(25)  'Peso Posicao
            
            'Deixar a coluna de OBSERVAÇÃO em negrito e vermelho
            ItemLst.ListSubItems(18).Bold = True
            'ItemLst.ListSubItems(18).ForeColor = vbRed
            ItemLst.ForeColor = &H404080
            ItemLst.ListSubItems(1).ForeColor = &H404080
            ItemLst.ListSubItems(2).ForeColor = &H404080
            ItemLst.ListSubItems(3).ForeColor = &H404080
            ItemLst.ListSubItems(4).ForeColor = &H404080
            ItemLst.ListSubItems(5).ForeColor = &H404080
            
            '--------------------
'----------
            Dim rsConstantes As New ADODB.Recordset
            Dim SqlConstantes As String
            SqlConstantes = "Select * from tbconstantes as a where a.idprd = '" & Val(Mid$(rsLisview.Fields(5), 15, 6)) & "'order by a.idprd"
            rsConstantes.Open SqlConstantes, cnBanco, adOpenKeyset, adLockOptimistic
            If rsConstantes.RecordCount > 0 Then
                If Val(rsLisview.Fields(5)) <> 0 Then ItemLst.SubItems(19) = rsConstantes.Fields(1) 'Peso Especifico
            End If
'----------
            If Val(ItemLst.SubItems(6)) <> 0 Then
                
                If Mid$(ItemLst.SubItems(7), 1, 5) <> "CHAPA" And Mid$(ItemLst.SubItems(7), 1, 5) <> "GRADE" And Val(ItemLst.SubItems(6)) <> 0 Then
'                If Mid$(ItemLst.SubItems(7), 1, 5) <> "CHAPA" And Val(ListView2.SelectedItem.ListSubItems.Item(6)) <> 0 Then
                    If ItemLst.SubItems(9) <> "" And ItemLst.SubItems(9) <> "-" Then
                        ItemLst.SubItems(21) = ItemLst.SubItems(9) * ItemLst.SubItems(10) * ItemLst.SubItems(12) 'Comprimento
                    Else
                        ItemLst.SubItems(21) = ItemLst.SubItems(14) / ItemLst.SubItems(19) * 1000 'Comprimento
                    End If
                Else
                    ItemLst.SubItems(21) = 0 'Comprimento
                End If
            Else
                ItemLst.SubItems(21) = 0 'Comprimento
            End If
'----------
            
            rsConstantes.Close
            Set rsConstantes = Nothing
            
            rsLisview.MoveNext
        Wend
    End If
    SomaListview
    Me.ListView2.ColumnHeaders(11).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(12).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(15).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(18).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(20).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(22).Alignment = lvwColumnRight
    
    rsLisview.Close
    ListView1.Refresh
    Set rsLisview = Nothing
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

Private Sub txtcadastro_GotFocus(Index As Integer)
On Error Resume Next
    mudaCorText txtcadastro(Index)
    'Abaixo - Deixa selecionado todo o texto do TextBox
    Dim X As Integer
    For X = 1 To txtcadastro.Count - 1
        txtcadastro(X).SelStart = 0
        txtcadastro(X).SelLength = Len(txtcadastro(X).Text)
    Next
End Sub

Private Sub criaTabelaExcLM()
'    cnBanco.Execute "CREATE TABLE " & sDatabaseName & ".dbo.tbExcluidosLM" & vTime & "(" & _
'    "fce NUMERIC NOT NULL," & _
'    "codlm NUMERIC NOT NULL," & _
'    "codseq NUMERIC NOT NULL)"
End Sub

Private Sub excluiTabelaExcLM()
'On Error GoTo Err
'    Dim rsExcluirTb As New ADODB.Recordset
'    Dim SqlExcluirTb As String
'    SqlExcluirTb = "Drop table tbExcluidosLM" & vTime
'    rsExcluirTb.Open SqlExcluirTb, cnBanco
'    Exit Sub
'Err:
'    If Err.Number = -2147467259 Then
'        While reestabeleceConexao = False
'        Wend
'        Resume
'    Else
'        Resume Next
'    End If
End Sub


