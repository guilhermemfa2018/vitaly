VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{34AD7171-8984-11D8-AD7F-BE723A6C8E7C}#1.0#0"; "IpToolTips.ocx"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCandidatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de candidatos"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "frmCandidatos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   15
      Left            =   720
      TabIndex        =   102
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   8760
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
      MICON           =   "frmCandidatos.frx":0CCA
      PICN            =   "frmCandidatos.frx":0CE6
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
      Index           =   14
      Left            =   120
      TabIndex        =   101
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   8760
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
      MICON           =   "frmCandidatos.frx":19C0
      PICN            =   "frmCandidatos.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame14 
      Caption         =   "ID"
      Height          =   735
      Left            =   8760
      TabIndex        =   98
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lblID 
         Height          =   255
         Left            =   120
         TabIndex        =   99
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
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
      Left            =   1920
      TabIndex        =   96
      Top             =   8760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   3855
      Index           =   0
      Left            =   8760
      TabIndex        =   92
      Top             =   120
      Width           =   2535
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   13
         Left            =   720
         TabIndex        =   104
         Tag             =   "Excluir foto"
         ToolTipText     =   "Excluir foto"
         Top             =   3120
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
         MICON           =   "frmCandidatos.frx":26B6
         PICN            =   "frmCandidatos.frx":26D2
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
         Index           =   12
         Left            =   120
         TabIndex        =   103
         Tag             =   "Adicionar foto"
         ToolTipText     =   "Adicionar foto"
         Top             =   3120
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
         MICON           =   "frmCandidatos.frx":33AC
         PICN            =   "frmCandidatos.frx":33C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture2 
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2715
         ScaleWidth      =   2235
         TabIndex        =   93
         Top             =   240
         Width           =   2295
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2775
            Left            =   0
            Top             =   0
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   4895
            Image           =   "frmCandidatos.frx":40A2
         End
         Begin VB.Label Label59 
            Alignment       =   2  'Center
            Caption         =   "A Imagem não se encontra no local especificado"
            Height          =   495
            Left            =   120
            TabIndex        =   94
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
   Begin VB.Frame Frame11 
      Caption         =   "Identificação "
      Height          =   3855
      Left            =   120
      TabIndex        =   88
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame8 
         Caption         =   "Módulo de avaliação "
         Height          =   2175
         Left            =   120
         TabIndex        =   151
         Top             =   1560
         Width           =   8295
         Begin MAESTRO.chameleonButton chameleonButton1 
            Height          =   375
            Left            =   360
            TabIndex        =   164
            Top             =   1440
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Avaliar"
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
            MICON           =   "frmCandidatos.frx":40BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   7680
            Top             =   1560
         End
         Begin VB.CheckBox chkAvaliador 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Formação escolar:"
            ForeColor       =   &H80000001&
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   159
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CheckBox chkAvaliador 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cursos/treinamentos:"
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
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   158
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chkAvaliador 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Habilidades:"
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
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   157
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox chkAvaliador 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Experiência:"
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
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   156
            Top             =   480
            Width           =   1215
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFFF&
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
            Left            =   240
            TabIndex        =   154
            Top             =   480
            Width           =   1335
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Left            =   60
               TabIndex        =   155
               Top             =   270
               Width           =   1215
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frame9"
            Height          =   975
            Left            =   4560
            TabIndex        =   152
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
            Begin SHDocVwCtl.WebBrowser WebBrowser1 
               Height          =   1215
               Left            =   -120
               TabIndex        =   153
               Top             =   -120
               Visible         =   0   'False
               Width           =   1695
               ExtentX         =   2990
               ExtentY         =   2143
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   ""
            End
         End
         Begin VB.Label Label39 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3735
            TabIndex        =   160
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label Label40 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3735
            TabIndex        =   161
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label Label38 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3735
            TabIndex        =   162
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   3735
            TabIndex        =   163
            Top             =   480
            Width           =   840
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H8000000D&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   1815
            Left            =   120
            Top             =   240
            Width           =   8055
         End
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
            Height          =   1575
            Left            =   5880
            Top             =   240
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   2778
            Image           =   "frmCandidatos.frx":40D6
            Opacity         =   30
            Props           =   5
         End
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   3
         Left            =   3840
         TabIndex        =   2
         Tag             =   "Nome do colaborador"
         ToolTipText     =   "Nome do colaborador"
         Top             =   480
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Tag             =   "Data de cadastro do colaborador"
         ToolTipText     =   "Data de cadastro do colaborador"
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   40505
      End
      Begin MSMask.MaskEdBox mskCadMatriz 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Tag             =   "CPF do colaborado"
         ToolTipText     =   "CPF do colaborado"
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###.###.###-##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCadMatriz 
         BackColor       =   &H8000000F&
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
         Height          =   285
         Index           =   4
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   89
         Tag             =   "Matriz e cargo do colaborador"
         ToolTipText     =   "Matriz e cargo do colaborador"
         Top             =   1080
         Width           =   8175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmCandidatos.frx":280F9
         TabIndex        =   113
         Top             =   840
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "frmCandidatos.frx":28171
         TabIndex        =   112
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   3840
         OleObjectBlob   =   "frmCandidatos.frx":281F1
         TabIndex        =   111
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmCandidatos.frx":28259
         TabIndex        =   110
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Determinações de competência "
      Height          =   4575
      Left            =   120
      TabIndex        =   76
      Top             =   4080
      Width           =   11175
      Begin TabDlg.SSTab SSTab1 
         Height          =   4215
         Left            =   120
         TabIndex        =   77
         Tag             =   "Histórico"
         Top             =   240
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7435
         _Version        =   393216
         Tabs            =   7
         Tab             =   3
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "Dados pessoais"
         TabPicture(0)   =   "frmCandidatos.frx":282C5
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(2)=   "Frame12"
         Tab(0).Control(3)=   "Frame3"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Experiências"
         TabPicture(1)   =   "frmCandidatos.frx":282E1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SkinLabel8"
         Tab(1).Control(1)=   "SkinLabel7"
         Tab(1).Control(2)=   "txtCadMatriz(9)"
         Tab(1).Control(3)=   "txtCadMatriz(8)"
         Tab(1).Control(4)=   "SkinLabel5"
         Tab(1).Control(5)=   "SkinLabel6"
         Tab(1).Control(6)=   "cmdCadastro(24)"
         Tab(1).Control(7)=   "cmdCad(0)"
         Tab(1).Control(8)=   "cmdCadastro(3)"
         Tab(1).Control(9)=   "cmdCadastro(2)"
         Tab(1).Control(10)=   "cmdCadastro(1)"
         Tab(1).Control(11)=   "cmdCadastro(0)"
         Tab(1).Control(12)=   "txtCadMatriz(1)"
         Tab(1).Control(13)=   "cboCadMatriz(2)"
         Tab(1).Control(14)=   "cboCadMatriz(3)"
         Tab(1).Control(15)=   "ListView1"
         Tab(1).ControlCount=   16
         TabCaption(2)   =   "Habilidades"
         TabPicture(2)   =   "frmCandidatos.frx":282FD
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListView2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Cursos"
         TabPicture(3)   =   "frmCandidatos.frx":28319
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "ListView3"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "cmdCadastro(8)"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "cmdCadastro(7)"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "cmdCadastro(5)"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).Control(4)=   "cmdCad(1)"
         Tab(3).Control(4).Enabled=   0   'False
         Tab(3).Control(5)=   "SkinLabel20"
         Tab(3).Control(5).Enabled=   0   'False
         Tab(3).Control(6)=   "SkinLabel21"
         Tab(3).Control(6).Enabled=   0   'False
         Tab(3).Control(7)=   "SkinLabel22"
         Tab(3).Control(7).Enabled=   0   'False
         Tab(3).Control(8)=   "cboCadMatriz(5)"
         Tab(3).Control(8).Enabled=   0   'False
         Tab(3).Control(9)=   "txtCadMatriz(11)"
         Tab(3).Control(9).Enabled=   0   'False
         Tab(3).Control(10)=   "txtCadMatriz(10)"
         Tab(3).Control(10).Enabled=   0   'False
         Tab(3).Control(11)=   "cmdCadastro(25)"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).ControlCount=   12
         TabCaption(4)   =   "Graduação"
         TabPicture(4)   =   "frmCandidatos.frx":28335
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "ListView4"
         Tab(4).Control(1)=   "cmdCadastro(16)"
         Tab(4).Control(2)=   "cmdCadastro(10)"
         Tab(4).Control(3)=   "cmdCadastro(9)"
         Tab(4).Control(4)=   "cmdCadastro(6)"
         Tab(4).Control(5)=   "cmdCad(2)"
         Tab(4).Control(6)=   "SkinLabel23"
         Tab(4).Control(7)=   "SkinLabel24"
         Tab(4).Control(8)=   "txtCadMatriz(12)"
         Tab(4).Control(9)=   "txtCadMatriz(13)"
         Tab(4).Control(10)=   "cmdCadastro(26)"
         Tab(4).ControlCount=   11
         TabCaption(5)   =   "Cargos Pret."
         TabPicture(5)   =   "frmCandidatos.frx":28351
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "txtCadMatriz(20)"
         Tab(5).Control(1)=   "txtCadMatriz(6)"
         Tab(5).Control(2)=   "txtCadMatriz(5)"
         Tab(5).Control(3)=   "txtCadMatriz(0)"
         Tab(5).Control(4)=   "SkinLabel28"
         Tab(5).Control(5)=   "SkinLabel27"
         Tab(5).Control(6)=   "SkinLabel26"
         Tab(5).Control(7)=   "SkinLabel25"
         Tab(5).Control(8)=   "cmdCad(3)"
         Tab(5).Control(9)=   "cmdCadastro(19)"
         Tab(5).Control(10)=   "cmdCadastro(20)"
         Tab(5).Control(11)=   "cmdCadastro(21)"
         Tab(5).Control(12)=   "cmdCadastro(22)"
         Tab(5).Control(13)=   "ListView5"
         Tab(5).ControlCount=   14
         TabCaption(6)   =   "Integração"
         TabPicture(6)   =   "frmCandidatos.frx":2836D
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "SSTab2"
         Tab(6).ControlCount=   1
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   26
            Left            =   -64920
            TabIndex        =   166
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":28389
            PICN            =   "frmCandidatos.frx":283A5
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
            Index           =   25
            Left            =   10080
            TabIndex        =   165
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2907F
            PICN            =   "frmCandidatos.frx":2909B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   20
            Left            =   -64920
            TabIndex        =   44
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   -71640
            TabIndex        =   43
            Tag             =   "Nome do cargo"
            ToolTipText     =   "Nome do cargo"
            Top             =   720
            Width           =   6615
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -72840
            TabIndex        =   38
            Tag             =   "Código do cargo"
            ToolTipText     =   "Código do cargo"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCadMatriz 
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
            Index           =   0
            Left            =   -74880
            TabIndex        =   37
            Tag             =   "Número do matriz"
            ToolTipText     =   "Número do matriz"
            Top             =   720
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
            Height          =   255
            Left            =   -64920
            OleObjectBlob   =   "frmCandidatos.frx":29D75
            TabIndex        =   137
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
            Height          =   255
            Left            =   -71640
            OleObjectBlob   =   "frmCandidatos.frx":29DDF
            TabIndex        =   136
            Top             =   480
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
            Height          =   255
            Left            =   -72840
            OleObjectBlob   =   "frmCandidatos.frx":29E53
            TabIndex        =   135
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmCandidatos.frx":29ECB
            TabIndex        =   134
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -73680
            TabIndex        =   32
            Tag             =   "Nome da formação escolar"
            ToolTipText     =   "Nome da formação escolar"
            Top             =   720
            Width           =   5895
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   12
            Left            =   -74880
            TabIndex        =   31
            Tag             =   "Código da formação escolar"
            ToolTipText     =   "Código da formação escolar"
            Top             =   720
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
            Height          =   255
            Left            =   -73680
            OleObjectBlob   =   "frmCandidatos.frx":29F3D
            TabIndex        =   133
            Top             =   480
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmCandidatos.frx":29FAD
            TabIndex        =   132
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   28
            Tag             =   "Código do treinamento"
            ToolTipText     =   "Código do treinamento"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   1320
            TabIndex        =   29
            Tag             =   "Nome do treinamento"
            ToolTipText     =   "Nome do treinamento"
            Top             =   720
            Width           =   5295
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   5
            Left            =   7440
            TabIndex        =   30
            Top             =   690
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
            Height          =   255
            Left            =   7440
            OleObjectBlob   =   "frmCandidatos.frx":2A019
            TabIndex        =   131
            Top             =   480
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmCandidatos.frx":2A083
            TabIndex        =   130
            Top             =   480
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmCandidatos.frx":2A115
            TabIndex        =   129
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   -67800
            OleObjectBlob   =   "frmCandidatos.frx":2A181
            TabIndex        =   117
            Top             =   540
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   -69840
            OleObjectBlob   =   "frmCandidatos.frx":2A1EF
            TabIndex        =   116
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -73680
            TabIndex        =   18
            Tag             =   "Nome do cargo da experiência"
            ToolTipText     =   "Nome do cargo da experiência"
            Top             =   780
            Width           =   3255
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   8
            Left            =   -74880
            TabIndex        =   17
            Tag             =   "Código do cargo da experiência"
            ToolTipText     =   "Código do cargo da experiência"
            Top             =   780
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   -73680
            OleObjectBlob   =   "frmCandidatos.frx":2A277
            TabIndex        =   115
            Top             =   540
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmCandidatos.frx":2A2F1
            TabIndex        =   114
            Top             =   540
            Width           =   1095
         End
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   24
            Left            =   -64920
            TabIndex        =   109
            Tag             =   "Experiência exigida pela matriz"
            ToolTipText     =   "Experiência exigida pela matriz"
            Top             =   1140
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
            MICON           =   "frmCandidatos.frx":2A369
            PICN            =   "frmCandidatos.frx":2A385
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   3
            Left            =   -73440
            TabIndex        =   108
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   2
            Left            =   -67680
            TabIndex        =   107
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   6720
            TabIndex        =   106
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   -70320
            TabIndex        =   105
            Top             =   840
            Width           =   375
         End
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   19
            Left            =   -73080
            TabIndex        =   52
            Tag             =   "Excluir histórico"
            ToolTipText     =   "Excluir histórico"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2B05F
            PICN            =   "frmCandidatos.frx":2B07B
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
            Index           =   20
            Left            =   -73680
            TabIndex        =   51
            Tag             =   "Editar histórico"
            ToolTipText     =   "Editar histórico"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2BD55
            PICN            =   "frmCandidatos.frx":2BD71
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
            Index           =   21
            Left            =   -74280
            TabIndex        =   50
            Tag             =   "Novo histórico"
            ToolTipText     =   "Novo histórico"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2CA4B
            PICN            =   "frmCandidatos.frx":2CA67
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
            Index           =   22
            Left            =   -74880
            TabIndex        =   49
            Tag             =   "Incluir histórico"
            ToolTipText     =   "Incluir histórico"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2D741
            PICN            =   "frmCandidatos.frx":2D75D
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
            Left            =   -73080
            TabIndex        =   42
            Tag             =   "Excluir escolaridade"
            ToolTipText     =   "Excluir escolaridade"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2E437
            PICN            =   "frmCandidatos.frx":2E453
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
            Left            =   -73680
            TabIndex        =   41
            Tag             =   "Editar escolaridade"
            ToolTipText     =   "Editar escolaridade"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2F12D
            PICN            =   "frmCandidatos.frx":2F149
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
            Index           =   10
            Left            =   -74280
            TabIndex        =   40
            Tag             =   "Novo escolaridade"
            ToolTipText     =   "Novo escolaridade"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":2FE23
            PICN            =   "frmCandidatos.frx":2FE3F
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
            Index           =   16
            Left            =   -74880
            TabIndex        =   39
            Tag             =   "Incluir escolaridade"
            ToolTipText     =   "Incluir escolaridade"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":30B19
            PICN            =   "frmCandidatos.frx":30B35
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
            Index           =   5
            Left            =   1320
            TabIndex        =   35
            Tag             =   "Excluir treinamento"
            ToolTipText     =   "Excluir treinamento"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":3180F
            PICN            =   "frmCandidatos.frx":3182B
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
            Left            =   720
            TabIndex        =   34
            Tag             =   "Novo treinamento"
            ToolTipText     =   "Novo treinamento"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":32505
            PICN            =   "frmCandidatos.frx":32521
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
            Left            =   120
            TabIndex        =   33
            Tag             =   "Incluir treinamento"
            ToolTipText     =   "Incluir treinamento"
            Top             =   1080
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
            MICON           =   "frmCandidatos.frx":331FB
            PICN            =   "frmCandidatos.frx":33217
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
            Index           =   3
            Left            =   -73080
            TabIndex        =   26
            Tag             =   "Excluir experiência"
            ToolTipText     =   "Excluir experiência"
            Top             =   1140
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
            MICON           =   "frmCandidatos.frx":33EF1
            PICN            =   "frmCandidatos.frx":33F0D
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
            Left            =   -73680
            TabIndex        =   25
            Tag             =   "Editar experiência"
            ToolTipText     =   "Editar experiência"
            Top             =   1140
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
            MICON           =   "frmCandidatos.frx":34BE7
            PICN            =   "frmCandidatos.frx":34C03
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
            Left            =   -74280
            TabIndex        =   24
            Tag             =   "Novo experiência"
            ToolTipText     =   "Novo experiência"
            Top             =   1140
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
            MICON           =   "frmCandidatos.frx":358DD
            PICN            =   "frmCandidatos.frx":358F9
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
            Left            =   -74880
            TabIndex        =   23
            Tag             =   "Incluir experiência"
            ToolTipText     =   "Incluir experiência"
            Top             =   1140
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
            MICON           =   "frmCandidatos.frx":365D3
            PICN            =   "frmCandidatos.frx":365EF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   100
            Top             =   480
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   6376
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "RM Sistemas"
            TabPicture(0)   =   "frmCandidatos.frx":372C9
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SkinLabel29"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SkinLabel30"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SkinLabel31"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SkinLabel32"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "SkinLabel33"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "SkinLabel34"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "SkinLabel35"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "SkinLabel36"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "SkinLabel37"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "SkinLabel38"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "SkinLabel39"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "SkinLabel40"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "SkinLabel41"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtCons(0)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "Combo(1)"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "Combo(2)"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "txtCons(1)"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "Combo(3)"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "txtCons(2)"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "Combo(4)"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).Control(20)=   "txtCons(3)"
            Tab(0).Control(20).Enabled=   0   'False
            Tab(0).Control(21)=   "Combo(5)"
            Tab(0).Control(21).Enabled=   0   'False
            Tab(0).Control(22)=   "txtCons(4)"
            Tab(0).Control(22).Enabled=   0   'False
            Tab(0).Control(23)=   "Combo(6)"
            Tab(0).Control(23).Enabled=   0   'False
            Tab(0).Control(24)=   "txtCons(5)"
            Tab(0).Control(24).Enabled=   0   'False
            Tab(0).Control(25)=   "Combo(7)"
            Tab(0).Control(25).Enabled=   0   'False
            Tab(0).Control(26)=   "txtCons(6)"
            Tab(0).Control(26).Enabled=   0   'False
            Tab(0).Control(27)=   "Combo(8)"
            Tab(0).Control(27).Enabled=   0   'False
            Tab(0).Control(28)=   "txtCons(7)"
            Tab(0).Control(28).Enabled=   0   'False
            Tab(0).Control(29)=   "Combo(9)"
            Tab(0).Control(29).Enabled=   0   'False
            Tab(0).Control(30)=   "txtCons(8)"
            Tab(0).Control(30).Enabled=   0   'False
            Tab(0).Control(31)=   "Combo(10)"
            Tab(0).Control(31).Enabled=   0   'False
            Tab(0).Control(32)=   "txtCons(9)"
            Tab(0).Control(32).Enabled=   0   'False
            Tab(0).Control(33)=   "Combo(11)"
            Tab(0).Control(33).Enabled=   0   'False
            Tab(0).Control(34)=   "txtCons(10)"
            Tab(0).Control(34).Enabled=   0   'False
            Tab(0).Control(35)=   "Combo(12)"
            Tab(0).Control(35).Enabled=   0   'False
            Tab(0).Control(36)=   "txtCons(11)"
            Tab(0).Control(36).Enabled=   0   'False
            Tab(0).Control(37)=   "Combo(13)"
            Tab(0).Control(37).Enabled=   0   'False
            Tab(0).Control(38)=   "txtCons(12)"
            Tab(0).Control(38).Enabled=   0   'False
            Tab(0).ControlCount=   39
            TabCaption(1)   =   "Microsiga"
            TabPicture(1)   =   "frmCandidatos.frx":372E5
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "SAP"
            TabPicture(2)   =   "frmCandidatos.frx":37301
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   12
               Left            =   120
               TabIndex        =   74
               Tag             =   "Membro sindical"
               ToolTipText     =   "Membro sindical"
               Top             =   3000
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   13
               Left            =   960
               TabIndex        =   75
               Tag             =   "Membro sindical"
               ToolTipText     =   "Membro sindical"
               Top             =   3000
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   11
               Left            =   7080
               TabIndex        =   72
               Tag             =   "RAIS situação"
               ToolTipText     =   "RAIS situação"
               Top             =   2400
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   12
               Left            =   7920
               TabIndex        =   73
               Tag             =   "RAIS situação"
               ToolTipText     =   "RAIS situação"
               Top             =   2400
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   10
               Left            =   3600
               TabIndex        =   70
               Tag             =   "Contribuição sindical"
               ToolTipText     =   "Contribuição sindical"
               Top             =   2400
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   11
               Left            =   4440
               TabIndex        =   71
               Tag             =   "Contribuição sindical"
               ToolTipText     =   "Contribuição sindical"
               Top             =   2400
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   9
               Left            =   120
               TabIndex        =   68
               Tag             =   "Seção"
               ToolTipText     =   "Seção"
               Top             =   2400
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   10
               ItemData        =   "frmCandidatos.frx":3731D
               Left            =   960
               List            =   "frmCandidatos.frx":3731F
               TabIndex        =   69
               Tag             =   "Seção"
               ToolTipText     =   "Seção"
               Top             =   2400
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   8
               Left            =   7080
               TabIndex        =   66
               Tag             =   "Função"
               ToolTipText     =   "Função"
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   9
               Left            =   7920
               TabIndex        =   67
               Tag             =   "Função"
               ToolTipText     =   "Função"
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   7
               Left            =   3600
               TabIndex        =   64
               Tag             =   "Horário de trabalho"
               ToolTipText     =   "Horário de trabalho"
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   8
               Left            =   4440
               TabIndex        =   65
               Tag             =   "Horário de trabalho"
               ToolTipText     =   "Horário de trabalho"
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   6
               Left            =   120
               TabIndex        =   62
               Tag             =   "Tipo de funcionário"
               ToolTipText     =   "Tipo de funcionário"
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   7
               Left            =   960
               TabIndex        =   63
               Tag             =   "Tipo de funcionário"
               ToolTipText     =   "Tipo de funcionário"
               Top             =   1800
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   5
               Left            =   7080
               TabIndex        =   60
               Tag             =   "Situação"
               ToolTipText     =   "Situação"
               Top             =   1200
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   6
               Left            =   7920
               TabIndex        =   61
               Tag             =   "Situação"
               ToolTipText     =   "Situação"
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   4
               Left            =   3600
               TabIndex        =   58
               Tag             =   "Forma de recebimento"
               ToolTipText     =   "Forma de recebimento"
               Top             =   1200
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   5
               Left            =   4440
               TabIndex        =   59
               Tag             =   "Forma de recebimento"
               ToolTipText     =   "Forma de recebimento"
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   56
               Tag             =   "Motivo da admissão"
               ToolTipText     =   "Motivo da admissão"
               Top             =   1200
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   4
               Left            =   960
               TabIndex        =   57
               Tag             =   "Motivo da admissão"
               ToolTipText     =   "Motivo da admissão"
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   2
               Left            =   7080
               TabIndex        =   54
               Tag             =   "Tipo de admissão"
               ToolTipText     =   "Tipo de admissão"
               Top             =   600
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   3
               Left            =   7920
               TabIndex        =   55
               Tag             =   "Tipo de admissão"
               ToolTipText     =   "Tipo de admissão"
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   1
               Left            =   3600
               TabIndex        =   48
               Tag             =   "Grau de instrução"
               ToolTipText     =   "Grau de instrução"
               Top             =   600
               Width           =   735
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   2
               Left            =   4440
               TabIndex        =   53
               Tag             =   "Grau de instrução"
               ToolTipText     =   "Grau de instrução"
               Top             =   600
               Width           =   2415
            End
            Begin VB.ComboBox Combo 
               Height          =   315
               Index           =   1
               Left            =   960
               TabIndex        =   47
               Tag             =   "Sexo"
               ToolTipText     =   "Sexo"
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox txtCons 
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   46
               Tag             =   "Sexo"
               ToolTipText     =   "Sexo"
               Top             =   600
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37321
               TabIndex        =   150
               Top             =   2760
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
               Height          =   255
               Left            =   7080
               OleObjectBlob   =   "frmCandidatos.frx":3739F
               TabIndex        =   149
               Top             =   2160
               Width           =   1815
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
               Height          =   255
               Left            =   3600
               OleObjectBlob   =   "frmCandidatos.frx":37419
               TabIndex        =   148
               Top             =   2160
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":374A3
               TabIndex        =   147
               Top             =   2160
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
               Height          =   255
               Left            =   7080
               OleObjectBlob   =   "frmCandidatos.frx":3750D
               TabIndex        =   146
               Top             =   1560
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
               Height          =   255
               Left            =   3600
               OleObjectBlob   =   "frmCandidatos.frx":37579
               TabIndex        =   145
               Top             =   1560
               Width           =   2295
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":375FF
               TabIndex        =   144
               Top             =   1560
               Width           =   2055
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
               Height          =   255
               Left            =   7080
               OleObjectBlob   =   "frmCandidatos.frx":37685
               TabIndex        =   143
               Top             =   960
               Width           =   1695
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
               Height          =   255
               Left            =   3600
               OleObjectBlob   =   "frmCandidatos.frx":376F5
               TabIndex        =   142
               Top             =   960
               Width           =   2655
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37779
               TabIndex        =   141
               Top             =   960
               Width           =   2655
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
               Height          =   255
               Left            =   7080
               OleObjectBlob   =   "frmCandidatos.frx":377FD
               TabIndex        =   140
               Top             =   360
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
               Height          =   255
               Left            =   3600
               OleObjectBlob   =   "frmCandidatos.frx":3787D
               TabIndex        =   139
               Top             =   360
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":378FF
               TabIndex        =   138
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Documentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -71160
            TabIndex        =   79
            Top             =   420
            Width           =   6975
            Begin VB.Frame Frame7 
               Caption         =   "CNH - Carteira Nacional de Habilitação"
               Height          =   975
               Left            =   3480
               TabIndex        =   81
               Top             =   240
               Width           =   3375
               Begin VB.TextBox txtCadMatriz 
                  Height          =   285
                  Index           =   18
                  Left            =   2160
                  TabIndex        =   13
                  Tag             =   "Tipo da carteira CNH"
                  ToolTipText     =   "Tipo da carteira CNH"
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.TextBox txtCadMatriz 
                  Height          =   285
                  Index           =   17
                  Left            =   120
                  TabIndex        =   12
                  Tag             =   "Número da carteira CNH"
                  ToolTipText     =   "Número da carteira CNH"
                  Top             =   480
                  Width           =   1935
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
                  Height          =   255
                  Left            =   2160
                  OleObjectBlob   =   "frmCandidatos.frx":37967
                  TabIndex        =   128
                  Top             =   240
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmCandidatos.frx":379CF
                  TabIndex        =   127
                  Top             =   240
                  Width           =   855
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Carteira de Trabalho e Previdência Social"
               Height          =   975
               Index           =   1
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   3255
               Begin VB.TextBox txtCadMatriz 
                  Height          =   285
                  Index           =   16
                  Left            =   2040
                  TabIndex        =   11
                  Tag             =   "Série da carteira de trabalho"
                  ToolTipText     =   "Série da carteira de trabalho"
                  Top             =   480
                  Width           =   1095
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
                  Height          =   255
                  Left            =   2040
                  OleObjectBlob   =   "frmCandidatos.frx":37A3B
                  TabIndex        =   126
                  Top             =   240
                  Width           =   855
               End
               Begin VB.TextBox txtCadMatriz 
                  Height          =   285
                  Index           =   15
                  Left            =   120
                  TabIndex        =   10
                  Tag             =   "Número da carteira de trabalho"
                  ToolTipText     =   "Número da carteira de trabalho"
                  Top             =   480
                  Width           =   1815
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmCandidatos.frx":37AA5
                  TabIndex        =   125
                  Top             =   240
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Observação "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   -71160
            TabIndex        =   78
            Top             =   1860
            Width           =   6975
            Begin VB.TextBox txtCadMatriz 
               Height          =   1095
               Index           =   19
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Tag             =   "Observação"
               ToolTipText     =   "Observação"
               Top             =   240
               Width           =   6735
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Contatos (Residencial/Celular)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -71160
            TabIndex        =   97
            Top             =   3420
            Width           =   6975
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   21
               Left            =   2880
               TabIndex        =   16
               Tag             =   "Telefone celular"
               ToolTipText     =   "Telefone celular"
               Top             =   240
               Width           =   2655
            End
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Tag             =   "Telefone residencial"
               ToolTipText     =   "Telefone residencial"
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Marcar/Desmarcar"
            Height          =   255
            Left            =   -74880
            TabIndex        =   83
            Top             =   420
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   1
            Left            =   -67800
            TabIndex        =   21
            Tag             =   "Nome da empresa"
            ToolTipText     =   "Nome da empresa"
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   2
            ItemData        =   "frmCandidatos.frx":37B11
            Left            =   -69720
            List            =   "frmCandidatos.frx":37B39
            TabIndex        =   19
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "001"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   780
            Width           =   735
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   3
            ItemData        =   "frmCandidatos.frx":37B79
            Left            =   -68880
            List            =   "frmCandidatos.frx":37B83
            TabIndex        =   20
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "Meses"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   780
            Width           =   855
         End
         Begin VB.Frame Frame3 
            Caption         =   "Dados "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3615
            Left            =   -74880
            TabIndex        =   82
            Top             =   420
            Width           =   3615
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   23
               Left            =   120
               TabIndex        =   9
               Tag             =   "Email"
               ToolTipText     =   "Email"
               Top             =   3240
               Width           =   3375
            End
            Begin VB.ComboBox cboCadMatriz 
               Height          =   315
               Index           =   4
               ItemData        =   "frmCandidatos.frx":37B94
               Left            =   2880
               List            =   "frmCandidatos.frx":37BE9
               TabIndex        =   8
               Tag             =   "Estado de naturalidade"
               ToolTipText     =   "Estado de naturalidade"
               Top             =   2640
               Width           =   615
            End
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   14
               Left            =   120
               TabIndex        =   7
               Tag             =   "Naturalidade"
               ToolTipText     =   "Naturalidade"
               Top             =   2640
               Width           =   2655
            End
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   6
               Tag             =   "Nacionalidade"
               ToolTipText     =   "Nacionalidade"
               Top             =   1920
               Width           =   3375
            End
            Begin VB.ComboBox cboCadMatriz 
               Height          =   315
               Index           =   1
               ItemData        =   "frmCandidatos.frx":37C59
               Left            =   120
               List            =   "frmCandidatos.frx":37C6F
               TabIndex        =   5
               Tag             =   "Estado civil"
               ToolTipText     =   "Estado civil"
               Top             =   1200
               Width           =   3375
            End
            Begin VB.ComboBox cboCadMatriz 
               Height          =   315
               Index           =   0
               ItemData        =   "frmCandidatos.frx":37CBE
               Left            =   1560
               List            =   "frmCandidatos.frx":37CC8
               TabIndex        =   4
               Tag             =   "Sexo"
               ToolTipText     =   "Sexo"
               Top             =   480
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   120
               TabIndex        =   3
               Tag             =   "Data de nascimento do colaborador"
               ToolTipText     =   "Data de nascimento do colaborador"
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   16187393
               CurrentDate     =   40499
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37CE1
               TabIndex        =   124
               Top             =   3000
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
               Height          =   255
               Left            =   2880
               OleObjectBlob   =   "frmCandidatos.frx":37D4B
               TabIndex        =   123
               Top             =   2400
               Width           =   375
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37DAF
               TabIndex        =   122
               Top             =   2400
               Width           =   1095
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37E27
               TabIndex        =   121
               Top             =   1680
               Width           =   1455
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37EA1
               TabIndex        =   120
               Top             =   960
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   1560
               OleObjectBlob   =   "frmCandidatos.frx":37F19
               TabIndex        =   119
               Top             =   240
               Width           =   615
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmCandidatos.frx":37F81
               TabIndex        =   118
               Top             =   240
               Width           =   1095
            End
         End
         Begin MSComctlLib.ListView ListView5 
            Height          =   2355
            Left            =   -74880
            TabIndex        =   45
            Top             =   1740
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4154
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
         Begin MSComctlLib.ListView ListView4 
            Height          =   2355
            Left            =   -74880
            TabIndex        =   36
            Top             =   1740
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4154
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
         Begin MSComctlLib.ListView ListView3 
            Height          =   2295
            Left            =   120
            TabIndex        =   84
            Top             =   1740
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   3615
            Left            =   -74880
            TabIndex        =   27
            Top             =   420
            Width           =   10695
            _ExtentX        =   18865
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
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   22
            Top             =   1800
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4048
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
         Begin VB.Label Label20 
            Caption         =   "Nome do cargo:"
            Height          =   255
            Left            =   -73680
            TabIndex        =   87
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Tempo de experiência:"
            Height          =   255
            Left            =   -67560
            TabIndex        =   86
            Top             =   540
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "Código cargo:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   85
            Top             =   540
            Width           =   1215
         End
      End
   End
   Begin IpToolTips.cIpToolTips cIpToolTips1 
      Left            =   3600
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status "
      Enabled         =   0   'False
      Height          =   735
      Left            =   10200
      TabIndex        =   90
      Top             =   8640
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
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
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   120
      TabIndex        =   95
      Top             =   3960
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "frmCandidatos"
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

Private rsColaboradores As New ADODB.Recordset
Private SqlColaboradores As String
Private rsCargos As New ADODB.Recordset 'OK
Private SqlCargos As String 'OK
Private rsCursos As New ADODB.Recordset
Private SqlCursos As String
Private rsEscolaridade As New ADODB.Recordset
Private sqlEscolaridade As String
Private rsHistorico As New ADODB.Recordset
Private SqlHistorico As String
Private CaMinho As String
Private Status As String
Private rsLocal As New ADODB.Recordset
Private Caminho1 As String

Private Sub chameleonButton1_Click()
    WebBrowser1.Visible = True
    Frame9.Visible = True
    Timer1.Enabled = True
End Sub

Private Sub cmdCad_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridCargo 8
        CarregaCargo 8
    Case 1
        ChamaGridCurso
        CarregaCurso
        CompoeComboNivel cboCadMatriz(5), txtCadMatriz(10)
    Case 2
        ChamaGridEscolaridade
        CarregaEscolaridade
    Case 3
        ChamaGridHistorico
        CarregaHistorico
    End Select
End Sub

Private Sub Combo_Click(Index As Integer)
    Select Case Index
    Case 1
        AchaComboTotvs Combo(Index), "PCODSEXO", "CODINTERNO", Index, "descricao"
    Case 2
        AchaComboTotvs Combo(Index), "PCODINSTRUCAO", "CODINTERNO", Index, "descricao"
    Case 3
        AchaComboTotvs Combo(Index), "PTPADMISSAO", "CODINTERNO", Index, "descricao"
    Case 4
        AchaComboTotvs Combo(Index), "PMOTADMISSAO", "CODINTERNO", Index, "descricao"
    Case 5
        AchaComboTotvs Combo(Index), "PCODRECEB", "CODINTERNO", Index, "descricao"
    Case 6
        AchaComboTotvs Combo(Index), "PCODSITUACAO", "CODINTERNO", Index, "descricao"
    Case 7
        AchaComboTotvs Combo(Index), "PTPFUNC", "CODINTERNO", Index, "descricao"
    Case 8
        AchaComboTotvs Combo(Index), "AHORARIO", "CODIGO", Index, "descricao"
    Case 9
        AchaComboTotvs Combo(Index), "PFUNCAO", "CODIGO", Index, "nome"
    Case 10
        AchaComboTotvs Combo(Index), "PSECAO", "CODIGO", Index, "descricao"
    Case 11
        AchaComboTotvs Combo(Index), "PCODCTSIND", "CODINTERNO", Index, "descricao"
    Case 12
        AchaComboTotvs Combo(Index), "PCODSITRAIS", "CODINTERNO", Index, "descricao"
    Case 13
        AchaComboTotvs Combo(Index), "PSINDIC", "CODIGO", Index, "nome"
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirExperiencia
        LimpaControlesExp
    Case 1
        LimpaControlesExp
    Case 2
        AlteraExperiencia
    Case 3
        mobjMsg.Abrir "Deseja EXCLUIR essa experiência do candidato?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView1
            LimpaControlesExp
        End If
    Case 5
        mobjMsg.Abrir "Deseja EXCLUIR esse curso/treinamento do candidato?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView3
            LimpaControlesTreinamento
        End If
    Case 6
        mobjMsg.Abrir "Deseja EXCLUIR essa formação escolar do candidato?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView4
            LimpaControlesEscolaridade
        End If
    Case 7
        LimpaControlesTreinamento
    Case 8
        IncluirTreinamento
        LimpaControlesTreinamento
    Case 9
        AlteraEscolaridade
    Case 10
        LimpaControlesEscolaridade
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
        Label53 = ""
    Case 14
        mobjMsg.Abrir "Deseja salvar os dados do candidato?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            gravaLog " CPF: " & mskCadMatriz & " Nome: " & txtCadMatriz(3), " Matriz/cargo: " & txtCadMatriz(4), "Média: " & Label41
            'AtivaLD
            'Unload Me
        End If
    Case 15
        mobjMsg.Abrir "Deseja sair da tela de cadastro do candidato?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Unload Me
            Set frmCandidatos = Nothing
        End If
    Case 16
        IncluirEscolaridade
        LimpaControlesEscolaridade
    Case 19
        mobjMsg.Abrir "Deseja EXCLUIR esse cargo do histórico funcional?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView5
            LimpaControlesHistorico
        End If
    Case 20
        AlteraHistorico
    Case 21
        LimpaControlesHistorico
    Case 22
        IncluirHistorico
        LimpaControlesHistorico
        CompoeLVHab
        CompoePontosLVHab
    Case 24
        Campo4 = 1
        frmAvisos.Show 1
    Case 25
        Campo4 = 2
        frmAvisos.Show 1
    Case 26
        Campo4 = 3
        frmAvisos.Show 1
    End Select
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    SSTab1.Tab = 0
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker4 = Date
    WebBrowser1.Navigate "about:<html><body scroll='no'><img src=" & "C:\Arquiv~1\SGCH\aguarde-01.gif" & " ></img></body></html>"
    listview_cabecalho
    If Status = "novo" Then
        'CompoeLVHab 'Compoe habilidade do cargo
    ElseIf Status = "editar" Then
        ResultPesq "editar"
        mskCadMatriz.PromptInclude = False
        CompoeLVExp 'Compoe Experiencia do cargo
        CompoeLVTrei 'Compoe Treinamento do cargo
        CompoeLVFor 'Compoe Formação escolar do cargo
        CompoeLVHist 'Compoe Historico funcional do cargo
        CompoeLVHab 'Compoe habilidade do cargo
        CompoePontosLVHab 'Compoe pontuação de habilidade do colaborador para a matriz
        mskCadMatriz.PromptInclude = True
        MudaCorLV5
    End If
    configControles
    If vIntegra = "S" Then ConexaoTotvs
    If vIntegra = "S" Then comporCombosTotvs
    If vIntegra = "S" Then
        If lblID <> "" Then comporControlesTotvs
    End If
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    'EXPERIÊNCIAS
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome do cargo", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Tempo de experiência", ListView1.Width / 5
    ListView1.ColumnHeaders.Add , , "Nome Empresa", ListView1.Width / 4
    
    'HABILIDADES
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView2.ColumnHeaders.Add , , "Habilidade", ListView1.Width / 1.5
    ListView2.ColumnHeaders.Add , , "Peso", ListView1.Width / 10
    ListView2.ColumnHeaders.Add , , "Avaliado", ListView1.Width / 10
    
    'CURSOS/TREINAMENTOS
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView3.ColumnHeaders.Add , , "Nome do curso/treinamento", ListView1.Width / 1.5
    ListView3.ColumnHeaders.Add , , "Origem", ListView1.Width / 10000
    ListView3.ColumnHeaders.Add , , "Nível", ListView3.Width / 4.5
    
    'ESCOLARIDADE
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView4.ColumnHeaders.Add , , "Formação escolar", ListView1.Width / 1.5
    
    'CARGOS PRETENDIDOS
    ListView5.ColumnHeaders.Clear
    ListView5.ColumnHeaders.Add , , "Matriz", ListView1.Width / 13
    ListView5.ColumnHeaders.Add , , "Código", ListView1.Width / 13
    ListView5.ColumnHeaders.Add , , "Nome do cargo", ListView1.Width / 3
    ListView5.ColumnHeaders.Add , , "Nível", ListView1.Width / 16
    ListView5.ColumnHeaders.Add , , "Ativo", ListView1.Width / 10000
    ListView5.ColumnHeaders.Add , , "Sequencia", ListView1.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

' (INICIO) >>>>>>>> LIMPA CRONTROLES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<
Private Sub LimpaControlesExp()
    Dim X As Integer
    txtCadMatriz(8).Enabled = True
    cmdCadastro(18).Enabled = True
    
    cboCadMatriz(2).Text = ""
    cboCadMatriz(3).Text = ""
    For X = 8 To 9
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(1) = ""
    txtCadMatriz(8).SetFocus
End Sub

' (INICIO) >>>>>>>> LIMPA CRONTROLES DA GUIA DE TREINAMENTOS <<<<<<<<<<
Private Sub LimpaControlesTreinamento()
    Dim X As Integer
    txtCadMatriz(10).Enabled = True
    cmdCadastro(4).Enabled = True
    
    For X = 10 To 11
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(10).SetFocus
End Sub

' (INICIO) >>>>>>>> LIMPA CRONTROLES DA GUIA DE ESCOLARIDADE <<<<<<<<<<
Private Sub LimpaControlesEscolaridade()
    Dim X As Integer
    txtCadMatriz(12).Enabled = True
    cmdCadastro(17).Enabled = True
    For X = 12 To 13
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(12).SetFocus
End Sub

' (INICIO) >>>>>>>> LIMPA CRONTROLES DA GUIA DE HISTORICO FUNCIONAL <<<<<<<<<<
Private Sub LimpaControlesHistorico()
    Dim X As Integer
    txtCadMatriz(0).Enabled = True
    cmdCad(3).Enabled = True
    txtCadMatriz(0) = ""
    txtCadMatriz(5) = ""
    txtCadMatriz(6) = ""
    txtCadMatriz(20) = ""
    DTPicker2 = Date
    txtCadMatriz(0).SetFocus
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    Dim ProCura As String
    ProCura = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
    
    For X = 0 To Len(ProCura)
        If Mid$(ProCura, X + 1, 1) = "E" Then chkAvaliador(0).Value = 1
        If Mid$(ProCura, X + 1, 1) = "H" Then chkAvaliador(1).Value = 1
        If Mid$(ProCura, X + 1, 1) = "T" Then chkAvaliador(2).Value = 1
        If Mid$(ProCura, X + 1, 1) = "F" Then chkAvaliador(3).Value = 1
    Next
    
    WebBrowser1.Visible = True
    Frame9.Visible = True
    Timer1.Enabled = True
    
    mskCadMatriz.PromptInclude = False
    mskCadMatriz.Text = rsColaboradores.Fields(0) 'cpf
    mskCadMatriz.PromptInclude = True
    'txtCadMatriz(2).Text = rsColaboradores.Fields(1) 'codigo do colaborador
    DTPicker4 = rsColaboradores.Fields(2) 'data do cadastro
    txtCadMatriz(3).Text = rsColaboradores.Fields(3) 'nome do colaborador
    DTPicker1 = rsColaboradores.Fields(4) 'Data nascimento
    cboCadMatriz(0).Text = rsColaboradores.Fields(5) 'sexo
    cboCadMatriz(1).Text = rsColaboradores.Fields(6) 'estado civil
    txtCadMatriz(7).Text = rsColaboradores.Fields(7) 'nacionalidade
    txtCadMatriz(14).Text = rsColaboradores.Fields(8) 'naturalidade
    cboCadMatriz(4).Text = rsColaboradores.Fields(9) 'ufnaturalidade
    txtCadMatriz(15).Text = rsColaboradores.Fields(10) 'ctps numero
    txtCadMatriz(16).Text = rsColaboradores.Fields(11) 'ctps serie
    txtCadMatriz(17).Text = rsColaboradores.Fields(12) 'cnh numero
    txtCadMatriz(18).Text = rsColaboradores.Fields(13) 'cnh tipo
    If Not IsNull(rsColaboradores.Fields(20)) Then txtCadMatriz(19).Text = rsColaboradores.Fields(20) 'observacao
    If rsColaboradores.Fields(17) = "S" Then 'ativo
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If Not IsNull(rsColaboradores.Fields(22)) Then txtCadMatriz(23).Text = rsColaboradores.Fields(22) 'email
    If Not IsNull(rsColaboradores.Fields(24)) Then txtCadMatriz(2).Text = rsColaboradores.Fields(24) 'Telefone
    If Not IsNull(rsColaboradores.Fields(25)) Then txtCadMatriz(21).Text = rsColaboradores.Fields(25) 'Celular
    
    If rsColaboradores.Fields(18) < MediaGlobal Then
        'Label41.ForeColor = &HC0&
    Else
        'Label41.ForeColor = &H8000&
    End If
    
    Label41.Caption = Format(rsColaboradores.Fields(18), "#,##0.00;(#,##0.00)") & " %" 'media geral
    
    If rsColaboradores.Fields(19) <> "Null" Then
        On Error GoTo TrataErro1
        Label53.Caption = rsColaboradores.Fields(19)
        aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    End If
    lblID = rsColaboradores.Fields(29)
    Exit Sub
TrataErro1:
    Label59.Visible = True
    Resume Next
End Sub

' (INICIO) >>>>>>>> COMPOE LISTVIEW1 DA GUIA DE EXPERIÊNCIA <<<<<<<<<<
Private Sub CompoeLVExp()
    Dim rsExp As New ADODB.Recordset
    Dim sqlExp As String
    sqlExp = "Select tbColaboradoresExp.*, tbcargos.nomecargo from tbColaboradoresExp,tbcargos where tbColaboradoresExp.codcoligada = '" & vCodcoligada & "' and tbColaboradoresExp.codcargo=tbcargos.codcargo and tbColaboradoresExp.CPF = '" & mskCadMatriz.Text & "'"
    rsExp.Open sqlExp, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsExp.EOF
    
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsExp.Fields(3), "000000"))
        ItemLst.SubItems(1) = "" & rsExp.Fields(6)
        ItemLst.SubItems(2) = "" & rsExp.Fields(4)
        ItemLst.SubItems(3) = "" & rsExp.Fields(2)
        rsExp.MoveNext
        X = X + 1
    Wend
    rsExp.Close
    Set rsExp = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE LISTVIEW2 DA GUIA DE HABILIDADES <<<<<<<<<<
Private Sub CompoeLVHab()
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidades As String
    sqlHabilidades = "Select tbMatrizHab.*, tbhabilidades.nomehabilidade, tbhabilidades.peso from tbMatrizHab, tbhabilidades where tbMatrizHab.codcoligada = '" & vCodcoligada & "' and tbMatrizHab.codhabilidade = tbhabilidades.codhabilidade and tbMatrizHab.codmatriz = '" & Val(Mid$(txtCadMatriz(4), 1, 6)) & "'order by tbMatrizHab.codhabilidade"
    rsHabilidade.Open sqlHabilidades, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    
    X = 0
    ListView2.ListItems.Clear
    While Not rsHabilidade.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsHabilidade.Fields(1), "00")) 'codigo da habilidade
        ItemLst.SubItems(1) = "" & rsHabilidade.Fields(3) 'nome da habilidade
        ItemLst.SubItems(2) = "" & rsHabilidade.Fields(4) 'peso da habilidade
        ItemLst.SubItems(3) = "" & 0 'avaliação
        ItemLst.ListSubItems(3).Bold = True
        
        rsHabilidade.MoveNext
        X = X + 1
    Wend
    rsHabilidade.Close
    Set rsHabilidade = Nothing
    Me.ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE PONTUAÇÃO DO LISTVIEW2 DA GUIA DE HABILIDADES <<<<<<<<<<
Private Sub CompoePontosLVHab()
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidades As String
    mskCadMatriz.PromptInclude = False
    sqlHabilidades = "Select tbColaboradoresHab.* from tbColaboradoresHab where tbColaboradoresHab.codcoligada = '" & vCodcoligada & "' and tbColaboradoresHab.cpf = '" & mskCadMatriz & "' and tbColaboradoresHab.codmatriz = '" & Val(Mid$(txtCadMatriz(4), 1, 6)) & "'order by tbColaboradoresHab.codhabilidade"
    rsHabilidade.Open sqlHabilidades, cnBanco, adOpenKeyset, adLockOptimistic
    
    'DAKI P BAIXO VOU BUSCAR DA MATRIZ DE CAPACITAÇÃO
    
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    While Not rsHabilidade.EOF
        For X = 1 To Y
            ListView2.ListItems(X).Selected = True
            If Val(ListView2.ListItems.Item(X)) = rsHabilidade.Fields(2) Then
                ListView2.SelectedItem.ListSubItems.Item(3) = rsHabilidade.Fields(3)
            End If
        Next
        rsHabilidade.MoveNext
    Wend
    mskCadMatriz.PromptInclude = False
    rsHabilidade.Close
    Set rsHabilidade = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE LISTVIEW3 DA GUIA DE TREINAMENTOS <<<<<<<<<<
Private Sub CompoeLVTrei()
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    'sqlTrei = "Select tbcolaboradoresCur.cpf,tbcolaboradoresCur.tipo,tbcolaboradoresCur.codtreinamento,tbcolaboradoresCur.origem,tbTreinamentos.nometreinamento from tbcolaboradoresCur,tbTreinamentos where tbcolaboradoresCur.codtreinamento=tbTreinamentos.codtreinamento and tbcolaboradoresCur.cpf = '" & mskCadMatriz.Text & "'"
    sqlTrei = "select a.cpf, a.tipo, a.codtreinamento, a.origem, b.nometreinamento, c.codnivel, c.nomenivel from tbcolaboradoresCur as a left join tbTreinamentos as b on a.codtreinamento = b.codtreinamento left join tbTreinamentosNiv as c on b.codtreinamento = c.codtreinamento and a.codnivel = c.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & mskCadMatriz.Text & "'"
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsTrei.EOF
        Set ItemLst = ListView3.ListItems.Add(, , Format(rsTrei.Fields(2), "000000"))
        ItemLst.SubItems(1) = "" & rsTrei.Fields(4)
        ItemLst.SubItems(2) = "" & rsTrei.Fields(3)
        If Not IsNull(rsTrei.Fields(5)) Then ItemLst.SubItems(3) = Format(rsTrei.Fields(5), "00") & " - " & rsTrei.Fields(6) Else ItemLst.SubItems(3) = "-"
        rsTrei.MoveNext
        X = X + 1
    Wend
    rsTrei.Close
    Set rsTrei = Nothing
    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 2
    Me.ListView3.SortOrder = lvwAscending
    MudaCorLV3
End Sub

' (INICIO) >>>>>>>> COMPOE LISTVIEW4 DA GUIA DE FORMAÇÃO ESCOLAR <<<<<<<<<<
Private Sub CompoeLVFor()
    Dim rsEsc As New ADODB.Recordset
    Dim sqlEsc As String
    sqlEsc = "Select tbcolaboradoresEsc.*, tbEscolaridade.nomeescolaridade from tbcolaboradoresEsc,tbEscolaridade where tbcolaboradoresEsc.codcoligada ='" & vCodcoligada & "' and tbcolaboradoresEsc.codescolaridade=tbEscolaridade.codescolaridade and tbcolaboradoresEsc.CPF = '" & mskCadMatriz.Text & "'"
    rsEsc.Open sqlEsc, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    
    X = 0
    While Not rsEsc.EOF
        Set ItemLst = ListView4.ListItems.Add(, , Format(rsEsc.Fields(2), "000000"))
        ItemLst.SubItems(1) = "" & rsEsc.Fields(4)
        rsEsc.MoveNext
        X = X + 1
    Wend
    rsEsc.Close
    Set rsEsc = Nothing
    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 0
    Me.ListView4.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> COMPOE LISTVIEW5 DA GUIA DE HISTORICO FUNCIONAL <<<<<<<<<<
Private Sub CompoeLVHist()
    Dim rsEsc As New ADODB.Recordset
    Dim sqlEsc As String
'    sqlEsc = "Select  a.*, c.codcargo,c.nomecargo,b.nivel from tbcolaboradoresHist as a inner join tbMatriz as b on b.codmatriz=a.codmatriz inner join tbcargos as c on c.codcargo = b.codcargo where a.CPF = '" & mskCadMatriz.Text & "'"
    sqlEsc = "Select  a.cpf,a.codmatriz,a.data,a.motivo,a.observacao,a.ativo,a.sequencia,a.tipo,a.codrequisicao,a.datasai, c.codcargo,c.nomecargo,b.nivel from tbcolaboradoresHist as a inner join tbMatriz as b on a.codcoligada = '" & vCodcoligada & "' and b.codmatriz=a.codmatriz inner join tbcargos as c on c.codcargo = b.codcargo where a.CPF = '" & mskCadMatriz.Text & "'"
    rsEsc.Open sqlEsc, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    
    X = 0
    While Not rsEsc.EOF
        Set ItemLst = ListView5.ListItems.Add(, , Format(rsEsc.Fields(1), "000000")) ' codigo da matriz
        ItemLst.SubItems(1) = "" & Format(rsEsc.Fields(10), "000000") 'codigo do cargo
        ItemLst.SubItems(2) = "" & rsEsc.Fields(11) 'nome do cargo
        ItemLst.SubItems(3) = "" & rsEsc.Fields(12) 'nivel da matriz do cargo
        ItemLst.SubItems(4) = "" & rsEsc.Fields(5) 'ativo
        If rsEsc.Fields(5) = "S" Then
            txtCadMatriz(4) = Format(rsEsc.Fields(1), "000000") & "-" & rsEsc.Fields(10) & " (" & rsEsc.Fields(11) & ")"
            'txtCadMatriz(4) = Format(rsEsc.Fields(1), "000000") & "-" & rsEsc.Fields(10)
        End If
        ItemLst.SubItems(5) = "" & rsEsc.Fields(6) 'sequencia
        rsEsc.MoveNext
        X = X + 1
    Wend
    rsEsc.Close
    Set rsEsc = Nothing
    Me.ListView5.Sorted = True
    Me.ListView5.SortKey = 5
    Me.ListView5.SortOrder = lvwDescending
End Sub

' (INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<
Private Sub IncluirExperiencia()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaExperiencia = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) & ListView1.SelectedItem.ListSubItems.Item(3) = Me.txtCadMatriz(8) & Me.txtCadMatriz(1) Then
                Me.txtCadMatriz(8) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(9)
                ListView1.SelectedItem.ListSubItems.Item(2) = Format(cboCadMatriz(2), "000") & " " & cboCadMatriz(3)
                ListView1.SelectedItem.ListSubItems.Item(3) = txtCadMatriz(1)
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadMatriz(8))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadMatriz(8))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadMatriz(9)
    ItemLst.SubItems(2) = Format(cboCadMatriz(2), "000") & " " & cboCadMatriz(3)
    ItemLst.SubItems(3) = txtCadMatriz(1)
    'txtCadMatriz(8).SetFocus
End Sub

Private Function ValidaExperiencia()
    ValidaExperiencia = False
    If txtCadMatriz(8).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(8).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(8).SetFocus
        Exit Function
    End If
    ValidaExperiencia = True
End Function

Private Sub AlteraExperiencia()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(8).Text = ListView1.ListItems.Item(X)
    Me.txtCadMatriz(9).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.cboCadMatriz(2).Text = Format(Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 3), "000")
    Me.cboCadMatriz(3).Text = Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 5, 10)
    Me.txtCadMatriz(1).Text = ListView1.SelectedItem.ListSubItems.Item(3)
    txtCadMatriz(8).Enabled = False
    txtCadMatriz(9).Enabled = False
    cmdCadastro(18).Enabled = False
End Sub
' (FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE CURSOS/TREINAMENTOS <<<<<<<<<<
Private Sub IncluirTreinamento()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaTreinamento = False Then Exit Sub
    Y = ListView3.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView3.ListItems.Item(X) = Me.txtCadMatriz(10) Then
                Me.txtCadMatriz(10) = ListView3.ListItems.Item(X)
                ListView3.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(11)
                ListView3.SelectedItem.ListSubItems.Item(2) = "C"
                ListView3.SelectedItem.ListSubItems.Item(3) = cboCadMatriz(5)
                Y = ListView3.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView3.ListItems.Add(, , txtCadMatriz(10))
        Y = ListView3.ListItems.Count
    Else
        Set ItemLst = ListView3.ListItems.Add(, , txtCadMatriz(10))
        Y = ListView3.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadMatriz(11)
    ItemLst.SubItems(2) = "C"
    ItemLst.SubItems(3) = cboCadMatriz(5)
    txtCadMatriz(10).SetFocus
End Sub

Private Function ValidaTreinamento()
    ValidaTreinamento = False
    If txtCadMatriz(11).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(11).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(10).SetFocus
        Exit Function
    End If
    ValidaTreinamento = True
End Function
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE CURSOS/TREINAMENTOS <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE FORMAÇÃO ESCOLAR <<<<<<<<<<
Private Sub IncluirEscolaridade()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaEscolaridade = False Then Exit Sub
    Y = ListView4.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView4.ListItems.Item(X) = Me.txtCadMatriz(12) Then
                Me.txtCadMatriz(12) = ListView4.ListItems.Item(X)
                ListView4.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(13)
                Y = ListView4.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView4.ListItems.Add(, , txtCadMatriz(12))
        Y = ListView4.ListItems.Count
    Else
        Set ItemLst = ListView4.ListItems.Add(, , txtCadMatriz(12))
        Y = ListView4.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadMatriz(13)
    txtCadMatriz(10).SetFocus
End Sub

Private Function ValidaEscolaridade()
    ValidaEscolaridade = False
    If txtCadMatriz(12).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(12).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(12).SetFocus
        Exit Function
    End If
    ValidaEscolaridade = True
End Function


Private Sub AlteraEscolaridade()
    If ListView4.ListItems.Count = 0 Then Exit Sub
    Dim Y As Integer, X As Integer
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        If ListView4.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(12).Text = ListView4.ListItems.Item(X)
    Me.txtCadMatriz(13).Text = ListView4.SelectedItem.ListSubItems.Item(1)
    'Me.txtCadMatriz(14).Text = ListView4.SelectedItem.ListSubItems.Item(2)
    txtCadMatriz(12).Enabled = False
    txtCadMatriz(13).Enabled = False
    cmdCadastro(17).Enabled = False
End Sub
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE FORMAÇÃO ESCOLAR <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE HISTÓRICO FUNCIONAL <<<<<<<<<<
Private Sub IncluirHistorico()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaHistorico = False Then Exit Sub
    
    '>>>>>> REMOVE A MARCAÇÃO DE ATIVO DA COLUNA 4
    Y = ListView5.ListItems.Count
    For X = 1 To Y
        ListView5.ListItems.Item(X).Selected = True
        ListView5.SelectedItem.ListSubItems.Item(4) = ""
    Next
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    
    Y = ListView5.ListItems.Count
    txtCadMatriz(4) = txtCadMatriz(0) & "-" & txtCadMatriz(6) & " (" & txtCadMatriz(20) & ")"
    If Y > 0 Then
        For X = 1 To Y
            If ListView5.ListItems.Item(X) = Me.txtCadMatriz(0) Then
                ListView5.ListItems.Item(X).Selected = True
                Me.txtCadMatriz(0) = ListView5.ListItems.Item(X)
                ListView5.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(5)
                ListView5.SelectedItem.ListSubItems.Item(2) = txtCadMatriz(6)
                ListView5.SelectedItem.ListSubItems.Item(3) = txtCadMatriz(20)
                ListView5.SelectedItem.ListSubItems.Item(4) = "S"
                Y = ListView5.ListItems.Count
                MudaCorLV5
                Exit Sub
            End If
        Next
        Set ItemLst = ListView5.ListItems.Add(, , txtCadMatriz(0))
        Y = ListView5.ListItems.Count
    Else
        Set ItemLst = ListView5.ListItems.Add(, , txtCadMatriz(0))
        Y = ListView5.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadMatriz(5)
    ItemLst.SubItems(2) = txtCadMatriz(6)
    ItemLst.SubItems(3) = txtCadMatriz(20)
    ItemLst.SubItems(4) = "S"
    ItemLst.SubItems(5) = ListView5.ListItems.Count
    
    Me.ListView5.Sorted = True
    Me.ListView5.SortKey = 5
    Me.ListView5.SortOrder = lvwDescending
    
    MudaCorLV5
    
    txtCadMatriz(0).SetFocus
End Sub

Private Function ValidaHistorico()
    ValidaHistorico = False
    If txtCadMatriz(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(0).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(0).SetFocus
        Exit Function
    End If
    ValidaHistorico = True
End Function

Private Sub AlteraHistorico()
    If ListView5.ListItems.Count = 0 Then Exit Sub
    Dim Y As Integer, X As Integer
    Y = ListView5.ListItems.Count
    For X = 1 To Y
        If ListView5.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(0).Text = ListView5.ListItems.Item(X)
    Me.txtCadMatriz(5).Text = ListView5.SelectedItem.ListSubItems.Item(1)
    Me.txtCadMatriz(6).Text = ListView5.SelectedItem.ListSubItems.Item(2)
    Me.txtCadMatriz(20).Text = ListView5.SelectedItem.ListSubItems.Item(3)
    txtCadMatriz(0).Enabled = False
    cmdCadastro(11).Enabled = False
End Sub
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE HISTÓRICO FUNCIONAL <<<<<<<<<<

Private Sub Form_Unload(Cancel As Integer)
    If vIntegra = "S" Then
        cnBancoTotvs.Close
        Set cnBancoTotvs = Nothing
    End If
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraExperiencia
    End If
End Sub

Private Sub ListView5_DblClick()
    If vEdi <> "N" Then
        AlteraHistorico
    End If
End Sub

Private Sub mskCadMatriz_LostFocus()
    If mskCadMatriz.Text = "" Then Exit Sub
    mskCadMatriz.PromptInclude = False
    If isCPF(mskCadMatriz.Text) = False Then
        mobjMsg.Abrir "CPF é inválido!", Ok, critico, "Atenção"
        mskCadMatriz.SetFocus
    End If
    If Status = "novo" Then
        ResultPesq "novo"
    Else
        ResultPesq "editar"
    End If
    mskCadMatriz.PromptInclude = False
End Sub

Private Sub Timer1_Timer()
    Avaliador "candidato"
    Frame9.Visible = False
    WebBrowser1.Visible = False
    Timer1.Enabled = False
    'chameleonButton1.SetFocus
End Sub

Private Sub txtCadMatriz_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaHistorico
        End If
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCargo 8
        End If
    Case 10
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCurso
            CompoeComboNivel cboCadMatriz(5), txtCadMatriz(10)
        End If
    Case 12
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaEscolaridade
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaCargo(indice As Integer)
    Dim X As Integer
    SqlCargos = "Select * from tbCargos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by codcargo"
    rsCargos.Open SqlCargos, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsCargos.EOF Then rsCargos.MoveFirst
    
    If indice = 5 Then rsCargos.Find "codcargo=" & "'" & Val(Me.txtCadMatriz(5)) & "'"
    If indice = 8 Then rsCargos.Find "codcargo=" & "'" & Val(Me.txtCadMatriz(8)) & "'"
    
    If rsCargos.EOF Then
        If indice = 5 Then txtCadMatriz(5).Text = Format(txtCadMatriz(5), "000000") & ""
        If indice = 8 Then txtCadMatriz(8).Text = Format(txtCadMatriz(8), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Cargo não cadastrado", Ok, critico, "Atenção"
            If indice = 5 Then txtCadMatriz(6) = ""
            If indice = 8 Then txtCadMatriz(9) = ""
        End If
    Else
        If indice = 5 Then txtCadMatriz(5).Text = Format(rsCargos.Fields(0), "000000") & ""
        If indice = 5 Then txtCadMatriz(6).Text = rsCargos.Fields(2)
        If indice = 8 Then txtCadMatriz(8).Text = Format(rsCargos.Fields(0), "000000") & ""
        If indice = 8 Then txtCadMatriz(9).Text = rsCargos.Fields(2)
    End If
    rsCargos.Close
    Set rsCargos = Nothing
End Sub

Private Sub ChamaGridCargo(indice As Integer)
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbCargos where codcoligada ='" & vCodcoligada & "' and ativo = 'S' order by nomecargo"
    procnom = "nomeCargo"
    campo = 2
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Cargo"
    Pesquisa = frmCandidatos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecargo=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If indice = 5 Then txtCadMatriz(5).Text = Format(rsLocal.Fields(0), "000000")
            If indice = 8 Then txtCadMatriz(8).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaCurso()
    Dim X As Integer
    SqlCursos = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbTreinamentos.codtreinamento"
    rsCursos.Open SqlCursos, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsCursos.EOF Then rsCursos.MoveFirst
    rsCursos.Find "codtreinamento=" & "'" & Val(Me.txtCadMatriz(10)) & "'"
    If rsCursos.EOF Then
        txtCadMatriz(10).Text = Format(txtCadMatriz(10), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Curso/Treinamento não cadastrado", Ok, critico, "Atenção"
            txtCadMatriz(11) = ""
        End If
    Else
        txtCadMatriz(10).Text = Format(rsCursos.Fields(0), "000000") & ""
        txtCadMatriz(11).Text = rsCursos.Fields(1)
    End If
    rsCursos.Close
    Set rsCursos = Nothing
End Sub

Private Sub ChamaGridCurso()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and introdutorio = 'N' order by tbTreinamentos.nometreinamento"
    procnom = "nometreinamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Treinamento"
    Pesquisa = frmCandidatos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nometreinamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadMatriz(10).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaEscolaridade()
    Dim X As Integer
    sqlEscolaridade = "Select * from tbEscolaridade where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbEscolaridade.codescolaridade"
    rsEscolaridade.Open sqlEscolaridade, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsEscolaridade.EOF Then rsEscolaridade.MoveFirst
    rsEscolaridade.Find "codescolaridade=" & "'" & Val(Me.txtCadMatriz(12)) & "'"
    If rsEscolaridade.EOF Then
        txtCadMatriz(12).Text = Format(txtCadMatriz(12), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Formação escolar não cadastrada", Ok, critico, "Atenção"
            txtCadMatriz(13) = ""
        End If
    Else
        txtCadMatriz(12).Text = Format(rsEscolaridade.Fields(0), "000000") & ""
        txtCadMatriz(13).Text = rsEscolaridade.Fields(1)
    End If
    rsEscolaridade.Close
    Set rsEscolaridade = Nothing
End Sub

Private Sub ChamaGridEscolaridade()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbEscolaridade where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbEscolaridade.codescolaridade"
    procnom = "nomeescolaridade"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de formação escolar"
    Pesquisa = frmCandidatos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomeescolaridade=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadMatriz(12).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaHistorico()
    Dim X As Integer
    SqlHistorico = "Select tbMatriz.codmatriz,tbMatriz.codcargo,tbMatriz.nivel,tbcargos.nomecargo from tbMatriz,tbcargos where tbmatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo and tbmatriz.ativo = 'S' order by tbMatriz.codmatriz"
    rsHistorico.Open SqlHistorico, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsHistorico.EOF Then rsHistorico.MoveFirst
    rsHistorico.Find "codmatriz=" & "'" & Val(Me.txtCadMatriz(0)) & "'"
    If rsHistorico.EOF Then
        txtCadMatriz(0).Text = Format(txtCadMatriz(0), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Matriz não cadastrada", Ok, critico, "Atenção"
            txtCadMatriz(5) = ""
            txtCadMatriz(6) = ""
            txtCadMatriz(20) = ""
        End If
    Else
        txtCadMatriz(0).Text = Format(rsHistorico.Fields(0), "000000") & ""
        txtCadMatriz(5).Text = Format(rsHistorico.Fields(1), "000000") & ""
        txtCadMatriz(6).Text = rsHistorico.Fields(3)
        txtCadMatriz(20).Text = rsHistorico.Fields(2)
    End If
    rsHistorico.Close
    Set rsHistorico = Nothing
End Sub

Private Sub ChamaGridHistorico()
On Error Resume Next
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select a.codmatriz,d.nomecargo,a.nivel,b.nomedepartamento,c.nomesetor from tbmatriz as a inner join tbdepartamentos as b on a.codcoligada = b.codcoligada inner join tbsetores as c on a.codcoligada = c.codcoligada inner join tbcargos as d on a.codcoligada = d.codcoligada " & _
            "where a.codcoligada = '" & vCodcoligada & "' and a.coddepartamento = b.coddepartamento and a.codsetor = c.codsetor and a.codcargo = d.codcargo and a.ativo = 'S' order by d.nomecargo,a.nivel"
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
            txtCadMatriz(0).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarColaborador As New ADODB.Recordset
    Dim SqlSalvarColaborador As String
    
    Dim rsSalvarColaboradorTotvs As New ADODB.Recordset
    Dim SqlSalvarColaboradorTotvs As String
    
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    Dim Y As Integer
    cnBanco.BeginTrans
   
    mskCadMatriz.PromptInclude = False

    If Status = "editar" Then
        SqlSalvarColaborador = "select * from tbColaboradores where codcoligada = '" & vCodcoligada & "' and id = '" & lblID & "'"
        rsSalvarColaborador.Open SqlSalvarColaborador, cnBanco, adOpenKeyset, adLockOptimistic
    Else
        SqlSalvarColaborador = "select * from tbColaboradores where codcoligada = '" & vCodcoligada & "' and cpf = '" & mskCadMatriz & "' and tipo = '" & "candidato" & "'"
        rsSalvarColaborador.Open SqlSalvarColaborador, cnBanco, adOpenKeyset, adLockOptimistic
        If rsSalvarColaborador.EOF Then
            rsSalvarColaborador.AddNew
        Else
            mobjMsg.Abrir "Já existe cadastrado um CANDIDATO com o mesmo CPF", Ok, critico, "Atenção"
            cnBanco.RollbackTrans
            rsSalvarColaborador.Close
            Set rsSalvarColaborador = Nothing
            Exit Sub
        End If
    End If
    If rsSalvarColaborador.EOF Then rsSalvarColaborador.AddNew
    rsSalvarColaborador.Fields(0) = mskCadMatriz 'cpf
    rsSalvarColaborador.Fields(1) = "-" 'código do colaborador
    rsSalvarColaborador.Fields(2) = DTPicker4 'data de cadastro
    rsSalvarColaborador.Fields(3) = txtCadMatriz(3) 'nome do colaborador
    rsSalvarColaborador.Fields(4) = DTPicker1 'data de nascimento do colaborador
    rsSalvarColaborador.Fields(5) = cboCadMatriz(0) 'sexo
    rsSalvarColaborador.Fields(6) = cboCadMatriz(1) 'estado civil
    rsSalvarColaborador.Fields(7) = txtCadMatriz(7) 'nacionalidade
    rsSalvarColaborador.Fields(8) = txtCadMatriz(14) 'naturalidade
    rsSalvarColaborador.Fields(9) = cboCadMatriz(4) 'uf da naturalidade
    rsSalvarColaborador.Fields(10) = txtCadMatriz(15) 'ctps numero
    rsSalvarColaborador.Fields(11) = txtCadMatriz(16) 'ctps serie
    rsSalvarColaborador.Fields(12) = txtCadMatriz(17) 'cnh numero
    rsSalvarColaborador.Fields(13) = txtCadMatriz(18) 'cnh tipo
    rsSalvarColaborador.Fields(20) = txtCadMatriz(19) 'observacao
    If Check1.Value = 1 Then rsSalvarColaborador.Fields(17) = "S" Else rsSalvarColaborador.Fields(17) = "N" 'ativo
    If Label41 <> "" Then rsSalvarColaborador.Fields(18) = RemoveMask(Label41) Else rsSalvarColaborador.Fields(18) = 0 'média geral
    rsSalvarColaborador.Fields(19) = Label53 'caminho da foto
    rsSalvarColaborador(21) = ""
    rsSalvarColaborador(22) = txtCadMatriz(23) ' email
    rsSalvarColaborador(23) = "candidato" ' Tipo
    rsSalvarColaborador(24) = txtCadMatriz(2) ' Telefone
    rsSalvarColaborador(25) = txtCadMatriz(21) ' Celular
    rsSalvarColaborador(31) = vCodcoligada ' Codigo da coligada
    For Y = 0 To 3
        If chkAvaliador(Y).Value = 1 Then
            If chkAvaliador(Y).Caption = "Experiência:" Then rsSalvarColaborador.Fields(21) = rsSalvarColaborador.Fields(21) & "E"
            If chkAvaliador(Y).Caption = "Habilidades:" Then rsSalvarColaborador.Fields(21) = rsSalvarColaborador.Fields(21) & "H"
            If chkAvaliador(Y).Caption = "Cursos/treinamentos:" Then rsSalvarColaborador.Fields(21) = rsSalvarColaborador.Fields(21) & "T"
            If chkAvaliador(Y).Caption = "Formação escolar:" Then rsSalvarColaborador.Fields(21) = rsSalvarColaborador.Fields(21) & "F"
        End If
    Next
    
    rsSalvarColaborador.Update
    If Status = "novo" Then lblID = rsSalvarColaborador.Fields(29)

    '>>>>>> GRAVAR EXPERIENCIA <<<<<<<<<
    sqlDeletar = "Delete from tbColaboradoresExp where codcoligada = '" & vCodcoligada & "' and tbColaboradoresExp.cpf = '" & mskCadMatriz.Text & "' and tipo = 'candidato'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView1.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbColaboradoresExp"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = mskCadMatriz.Text
            rsSalvar.Fields(1) = "candidato"
            rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(3)
            rsSalvar.Fields(3) = Val(ListView1.ListItems.Item(X))
            rsSalvar.Fields(4) = ListView1.SelectedItem.ListSubItems.Item(2)
            rsSalvar.Fields(5) = vCodcoligada 'Codigo da coligada
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    '>>>>>> GRAVAR HABILIDADE <<<<<<<<<
    If ListView2.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbColaboradoresHab where codcoligada = '" & vCodcoligada & "' and tbColaboradoresHab.cpf = '" & mskCadMatriz.Text & "' and tbColaboradoresHab.codmatriz = '" & Mid$(txtCadMatriz(4), 1, 6) & "' and tipo = 'candidato'"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView2.ListItems.Count
            ListView2.ListItems.Item(X).Selected = True
            If ListView2.ListItems.Item(X).Checked = True Then
                'On Error GoTo TrataErro1
            
                rsSalvar.Find "codhabilidade=" & "'" & Val(ListView2.ListItems.Item(X)) & "'"
                If rsSalvar.EOF Then rsSalvar.AddNew
                'rsSalvar.AddNew
                rsSalvar.Fields(0) = mskCadMatriz.Text
                rsSalvar.Fields(1) = "candidato"
                rsSalvar.Fields(2) = Val(ListView2.ListItems.Item(X))
                rsSalvar.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(3)
                rsSalvar.Fields(4) = Val(Mid$(txtCadMatriz(4), 1, 6))
                rsSalvar.Fields(5) = vCodcoligada ' Codigo da coligada
            End If
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    '>>>>>> GRAVAR CURSO/TREINAMENTO <<<<<<<<<
    sqlDeletar = "Delete from tbColaboradoresCur where codcoligada = '" & vCodcoligada & "' and tbColaboradoresCur.cpf = '" & mskCadMatriz.Text & "' and tipo = 'candidato'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView3.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbColaboradoresCur"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView3.ListItems.Count
            ListView3.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = mskCadMatriz.Text
            rsSalvar.Fields(1) = "candidato"
            rsSalvar.Fields(2) = Val(ListView3.ListItems.Item(X))
            rsSalvar.Fields(3) = ListView3.SelectedItem.ListSubItems.Item(2)
            If ListView3.SelectedItem.ListSubItems.Item(3) <> "-" Then rsSalvar.Fields(5) = Val(Mid$(ListView3.SelectedItem.ListSubItems.Item(3), 1, 2)) Else rsSalvar.Fields(5) = 0
            rsSalvar.Fields(6) = vCodcoligada ' Codigo da coligada
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    '>>>>>> GRAVAR ESCOLARIDADE <<<<<<<<<
    sqlDeletar = "Delete from tbColaboradoresEsc where codcoligada = '" & vCodcoligada & "' and tbColaboradoresEsc.cpf = '" & mskCadMatriz.Text & "' and tipo = 'candidato'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView4.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbColaboradoresEsc"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView4.ListItems.Count
            ListView4.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = mskCadMatriz.Text
            rsSalvar.Fields(1) = "candidato"
            rsSalvar.Fields(2) = Val(ListView4.ListItems.Item(X))
            rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    '>>>>>> GRAVAR CARGOS PRETENDIDOS <<<<<<<<<
    sqlDeletar = "Delete from tbColaboradoresHist where codcoligada ='" & vCodcoligada & "' and tbColaboradoresHist.cpf = '" & mskCadMatriz.Text & "' and tipo = 'candidato'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView5.ListItems.Count > 0 Then
        SqlSalvar = "Select * from tbColaboradoresHist"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
        For X = 1 To ListView5.ListItems.Count
            ListView5.ListItems.Item(X).Selected = True
            rsSalvar.AddNew
            rsSalvar.Fields(0) = mskCadMatriz.Text
            rsSalvar.Fields(1) = Val(ListView5.ListItems.Item(X))
            rsSalvar.Fields(2) = Date
            rsSalvar.Fields(5) = ListView5.SelectedItem.ListSubItems.Item(4)
            rsSalvar.Fields(6) = ListView5.SelectedItem.ListSubItems.Item(5)
            rsSalvar.Fields(7) = "candidato"
            rsSalvar.Fields(11) = vCodcoligada 'Codigo da coligada
        Next
        rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    
    mskCadMatriz.PromptInclude = True
    
    '>>>>>> GRAVAR DADOS INTEGRAÇÃO TOTVS <<<<<<<<<
    sqlDeletar = "Delete from tbColaboradoresIntTotvs where codcoligada = '" & vCodcoligada & "' and tbColaboradoresIntTotvs.id = '" & Val(lblID.Caption) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvarColaboradorTotvs = "select * from tbColaboradoresIntTotvs"
    rsSalvarColaboradorTotvs.Open SqlSalvarColaboradorTotvs, cnBanco, adOpenKeyset, adLockOptimistic
    rsSalvarColaboradorTotvs.AddNew
    rsSalvarColaboradorTotvs.Fields(0) = Val(lblID) 'Identificador do colaborador
    rsSalvarColaboradorTotvs.Fields(1) = "1.1" 'código do colaborador
    rsSalvarColaboradorTotvs.Fields(2) = txtCons(0) 'Sexo
    rsSalvarColaboradorTotvs.Fields(3) = txtCons(1) 'Grau de instrução
    rsSalvarColaboradorTotvs.Fields(4) = txtCons(2) 'Tipo de admissão
    rsSalvarColaboradorTotvs.Fields(5) = txtCons(3) 'Motivo da admissão
    rsSalvarColaboradorTotvs.Fields(6) = txtCons(4) 'forma de recebimento
    rsSalvarColaboradorTotvs.Fields(7) = txtCons(5) 'situação
    rsSalvarColaboradorTotvs.Fields(8) = txtCons(6) 'Tipo de funcionario
    rsSalvarColaboradorTotvs.Fields(9) = txtCons(7) 'Horario de trabalho
    rsSalvarColaboradorTotvs.Fields(10) = txtCons(8) 'Função
    rsSalvarColaboradorTotvs.Fields(11) = txtCons(9) 'Seção
    rsSalvarColaboradorTotvs.Fields(12) = txtCons(10) 'Contribuição sindical
    rsSalvarColaboradorTotvs.Fields(13) = txtCons(11) 'Rais
    rsSalvarColaboradorTotvs.Fields(14) = txtCons(12) 'Membro sindical
    rsSalvarColaboradorTotvs.Fields(15) = vCodcoligada 'Codigo da coligada
    
    rsSalvarColaboradorTotvs.Update
    rsSalvarColaboradorTotvs.Close
    Set rsSalvarColaboradorTotvs = Nothing
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    cnBanco.CommitTrans
    rsSalvarColaborador.Close
    Set rsSalvarColaborador = Nothing
    
    mobjMsg.Abrir "Os dados do Candidato foram salvos com sucesso", Ok, critico, "Atenção"
    AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    'CPF
    mskCadMatriz.PromptInclude = False
    If mskCadMatriz.Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.mskCadMatriz.Tag, Ok, critico, "Atenção"
        Me.mskCadMatriz.SetFocus
        Exit Function
    End If
    mskCadMatriz.PromptInclude = True
    
    'NOME
    If txtCadMatriz(3).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(3).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(3).SetFocus
        Exit Function
    End If
    
    'MATRIZ
    If txtCadMatriz(4).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(4).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(0).SetFocus
        Exit Function
    End If
    
    'SEXO
    If cboCadMatriz(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.cboCadMatriz(0).Tag, Ok, critico, "Atenção"
        Me.cboCadMatriz(0).SetFocus
        Exit Function
    End If
    
    'ESTADO CIVIL
    If cboCadMatriz(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.cboCadMatriz(1).Tag, Ok, critico, "Atenção"
        Me.cboCadMatriz(1).SetFocus
        Exit Function
    End If
    
    'NACIONALIDADE
    If txtCadMatriz(7).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(7).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(7).SetFocus
        Exit Function
    End If
    
    'NATURALIDADE
    If txtCadMatriz(14).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(14).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(14).SetFocus
        Exit Function
    End If
    
    'UF
    If cboCadMatriz(4).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.cboCadMatriz(4).Tag, Ok, critico, "Atenção"
        Me.cboCadMatriz(4).SetFocus
        Exit Function
    End If
    
    'TELEFONE
    If txtCadMatriz(2).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(2).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(2).SetFocus
        Exit Function
    End If
    'Irá realizar o procedimento abaixo somente se houver integração com Totvs
   ' If vIntegra = "S" Then
   '     For X = 0 To 12
   '         If Me.txtCons(X) = "" Then
   '             MsgBox "Favor informar o campo " & Me.txtCons(X).Tag, vbInformation, "Atenção"
   '             Me.txtCons(X).SetFocus
   '             Exit Function
   '         End If
   '     Next
   ' End If
    
    ValidaCampo = True
End Function

Private Sub ResultPesq(Status As String)
    If Status = "editar" Then
        SqlColaboradores = "Select * from tbColaboradores Where codcoligada ='" & vCodcoligada & "' and cpf = '" & Mid$(varGlobal, 1, 11) & "' and tipo = 'candidato' order by cpf"
        rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockReadOnly
        If rsColaboradores.RecordCount > 0 Then
            CompoeControles
        Else
            mobjMsg.Abrir "Candidato não encontrado", Ok, critico, "Atenção"
        End If
    ElseIf Status = "novo" Then
        SqlColaboradores = "Select * from tbColaboradores Where codcoligada ='" & vCodcoligada & "' and cpf = '" & mskCadMatriz & "' and tipo = 'candidato' or cpf = '" & mskCadMatriz & "' and tipo = 'colaborador' and ativo = 'S' order by cpf"
        rsColaboradores.Open SqlColaboradores, cnBanco, adOpenKeyset, adLockReadOnly
        If rsColaboradores.RecordCount > 0 Then
            mobjMsg.Abrir "Já existe um colaborador cadastrado no sistema com esse CPF", Ok, critico, "Atenção"
            mskCadMatriz.SetFocus
        End If
    End If
    rsColaboradores.Close
    Set rsColaboradores = Nothing
End Sub

Private Sub AtualizaListview()
'    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        mskCadMatriz.PromptInclude = False
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , mskCadMatriz) 'CPF
        mskCadMatriz.PromptInclude = True
        'ItemLst.SubItems(1) = txtCadMatriz(2).Text ' Registro
        ItemLst.SubItems(1) = txtCadMatriz(3).Text 'nome
        ItemLst.SubItems(2) = txtCadMatriz(15).Text 'CTPS nº
        ItemLst.SubItems(3) = txtCadMatriz(16).Text 'CTPS série
        ItemLst.SubItems(4) = Label41 'Média
        If Check1.Value = 0 Then 'Ativo
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "OK"
        End If
        ItemLst.SubItems(6) = ""
        For Y = 0 To 3
            If chkAvaliador(Y).Value = 1 Then
                If chkAvaliador(Y).Caption = "Experiência:" Then ItemLst.SubItems(6) = ItemLst.SubItems(6) & "E"
                If chkAvaliador(Y).Caption = "Habilidades:" Then ItemLst.SubItems(6) = ItemLst.SubItems(6) & "H"
                If chkAvaliador(Y).Caption = "Cursos/treinamentos:" Then ItemLst.SubItems(6) = ItemLst.SubItems(6) & "T"
                If chkAvaliador(Y).Caption = "Formação escolar:" Then ItemLst.SubItems(6) = ItemLst.SubItems(6) & "F"
            End If
        Next
        ItemLst.SubItems(7) = Mid$(txtCadMatriz(4), 8, 50)
        If RemoveMask(ItemLst.SubItems(5)) >= MediaGlobal Then
            ItemLst.ListSubItems(4).ForeColor = &H8000&
        Else
            ItemLst.ListSubItems(4).ForeColor = &HC0&
        End If
        Status = "editar"
    Else
        'frmPesqcolaboradores.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(2).Text ' Registro
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(3).Text 'Nome
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadMatriz(15).Text 'CTPS nº
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtCadMatriz(16).Text 'CTPS série
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = Label41 'Média
        If Check2.Value = 0 Then 'Ativo
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "OK"
        End If
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
        For Y = 0 To 3
            If chkAvaliador(Y).Value = 1 Then
                If chkAvaliador(Y).Caption = "Experiência:" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) & "E"
                If chkAvaliador(Y).Caption = "Habilidades:" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) & "H"
                If chkAvaliador(Y).Caption = "Cursos/treinamentos:" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) & "T"
                If chkAvaliador(Y).Caption = "Formação escolar:" Then MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) & "F"
            End If
        Next
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = Mid$(txtCadMatriz(4), 8, 50)
        If RemoveMask(MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)) >= MediaGlobal Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &H8000&
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ForeColor = &HC0&
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub MudaCorLV5()
    Dim ItemLst5 As ListItem

    Y = ListView5.ListItems.Count
    For X = 1 To Y
        ListView5.ListItems.Item(X).Selected = True
        
        If ListView5.SelectedItem.ListSubItems.Item(4) <> "S" Then
            ListView5.ListItems.Item(X).Bold = False
            ListView5.SelectedItem.ListSubItems.Item(1).Bold = False
            ListView5.SelectedItem.ListSubItems.Item(2).Bold = False
            ListView5.SelectedItem.ListSubItems.Item(3).Bold = False
            ListView5.SelectedItem.ListSubItems.Item(4).Bold = False
            ListView5.SelectedItem.ListSubItems.Item(5).Bold = False
            'ListView5.SelectedItem.ListSubItems.Item(6).Bold = False
        
            ListView5.ListItems.Item(X).ForeColor = &H800000
            ListView5.SelectedItem.ListSubItems.Item(1).ForeColor = &H800000
            ListView5.SelectedItem.ListSubItems.Item(2).ForeColor = &H800000
            ListView5.SelectedItem.ListSubItems.Item(3).ForeColor = &H800000
            ListView5.SelectedItem.ListSubItems.Item(4).ForeColor = &H800000
            ListView5.SelectedItem.ListSubItems.Item(5).ForeColor = &H800000
            'ListView5.SelectedItem.ListSubItems.Item(6).ForeColor = &H800000
        Else
            ListView5.ListItems.Item(X).Bold = True
            ListView5.SelectedItem.ListSubItems.Item(1).Bold = True
            ListView5.SelectedItem.ListSubItems.Item(2).Bold = True
            ListView5.SelectedItem.ListSubItems.Item(3).Bold = True
            ListView5.SelectedItem.ListSubItems.Item(4).Bold = True
            ListView5.SelectedItem.ListSubItems.Item(5).Bold = True
            'ListView5.SelectedItem.ListSubItems.Item(6).Bold = True
        
            ListView5.ListItems.Item(X).ForeColor = &H8000&
            ListView5.SelectedItem.ListSubItems.Item(1).ForeColor = &H8000&
            ListView5.SelectedItem.ListSubItems.Item(2).ForeColor = &H8000&
            ListView5.SelectedItem.ListSubItems.Item(3).ForeColor = &H8000&
            ListView5.SelectedItem.ListSubItems.Item(4).ForeColor = &H8000&
            ListView5.SelectedItem.ListSubItems.Item(5).ForeColor = &H8000&
            'ListView5.SelectedItem.ListSubItems.Item(6).ForeColor = &H8000&
        End If
    Next
End Sub

Private Sub MudaCorLV3()
    Y = ListView3.ListItems.Count
    For X = 1 To Y
        ListView3.ListItems.Item(X).Selected = True
        
        If ListView3.SelectedItem.ListSubItems.Item(2) <> "C" Then
            If ListView3.SelectedItem.ListSubItems.Item(2) = "SA" Then
                ListView3.ListItems.Item(X).ForeColor = &H80000011
                ListView3.SelectedItem.ListSubItems.Item(1).ForeColor = &H80000011
                ListView3.SelectedItem.ListSubItems.Item(2).ForeColor = &H80000011
                ListView3.SelectedItem.ListSubItems.Item(3).ForeColor = &H80000011
            Else
                ListView3.ListItems.Item(X).ForeColor = &H8080FF
                ListView3.SelectedItem.ListSubItems.Item(1).ForeColor = &H8080FF
                ListView3.SelectedItem.ListSubItems.Item(2).ForeColor = &H8080FF
                ListView3.SelectedItem.ListSubItems.Item(3).ForeColor = &H8080FF
            End If
        Else
            ListView3.ListItems.Item(X).ForeColor = &H800000
            ListView3.SelectedItem.ListSubItems.Item(1).ForeColor = &H800000
            ListView3.SelectedItem.ListSubItems.Item(2).ForeColor = &H800000
            ListView3.SelectedItem.ListSubItems.Item(3).ForeColor = &H800000
        End If
    Next
End Sub

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
        For i = 4 To 4
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
            If ListView2.ListItems.Count >= 16 Then
                If txtLeft < 11000 Then .Left = txtLeft + 450 Else .Left = txtLeft - 140
            Else
                If txtLeft < 11000 Then .Left = txtLeft + 220 Else .Left = txtLeft - 140
            End If
            .Top = txtTop + 4320
            .Width = txtWidth - 530
            .Height = ListView2.SelectedItem.Height - 8
            
        End With
    End If
End Sub

Private Sub txtCons_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(0) <> "" Then CarregaComboTotvs "PCODSEXO", "CODINTERNO", txtCons(0).Text, Combo(1).Text, Index, "descricao"
        End If
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(1) <> "" Then CarregaComboTotvs "PCODINSTRUCAO", "CODINTERNO", txtCons(1).Text, Combo(2).Text, Index, "descricao"
        End If
    Case 2
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(2) <> "" Then CarregaComboTotvs "PTPADMISSAO", "CODINTERNO", txtCons(2).Text, Combo(3).Text, Index, "descricao"
        End If
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(3) <> "" Then CarregaComboTotvs "PMOTADMISSAO", "CODINTERNO", txtCons(3).Text, Combo(4).Text, Index, "descricao"
        End If
    Case 4
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(4) <> "" Then CarregaComboTotvs "PCODRECEB", "CODINTERNO", txtCons(4).Text, Combo(5).Text, Index, "descricao"
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(5) <> "" Then CarregaComboTotvs "PCODSITUACAO", "CODINTERNO", txtCons(5).Text, Combo(6).Text, Index, "descricao"
        End If
    Case 6
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(6) <> "" Then CarregaComboTotvs "PTPFUNC", "CODINTERNO", txtCons(6).Text, Combo(7).Text, Index, "descricao"
        End If
    Case 7
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(7) <> "" Then CarregaComboTotvs "AHORARIO", "CODIGO", txtCons(7).Text, Combo(8).Text, Index, "descricao"
        End If
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(8) <> "" Then CarregaComboTotvs "PFUNCAO", "CODIGO", txtCons(8).Text, Combo(9).Text, Index, "nome"
        End If
    Case 9
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(9) <> "" Then CarregaComboTotvs "PSECAO", "CODIGO", txtCons(9).Text, Combo(10).Text, Index, "descricao"
        End If
    Case 10
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(10) <> "" Then CarregaComboTotvs "PCODCTSIND", "CODINTERNO", txtCons(10).Text, Combo(11).Text, Index, "descricao"
        End If
    Case 11
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(11) <> "" Then CarregaComboTotvs "PCODSITRAIS", "CODINTERNO", txtCons(11).Text, Combo(12).Text, Index, "descricao"
        End If
    Case 12
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(12) <> "" Then CarregaComboTotvs "PSINDIC", "CODIGO", txtCons(12).Text, Combo(13).Text, Index, "nome"
        End If
    End Select
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
        ListView2.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView2.ListItems(ListView2.ListItems.Count) <> c_EntryTxt Then ListView2.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    If ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) = "-" And ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) < ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
        Exit Sub
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
        ListView2.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
    End If
    
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    ListView2.SetFocus
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

'FUNCAO PARA MUDAR TOOLTIPS
Private Sub MudaTool()
    On Error Resume Next
    Dim Ctl As Control
    Dim i As Integer
    With Me.cIpToolTips1
        .Create
        .Title = "Atenção:" 'Titulo do tooltip
        .MyIcon = itInfoIcon 'Icone do tooltip
        .BackColor = &H80000018  'Cor de fundo
        .ForeColor = &H800000    'Cor da letra e bordas
        For Each Ctl In Me.Controls
            If Ctl.Tag <> "" Then
                .AddTool Ctl, tfAbsolute, Replace(Ctl.Tag, "|", vbCrLf)
            End If
        Next
    End With
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(8).UseGreyscale = True
        cmdCadastro(8).DragMode = 1
        cmdCadastro(8).SpecialEffect = cbEngraved
    
        cmdCadastro(16).UseGreyscale = True
        cmdCadastro(16).DragMode = 1
        cmdCadastro(16).SpecialEffect = cbEngraved
    
        cmdCadastro(22).UseGreyscale = True
        cmdCadastro(22).DragMode = 1
        cmdCadastro(22).SpecialEffect = cbEngraved
    
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    
        cmdCadastro(7).UseGreyscale = True
        cmdCadastro(7).DragMode = 1
        cmdCadastro(7).SpecialEffect = cbEngraved
    
        cmdCadastro(10).UseGreyscale = True
        cmdCadastro(10).DragMode = 1
        cmdCadastro(10).SpecialEffect = cbEngraved
    
        cmdCadastro(21).UseGreyscale = True
        cmdCadastro(21).DragMode = 1
        cmdCadastro(21).SpecialEffect = cbEngraved
    
        cmdCadastro(12).UseGreyscale = True
        cmdCadastro(12).DragMode = 1
        cmdCadastro(12).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    
        cmdCadastro(9).UseGreyscale = True
        cmdCadastro(9).DragMode = 1
        cmdCadastro(9).SpecialEffect = cbEngraved
    
        cmdCadastro(20).UseGreyscale = True
        cmdCadastro(20).DragMode = 1
        cmdCadastro(20).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(14).UseGreyscale = True
        cmdCadastro(14).DragMode = 1
        cmdCadastro(14).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    
        cmdCadastro(5).UseGreyscale = True
        cmdCadastro(5).DragMode = 1
        cmdCadastro(5).SpecialEffect = cbEngraved
    
        cmdCadastro(6).UseGreyscale = True
        cmdCadastro(6).DragMode = 1
        cmdCadastro(6).SpecialEffect = cbEngraved
    
        cmdCadastro(19).UseGreyscale = True
        cmdCadastro(19).DragMode = 1
        cmdCadastro(19).SpecialEffect = cbEngraved
    
        cmdCadastro(13).UseGreyscale = True
        cmdCadastro(13).DragMode = 1
        cmdCadastro(13).SpecialEffect = cbEngraved
    End If
    If vAva = "N" Then
        chameleonButton1.UseGreyscale = True
        chameleonButton1.DragMode = 1
        chameleonButton1.SpecialEffect = cbEngraved
    End If
    If vIntegra = "S" Then SSTab1.TabEnabled(6) = True Else SSTab1.TabEnabled(6) = False
End Sub

Private Sub comporCombosTotvs()
    Dim X As Integer
    CompoeComboTotvs Combo(1), "PCODSEXO", "codinterno", "descricao"
    CompoeComboTotvs Combo(2), "PCODINSTRUCAO", "codinterno", "descricao"
    CompoeComboTotvs Combo(3), "PTPADMISSAO", "codinterno", "descricao"
    CompoeComboTotvs Combo(4), "PMOTADMISSAO", "codinterno", "descricao"
    CompoeComboTotvs Combo(5), "PCODRECEB", "codinterno", "descricao"
    CompoeComboTotvs Combo(6), "PCODSITUACAO", "codinterno", "descricao"
    CompoeComboTotvs Combo(7), "PTPFUNC", "codinterno", "descricao"
    CompoeComboTotvs Combo(8), "AHORARIO", "codigo", "descricao"
    CompoeComboTotvs Combo(9), "PFUNCAO", "codigo", "nome"
    CompoeComboTotvs Combo(10), "PSECAO", "codigo", "descricao"
    CompoeComboTotvs Combo(11), "PCODCTSIND", "codinterno", "descricao"
    CompoeComboTotvs Combo(12), "PCODSITRAIS", "codinterno", "descricao"
    CompoeComboTotvs Combo(13), "PSINDIC", "codigo", "nome"
    
    For X = 0 To Combo(10).ListCount - 1
        Combo(10).ListIndex = X
        If Combo(10).List(X) = "SGC" Then
            Combo(10).Text = Combo(10).List(X)
            Combo_Click (10)
            Combo(10).Enabled = False
            txtCons(9).Enabled = False
            Exit For
        End If
    Next

    For X = 0 To Combo(6).ListCount - 1
        Combo(6).ListIndex = X
        If Combo(6).List(X) = "Ativo" Then
            Combo(6).Text = Combo(6).List(X)
            Combo_Click (6)
            Combo(6).Enabled = False
            txtCons(5).Enabled = False
            Exit For
        End If
    Next
End Sub

Private Sub comporControlesTotvs()
    On Error Resume Next
    Dim rsContrTotvs As New ADODB.Recordset
    Dim SqlContrTotvs As String
        
    SqlContrTotvs = "select * from tbColaboradoresIntTotvs where codcoligada = '" & vCodcoligada & "' and id = '" & lblID & "'"
    rsContrTotvs.Open SqlContrTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    
    txtCons(0) = rsContrTotvs.Fields(2)
    txtCons(1) = rsContrTotvs.Fields(3)
    txtCons(2) = rsContrTotvs.Fields(4)
    txtCons(3) = rsContrTotvs.Fields(5)
    txtCons(4) = rsContrTotvs.Fields(6)
    If Combo(6) <> "Ativo" Then txtCons(5) = rsContrTotvs.Fields(7)
    txtCons(6) = rsContrTotvs.Fields(8)
    txtCons(7) = rsContrTotvs.Fields(9)
    txtCons(8) = rsContrTotvs.Fields(10)
    If Combo(10) <> "SGC" Then txtCons(9) = rsContrTotvs.Fields(11)
    txtCons(10) = rsContrTotvs.Fields(12)
    txtCons(11) = rsContrTotvs.Fields(13)
    txtCons(12) = rsContrTotvs.Fields(14)
    txtCons_KeyDown 0, 13, 0
    txtCons_KeyDown 1, 13, 1
    txtCons_KeyDown 2, 13, 2
    txtCons_KeyDown 3, 13, 3
    txtCons_KeyDown 4, 13, 4
    If Combo(6) <> "Ativo" Then txtCons_KeyDown 5, 13, 5
    txtCons_KeyDown 6, 13, 6
    txtCons_KeyDown 7, 13, 7
    txtCons_KeyDown 8, 13, 8
    If Combo(10) <> "SGC" Then txtCons_KeyDown 9, 13, 9
    txtCons_KeyDown 10, 13, 10
    txtCons_KeyDown 11, 13, 11
    txtCons_KeyDown 12, 13, 12
    rsContrTotvs.Close
    Set rsContrTotvs = Nothing
End Sub
