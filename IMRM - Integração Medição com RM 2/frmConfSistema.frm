VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmConfSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfSistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin IMRM.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   65
      Top             =   8160
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
      MICON           =   "frmConfSistema.frx":37E04
      PICN            =   "frmConfSistema.frx":37E20
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin IMRM.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   64
      Top             =   8160
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
      MICON           =   "frmConfSistema.frx":38AFA
      PICN            =   "frmConfSistema.frx":38B16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   13996
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
      TabCaption(0)   =   "Sincronização"
      TabPicture(0)   =   "frmConfSistema.frx":397F0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame27"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Parametrizações"
      TabPicture(1)   =   "frmConfSistema.frx":3980C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Empresa/Coligadas"
      TabPicture(2)   =   "frmConfSistema.frx":39828
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Servidor - email"
      TabPicture(3)   =   "frmConfSistema.frx":39844
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "Frame21"
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(3)=   "Frame9"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Menu"
      TabPicture(4)   =   "frmConfSistema.frx":39860
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      Begin VB.CommandButton Command4 
         Caption         =   "Envia EMAIL"
         Height          =   375
         Left            =   -74880
         TabIndex        =   181
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Frame Frame27 
         Caption         =   "Sincornizar tabelas "
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
         TabIndex        =   156
         Top             =   480
         Width           =   9135
         Begin VB.CommandButton Command3 
            Caption         =   "Iniciar sincronização"
            Height          =   375
            Left            =   120
            TabIndex        =   173
            Top             =   5280
            Width           =   2295
         End
         Begin VB.Frame Frame28 
            Caption         =   "Servidor a ser sincronizado (ONLINE)"
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
            TabIndex        =   167
            Top             =   2400
            Width           =   8895
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   9
               Left            =   6120
               TabIndex        =   170
               Top             =   600
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
               Height          =   255
               Left            =   6120
               OleObjectBlob   =   "frmConfSistema.frx":3987C
               TabIndex        =   180
               Top             =   360
               Width           =   2295
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   13
               Left            =   120
               TabIndex        =   168
               Tag             =   "Nome do Servidor Local"
               ToolTipText     =   "Nome do servidor local (SQL Server)"
               Top             =   600
               Width           =   3135
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   12
               Left            =   3480
               TabIndex        =   169
               Tag             =   "Nome do Banco Local"
               ToolTipText     =   "Nome do banco da ferramentaria que será criado no servidor local"
               Top             =   600
               Width           =   2535
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   11
               Left            =   120
               TabIndex        =   171
               Tag             =   "Usuário do Banco Local"
               ToolTipText     =   "Usuário do servidor local"
               Top             =   1200
               Width           =   3135
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   10
               Left            =   3480
               PasswordChar    =   "*"
               TabIndex        =   172
               Tag             =   "Senha do Usuário do Banco Local"
               ToolTipText     =   "Senha de acesso ao servidor local"
               Top             =   1200
               Width           =   2535
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":39902
               TabIndex        =   174
               Top             =   1560
               Width           =   2175
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
               Height          =   735
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":3996A
               TabIndex        =   175
               Top             =   1920
               Width           =   8655
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":39B64
               TabIndex        =   176
               Top             =   360
               Width           =   2775
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
               Height          =   255
               Left            =   3480
               OleObjectBlob   =   "frmConfSistema.frx":39BF4
               TabIndex        =   177
               Top             =   360
               Width           =   3855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":39C7A
               TabIndex        =   178
               Top             =   960
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
               Height          =   255
               Left            =   3480
               OleObjectBlob   =   "frmConfSistema.frx":39CE2
               TabIndex        =   179
               Top             =   960
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "frmConfSistema.frx":39D46
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   166
            Top             =   1800
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "frmConfSistema.frx":3AA10
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   165
            Top             =   1440
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "frmConfSistema.frx":3B6DA
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   164
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "frmConfSistema.frx":3C3A4
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   163
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3360
            Picture         =   "frmConfSistema.frx":3D06E
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   162
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Funções/Seções (Importar)"
            Height          =   255
            Left            =   240
            TabIndex        =   161
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Movimentos (Exportar)"
            Height          =   255
            Left            =   240
            TabIndex        =   160
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Produtos (Importar)"
            Height          =   255
            Left            =   240
            TabIndex        =   159
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Usuários (Importar)"
            Height          =   255
            Left            =   240
            TabIndex        =   158
            Top             =   840
            Width           =   3015
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Funcionários (Importar)"
            Height          =   255
            Left            =   240
            TabIndex        =   157
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Porta "
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
         Left            =   -69000
         TabIndex        =   111
         Top             =   480
         Width           =   3015
         Begin VB.TextBox txtEmail 
            Height          =   330
            Index           =   3
            Left            =   120
            TabIndex        =   112
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Estrutura do Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   9015
         Begin IMRM.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   4
            Left            =   1920
            TabIndex        =   96
            Tag             =   "Excluir"
            ToolTipText     =   "Excluir"
            Top             =   1680
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
            MICON           =   "frmConfSistema.frx":3DD38
            PICN            =   "frmConfSistema.frx":3DD54
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin IMRM.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   5
            Left            =   1320
            TabIndex        =   95
            Tag             =   "Novo"
            ToolTipText     =   "Novo"
            Top             =   1680
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
            MICON           =   "frmConfSistema.frx":3EA2E
            PICN            =   "frmConfSistema.frx":3EA4A
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
            Height          =   1215
            Left            =   7200
            TabIndex        =   90
            Top             =   360
            Width           =   1695
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
               Height          =   375
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":3F724
               TabIndex        =   92
               Top             =   600
               Width           =   1455
            End
         End
         Begin IMRM.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   720
            TabIndex        =   94
            Tag             =   "Editar"
            ToolTipText     =   "Editar"
            Top             =   1680
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
            MICON           =   "frmConfSistema.frx":3F77E
            PICN            =   "frmConfSistema.frx":3F79A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin IMRM.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   93
            Tag             =   "Incluir"
            ToolTipText     =   "Incluir"
            Top             =   1680
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
            MICON           =   "frmConfSistema.frx":40474
            PICN            =   "frmConfSistema.frx":40490
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            TabIndex        =   81
            Top             =   360
            Width           =   1335
            Begin IMRM.chameleonButton cmdCadastro 
               Height          =   255
               Index           =   6
               Left            =   840
               TabIndex        =   97
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
               MICON           =   "frmConfSistema.frx":4116A
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
               TabIndex        =   89
               Tag             =   "Ícone"
               ToolTipText     =   "Ícone"
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox Check9 
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
               Height          =   255
               Left            =   120
               TabIndex        =   88
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
            TabIndex        =   78
            Top             =   360
            Width           =   1215
            Begin VB.ComboBox Combo2 
               Enabled         =   0   'False
               Height          =   345
               ItemData        =   "frmConfSistema.frx":41186
               Left            =   120
               List            =   "frmConfSistema.frx":41193
               TabIndex        =   80
               Tag             =   "Tipo"
               Text            =   "TAB"
               ToolTipText     =   "Tipo"
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox Check8 
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
               Left            =   120
               TabIndex        =   79
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "     Botão "
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
            TabIndex        =   77
            Top             =   360
            Width           =   1335
            Begin VB.ComboBox Combo4 
               Enabled         =   0   'False
               Height          =   345
               Left            =   120
               TabIndex        =   87
               Tag             =   "Botão"
               ToolTipText     =   "Botão"
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox Check7 
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
               Height          =   255
               Left            =   120
               TabIndex        =   86
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
            TabIndex        =   76
            Top             =   360
            Width           =   1455
            Begin VB.ComboBox Combo3 
               Enabled         =   0   'False
               Height          =   345
               ItemData        =   "frmConfSistema.frx":411A6
               Left            =   120
               List            =   "frmConfSistema.frx":411A8
               TabIndex        =   85
               Tag             =   "Submenu"
               ToolTipText     =   "Submenu"
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox Check3 
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
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "     Menu "
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
            TabIndex        =   75
            Top             =   360
            Width           =   1095
            Begin VB.ComboBox Combo1 
               Enabled         =   0   'False
               Height          =   345
               ItemData        =   "frmConfSistema.frx":411AA
               Left            =   120
               List            =   "frmConfSistema.frx":411CC
               TabIndex        =   83
               Tag             =   "Menu"
               Text            =   "01"
               ToolTipText     =   "Menu"
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox Check2 
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
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   91
            Tag             =   "Nome"
            ToolTipText     =   "Nome"
            Top             =   1320
            Width           =   6855
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4935
            Left            =   120
            TabIndex        =   73
            Top             =   2400
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   8705
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1080
            Width           =   735
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   11456
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
         TabCaption(0)   =   "Empresa/coligada ativa"
         TabPicture(0)   =   "frmConfSistema.frx":411F8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Empresas/Coligadas"
         TabPicture(1)   =   "frmConfSistema.frx":41214
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ListView3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "imgColigada"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chameleonButton4"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chameleonButton5"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin IMRM.chameleonButton chameleonButton5 
            Height          =   615
            Left            =   720
            TabIndex        =   69
            Top             =   480
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
            MICON           =   "frmConfSistema.frx":41230
            PICN            =   "frmConfSistema.frx":4124C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin IMRM.chameleonButton chameleonButton4 
            Height          =   615
            Left            =   120
            TabIndex        =   68
            Top             =   480
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
            MICON           =   "frmConfSistema.frx":41F26
            PICN            =   "frmConfSistema.frx":41F42
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ImageList imgColigada 
            Left            =   240
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
                  Picture         =   "frmConfSistema.frx":42C1C
                  Key             =   "OK"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmConfSistema.frx":4362E
                  Key             =   "EXC"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   5175
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   9128
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
            Height          =   5895
            Left            =   -74880
            TabIndex        =   19
            Top             =   360
            Width           =   8775
            Begin VB.Frame Frame29 
               Caption         =   "Série "
               Height          =   735
               Left            =   120
               TabIndex        =   182
               Top             =   3720
               Width           =   975
               Begin VB.TextBox txtDadosEmpresa 
                  Height          =   375
                  Index           =   12
                  Left            =   120
                  MaxLength       =   3
                  TabIndex        =   183
                  ToolTipText     =   "Deve conter 3 caracteres"
                  Top             =   240
                  Width           =   735
               End
            End
            Begin IMRM.chameleonButton cmdCadastro 
               Height          =   615
               Index           =   16
               Left            =   720
               TabIndex        =   67
               Top             =   4560
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
               MICON           =   "frmConfSistema.frx":44040
               PICN            =   "frmConfSistema.frx":4405C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin IMRM.chameleonButton cmdCadastro 
               Height          =   615
               Index           =   15
               Left            =   120
               TabIndex        =   66
               Top             =   4560
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
               MICON           =   "frmConfSistema.frx":44D36
               PICN            =   "frmConfSistema.frx":44D52
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   10
               Left            =   3720
               TabIndex        =   38
               Top             =   3240
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
               Height          =   255
               Left            =   3360
               OleObjectBlob   =   "frmConfSistema.frx":45A2C
               TabIndex        =   53
               Top             =   3360
               Width           =   375
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   8
               Left            =   3720
               TabIndex        =   36
               Top             =   2880
               Width           =   1935
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
               Height          =   255
               Left            =   3360
               OleObjectBlob   =   "frmConfSistema.frx":45A8C
               TabIndex        =   52
               Top             =   3000
               Width           =   495
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   5
               Left            =   1320
               TabIndex        =   33
               Top             =   2160
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   4
               Left            =   2760
               TabIndex        =   32
               Top             =   1800
               Width           =   1575
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
               Height          =   255
               Left            =   2280
               OleObjectBlob   =   "frmConfSistema.frx":45AEC
               TabIndex        =   51
               Top             =   1920
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45B4C
               TabIndex        =   50
               Top             =   3360
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45BAE
               TabIndex        =   49
               Top             =   3000
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45C18
               TabIndex        =   48
               Top             =   2640
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45C7A
               TabIndex        =   47
               Top             =   2280
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45CDE
               TabIndex        =   46
               Top             =   1920
               Width           =   495
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45D3E
               TabIndex        =   45
               Top             =   1560
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45DA4
               TabIndex        =   44
               Top             =   1200
               Width           =   855
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45E0A
               TabIndex        =   43
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtDadosEmpresa 
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   1320
               TabIndex        =   26
               Tag             =   "Código da coligada"
               ToolTipText     =   "Código da coligada"
               Top             =   360
               Width           =   735
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":45E74
               TabIndex        =   42
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   0
               Left            =   2160
               TabIndex        =   27
               Tag             =   "Razão social"
               ToolTipText     =   "Razão social"
               Top             =   360
               Width           =   3495
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   28
               Top             =   720
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   2
               Left            =   1320
               TabIndex        =   29
               Top             =   1080
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   3
               Left            =   1320
               TabIndex        =   30
               Top             =   1440
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   6
               Left            =   1320
               TabIndex        =   34
               Top             =   2520
               Width           =   4335
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   7
               Left            =   1320
               TabIndex        =   35
               Top             =   2880
               Width           =   1935
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   9
               Left            =   1320
               TabIndex        =   37
               Top             =   3240
               Width           =   1935
            End
            Begin VB.ComboBox cboDadosEmpresa 
               Height          =   345
               ItemData        =   "frmConfSistema.frx":45EE6
               Left            =   1320
               List            =   "frmConfSistema.frx":45F3B
               TabIndex        =   31
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
               TabIndex        =   20
               Top             =   240
               Width           =   2895
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   13
                  Left            =   720
                  TabIndex        =   71
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
                  MICON           =   "frmConfSistema.frx":45FAB
                  PICN            =   "frmConfSistema.frx":45FC7
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   12
                  Left            =   120
                  TabIndex        =   70
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
                  MICON           =   "frmConfSistema.frx":46CA1
                  PICN            =   "frmConfSistema.frx":46CBD
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
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2775
                  Left            =   120
                  ScaleHeight     =   2715
                  ScaleWidth      =   2595
                  TabIndex        =   39
                  Top             =   240
                  Width           =   2655
                  Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
                     Height          =   2655
                     Left            =   120
                     Top             =   0
                     Width           =   2415
                     _ExtentX        =   4260
                     _ExtentY        =   4683
                     Image           =   "frmConfSistema.frx":47997
                  End
                  Begin VB.Label Label59 
                     Alignment       =   2  'Center
                     Caption         =   "A Imagem não se encontra no local especificado"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   495
                     Left            =   240
                     TabIndex        =   21
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
               Left            =   120
               TabIndex        =   22
               Top             =   5280
               Visible         =   0   'False
               Width           =   5415
            End
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7215
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   12726
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
         TabCaption(0)   =   "Gerais"
         TabPicture(0)   =   "frmConfSistema.frx":479AF
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame8"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Integração"
         TabPicture(1)   =   "frmConfSistema.frx":479CB
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "Frame15"
         Tab(1).Control(2)=   "Check4"
         Tab(1).Control(3)=   "SSTab3"
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
            Height          =   615
            Left            =   -74880
            TabIndex        =   57
            Top             =   780
            Width           =   4095
            Begin VB.OptionButton optIntegra 
               Caption         =   "SQL Server"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Value           =   -1  'True
               Width           =   1575
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
            Height          =   615
            Left            =   -70680
            TabIndex        =   55
            Top             =   780
            Width           =   4455
            Begin VB.OptionButton chkIntegra 
               Caption         =   "Totvs RM"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   135
               Top             =   240
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton chkIntegra 
               Caption         =   "SAP"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   56
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Realizar integração"
            Height          =   255
            Left            =   -74880
            TabIndex        =   54
            Top             =   420
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
            Height          =   6735
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   8775
            Begin VB.Frame Frame24 
               Caption         =   "Período de avaliação de fornec."
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
               TabIndex        =   131
               Top             =   5880
               Visible         =   0   'False
               Width           =   3135
               Begin VB.ComboBox Combo5 
                  Height          =   345
                  ItemData        =   "frmConfSistema.frx":479E7
                  Left            =   960
                  List            =   "frmConfSistema.frx":47A0F
                  TabIndex        =   134
                  Text            =   "1"
                  Top             =   240
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
                  Height          =   255
                  Left            =   1800
                  OleObjectBlob   =   "frmConfSistema.frx":47A3A
                  TabIndex        =   133
                  Top             =   240
                  Width           =   855
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":47A9C
                  TabIndex        =   132
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame23 
               Caption         =   "Classificação "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3495
               Left            =   120
               TabIndex        =   120
               Top             =   2280
               Visible         =   0   'False
               Width           =   3135
               Begin MSComctlLib.ListView ListView2 
                  Height          =   1695
                  Left            =   120
                  TabIndex        =   126
                  Top             =   1680
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   2990
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
               Begin VB.TextBox Text6 
                  Alignment       =   2  'Center
                  DataField       =   "&H00008000&"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   21.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00008000&
                  Height          =   570
                  Left            =   1800
                  TabIndex        =   125
                  Text            =   "A"
                  Top             =   240
                  Width           =   735
               End
               Begin VB.TextBox Text5 
                  Height          =   330
                  Left            =   840
                  TabIndex        =   124
                  Top             =   480
                  Width           =   615
               End
               Begin VB.TextBox Text4 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   123
                  Top             =   480
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
                  Height          =   255
                  Left            =   840
                  OleObjectBlob   =   "frmConfSistema.frx":47B04
                  TabIndex        =   122
                  Top             =   240
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":47B64
                  TabIndex        =   121
                  Top             =   240
                  Width           =   615
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   11
                  Left            =   1920
                  TabIndex        =   127
                  Tag             =   "Excluir"
                  ToolTipText     =   "Excluir"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":47BC2
                  PICN            =   "frmConfSistema.frx":47BDE
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   14
                  Left            =   1320
                  TabIndex        =   128
                  Tag             =   "Novo"
                  ToolTipText     =   "Novo"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":488B8
                  PICN            =   "frmConfSistema.frx":488D4
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   17
                  Left            =   720
                  TabIndex        =   129
                  Tag             =   "Editar"
                  ToolTipText     =   "Editar"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":495AE
                  PICN            =   "frmConfSistema.frx":495CA
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   18
                  Left            =   120
                  TabIndex        =   130
                  Tag             =   "Incluir"
                  ToolTipText     =   "Incluir"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":4A2A4
                  PICN            =   "frmConfSistema.frx":4A2C0
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
            Begin VB.Frame Frame20 
               Caption         =   "E-mails Avaliação de Recebimento "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3495
               Left            =   3360
               TabIndex        =   103
               Top             =   2280
               Visible         =   0   'False
               Width           =   5295
               Begin VB.TextBox txtCadParametro 
                  Height          =   375
                  Index           =   1
                  Left            =   120
                  TabIndex        =   105
                  Top             =   480
                  Width           =   5055
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "frmConfSistema.frx":4AF9A
                  TabIndex        =   106
                  Top             =   240
                  Width           =   1815
               End
               Begin MSComctlLib.ListView ListView1 
                  Height          =   1695
                  Left            =   120
                  TabIndex        =   104
                  Top             =   1680
                  Width           =   5055
                  _ExtentX        =   8916
                  _ExtentY        =   2990
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
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   7
                  Left            =   1920
                  TabIndex        =   107
                  Tag             =   "Excluir"
                  ToolTipText     =   "Excluir"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":4B000
                  PICN            =   "frmConfSistema.frx":4B01C
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   8
                  Left            =   1320
                  TabIndex        =   108
                  Tag             =   "Novo"
                  ToolTipText     =   "Novo"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":4BCF6
                  PICN            =   "frmConfSistema.frx":4BD12
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   9
                  Left            =   720
                  TabIndex        =   109
                  Tag             =   "Editar"
                  ToolTipText     =   "Editar"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":4C9EC
                  PICN            =   "frmConfSistema.frx":4CA08
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin IMRM.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   10
                  Left            =   120
                  TabIndex        =   110
                  Tag             =   "Incluir"
                  ToolTipText     =   "Incluir"
                  Top             =   960
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
                  MICON           =   "frmConfSistema.frx":4D6E2
                  PICN            =   "frmConfSistema.frx":4D6FE
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
            Begin VB.Frame Frame19 
               Caption         =   "Média "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               TabIndex        =   100
               Top             =   1200
               Visible         =   0   'False
               Width           =   3135
               Begin VB.TextBox txtCadParametro 
                  Height          =   285
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   101
                  Tag             =   "Média para aprovação"
                  ToolTipText     =   "Média para aprovação"
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Aprovação (%):"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   102
                  Top             =   360
                  Width           =   1335
               End
            End
            Begin VB.Frame Frame18 
               Caption         =   "Data de início da Avaliação das OC - Ordens de Compra "
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
               TabIndex        =   98
               Top             =   1200
               Visible         =   0   'False
               Width           =   5295
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   375
                  Left            =   960
                  TabIndex        =   99
                  Top             =   360
                  Width           =   2775
                  _ExtentX        =   4895
                  _ExtentY        =   661
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   219742209
                  CurrentDate     =   42270
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
               TabIndex        =   40
               Top             =   240
               Visible         =   0   'False
               Width           =   5295
               Begin VB.CommandButton cmdCad 
                  Caption         =   "..."
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
                  Height          =   255
                  Index           =   2
                  Left            =   4680
                  TabIndex        =   15
                  Top             =   360
                  Width           =   375
               End
               Begin MSComDlg.CommonDialog cdlTXT2 
                  Left            =   4680
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
                  TabIndex        =   13
                  Top             =   0
                  Width           =   375
               End
               Begin VB.TextBox Text2 
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   120
                  TabIndex        =   14
                  Text            =   "Informe o caminho do executável: AtualizaSAF.exe"
                  Top             =   360
                  Width           =   4455
               End
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Exibir avisos ao logar"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   720
               Width           =   2175
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Bloquear colaborador que possui ferramentas vencidas"
               Height          =   375
               Left            =   120
               TabIndex        =   11
               Top             =   300
               Width           =   6135
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   5535
            Left            =   -74880
            TabIndex        =   59
            Top             =   1500
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   9763
            _Version        =   393216
            Tabs            =   1
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
            TabCaption(0)   =   "Dados do Servidor TOTVS"
            TabPicture(0)   =   "frmConfSistema.frx":4E3D8
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SkinLabel21"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SkinLabel22"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SkinLabel23"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SkinLabel24"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "txtIntegra(0)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txtIntegra(1)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txtIntegra(2)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtIntegra(3)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Frame25"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Command5"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).ControlCount=   10
            Begin VB.CommandButton Command5 
               Caption         =   "TESTE"
               Height          =   855
               Left            =   6240
               TabIndex        =   184
               Top             =   840
               Width           =   2295
            End
            Begin VB.Frame Frame25 
               Caption         =   "OFFLINE"
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
               TabIndex        =   140
               Top             =   2040
               Width           =   8535
               Begin VB.CommandButton Command2 
                  Caption         =   "Criar Banco/Tabelas"
                  Height          =   375
                  Left            =   4560
                  TabIndex        =   154
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Clonar Dados"
                  Height          =   375
                  Left            =   6600
                  TabIndex        =   153
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.Frame Frame26 
                  Caption         =   "Dados do servidor local"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2655
                  Left            =   120
                  TabIndex        =   142
                  Top             =   600
                  Width           =   8295
                  Begin VB.TextBox txtIntegra 
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   8
                     Left            =   6000
                     TabIndex        =   155
                     Text            =   "CORPORERM_OFF"
                     Top             =   600
                     Width           =   1935
                  End
                  Begin VB.TextBox txtIntegra 
                     Height          =   330
                     IMEMode         =   3  'DISABLE
                     Index           =   4
                     Left            =   3960
                     PasswordChar    =   "*"
                     TabIndex        =   146
                     Tag             =   "Senha do Usuário do Banco Local"
                     ToolTipText     =   "Senha de acesso ao servidor local"
                     Top             =   1200
                     Width           =   3975
                  End
                  Begin VB.TextBox txtIntegra 
                     Height          =   330
                     Index           =   5
                     Left            =   120
                     TabIndex        =   145
                     Tag             =   "Usuário do Banco Local"
                     ToolTipText     =   "Usuário do servidor local"
                     Top             =   1200
                     Width           =   3615
                  End
                  Begin VB.TextBox txtIntegra 
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   6
                     Left            =   3960
                     TabIndex        =   144
                     Tag             =   "Nome do Banco Local"
                     Text            =   "FERRAMENTARIA_OFF"
                     ToolTipText     =   "Nome do banco da ferramentaria que será criado no servidor local"
                     Top             =   600
                     Width           =   1935
                  End
                  Begin VB.TextBox txtIntegra 
                     Height          =   330
                     Index           =   7
                     Left            =   120
                     TabIndex        =   143
                     Tag             =   "Nome do Servidor Local"
                     ToolTipText     =   "Nome do servidor local (SQL Server)"
                     Top             =   600
                     Width           =   3615
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
                     Height          =   255
                     Left            =   120
                     OleObjectBlob   =   "frmConfSistema.frx":4E3F4
                     TabIndex        =   147
                     Top             =   1560
                     Width           =   2175
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
                     Height          =   615
                     Left            =   120
                     OleObjectBlob   =   "frmConfSistema.frx":4E45C
                     TabIndex        =   148
                     Top             =   1920
                     Width           =   7935
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
                     Height          =   255
                     Left            =   120
                     OleObjectBlob   =   "frmConfSistema.frx":4E5FC
                     TabIndex        =   149
                     Top             =   360
                     Width           =   1935
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
                     Height          =   255
                     Left            =   3960
                     OleObjectBlob   =   "frmConfSistema.frx":4E676
                     TabIndex        =   150
                     Top             =   360
                     Width           =   3855
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
                     Height          =   255
                     Left            =   120
                     OleObjectBlob   =   "frmConfSistema.frx":4E6FE
                     TabIndex        =   151
                     Top             =   960
                     Width           =   1215
                  End
                  Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
                     Height          =   255
                     Left            =   3960
                     OleObjectBlob   =   "frmConfSistema.frx":4E766
                     TabIndex        =   152
                     Top             =   960
                     Width           =   975
                  End
               End
               Begin VB.CheckBox Check11 
                  Caption         =   "Funcionar independente do banco do TOTVS RM"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   141
                  Top             =   240
                  Width           =   4575
               End
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   3000
               PasswordChar    =   "*"
               TabIndex        =   63
               Top             =   1560
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   62
               Top             =   1560
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   1
               Left            =   3000
               TabIndex        =   61
               Top             =   960
               Width           =   2655
            End
            Begin VB.TextBox txtIntegra 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   960
               Width           =   2655
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
               Height          =   255
               Left            =   3000
               OleObjectBlob   =   "frmConfSistema.frx":4E7D0
               TabIndex        =   139
               Top             =   1320
               Width           =   975
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":4E83A
               TabIndex        =   138
               Top             =   1320
               Width           =   1215
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
               Height          =   255
               Left            =   3000
               OleObjectBlob   =   "frmConfSistema.frx":4E8A2
               TabIndex        =   137
               Top             =   720
               Width           =   2055
            End
            Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
               Height          =   255
               Left            =   120
               OleObjectBlob   =   "frmConfSistema.frx":4E916
               TabIndex        =   136
               Top             =   720
               Width           =   1935
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
         Height          =   3495
         Left            =   -74880
         TabIndex        =   8
         Top             =   1560
         Width           =   8895
         Begin VB.CheckBox Check10 
            Caption         =   "Requer conexão segura (SSL)"
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Frame Frame22 
            Caption         =   "Tipo de autenticação"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   115
            Top             =   1680
            Width           =   8655
            Begin VB.OptionButton Option4 
               Caption         =   "NTLM ( autenticação de senha segura no Microsoft® Outlook® Express)"
               Height          =   495
               Left            =   120
               TabIndex        =   118
               Top             =   840
               Width           =   6375
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Básica"
               Height          =   255
               Left            =   120
               TabIndex        =   117
               Top             =   600
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Não requer"
               Height          =   255
               Left            =   120
               TabIndex        =   116
               Top             =   240
               Width           =   3135
            End
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   4920
            PasswordChar    =   "*"
            TabIndex        =   3
            Tag             =   "Senha do usuario de autenticação"
            ToolTipText     =   "Senha do usuario de autenticação"
            Top             =   600
            Width           =   3855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   4920
            OleObjectBlob   =   "frmConfSistema.frx":4E990
            TabIndex        =   114
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Tag             =   "usuario de autenticação"
            ToolTipText     =   "usuario de autenticação"
            Top             =   600
            Width           =   4575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":4E9F4
            TabIndex        =   113
            Top             =   360
            Width           =   1575
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
         TabIndex        =   7
         Top             =   480
         Width           =   5775
         Begin ACTIVESKINLibCtl.SkinLabel Label16 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmConfSistema.frx":4EA5C
            TabIndex        =   41
            Top             =   600
            Width           =   5535
         End
         Begin VB.TextBox txtEmail 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Tag             =   "Endereço do servidor de SMTP"
            ToolTipText     =   "Endereço do servidor de SMTP"
            Top             =   240
            Width           =   5535
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
         Left            =   -74880
         TabIndex        =   5
         Top             =   6360
         Visible         =   0   'False
         Width           =   9135
         Begin IMRM.chameleonButton chameleonButton1 
            Height          =   735
            Left            =   1800
            TabIndex        =   0
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
            MICON           =   "frmConfSistema.frx":4EB06
            PICN            =   "frmConfSistema.frx":4EB22
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
            TabIndex        =   16
            Top             =   240
            Width           =   3735
            Begin VB.CommandButton cmdCad 
               Caption         =   "Importar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   24
               Tag             =   "Importar"
               ToolTipText     =   "Importar"
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmdCad 
               Caption         =   "Localizar..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
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
               TabIndex        =   23
               Tag             =   "Localizar"
               ToolTipText     =   "Localizar"
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text1 
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
               Left            =   120
               TabIndex        =   17
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
            Height          =   255
            Left            =   240
            TabIndex        =   6
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
Public vTab As String
Public vCAT As String
Public X As Integer
Public no As Node
Public vChave As String
Public vChaveTAB As String
Public vIDGrupo As Integer
Public vIDCriterio As Integer
Public vIDSubCriterio As Integer
Private vPonte1 As TextBox, vPonte2 As TextBox

Private Sub chameleonButton1_Click()
On Error Resume Next
    mobjMsg.Abrir "Deseja realmente importar os dados das tabelas selecionadas?", YesNo, pergunta, "IMRM"
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
        mobjMsg.Abrir "Nenhuma tabela selecionada. Marque a tabela a ser importada", Ok, critico, "IMRM"
        Exit Sub
    End If
    
    'A ROTINA ABAIXO VC SELECIONA UM PROCESSO Q ESTA NA MEMORIA P SER REMOVIDO
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    '--------------------------------------------------------------------------
    mobjMsg.Abrir "Dados importados com sucesso. Para vizualisar os dados feche a tabela e abra novamente", Ok, informacao, "IMRM"
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub ImportaDadosCargo()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de CARGOS..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsCargos.AddNew
            rsCargos.Fields(0) = .Range("A" & J).Value 'Código do CARGO
            rsCargos.Fields(1) = .Range("B" & J).Value 'Código do CBO
            rsCargos.Fields(2) = .Range("C" & J).Value 'Nome do CARGO
            rsCargos.Fields(5) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de cargos do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub ImportaDadosHabilidade()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Habilidade..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsHabilidade.AddNew
            rsHabilidade.Fields(0) = .Range("A" & J).Value 'Código da Habilidade
            rsHabilidade.Fields(1) = .Range("B" & J).Value 'Habilidade
            rsHabilidade.Fields(2) = .Range("D" & J).Value 'Peso da Habilidade
            rsHabilidade.Fields(3) = .Range("C" & J).Value 'Descrição da Habilidade
            rsHabilidade.Fields(4) = "S" 'Status
            rsHabilidade.Fields(5) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de Habilidades do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub ImportaDadosAvaliacao()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - AVALIACAO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Avaliação do Treinamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsAvaliacao.AddNew
            rsAvaliacao.Fields(0) = .Range("A" & J).Value 'Código da avaliação
            rsAvaliacao.Fields(1) = .Range("B" & J).Value 'Nome da avaliação
            rsAvaliacao.Fields(2) = .Range("C" & J).Value 'Tipo da avaliação
            rsAvaliacao.Fields(3) = .Range("D" & J).Value 'Peso da avaliação
            rsAvaliacao.Fields(4) = "S" 'Status
            rsAvaliacao.Fields(5) = .Range("E" & J).Value 'Descrição da avaliação
            rsAvaliacao.Fields(6) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de Avaliação do Treinamento do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub ImportaDadosEscolaridade()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - ESCOLARIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Escolaridade..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsEscolaridade.AddNew
            rsEscolaridade.Fields(0) = .Range("A" & J).Value 'Código da escolaridade
            rsEscolaridade.Fields(1) = .Range("B" & J).Value 'Nome da escolaridade
            rsEscolaridade.Fields(2) = .Range("C" & J).Value 'Peso da escolaridade
            rsEscolaridade.Fields(3) = "S" 'Status
            rsEscolaridade.Fields(4) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de Escolaridade do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub


Private Sub ImportaDadosDepartamento()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - DEPARTAMENTO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Departamento..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsDepartamento.AddNew
            rsDepartamento.Fields(0) = .Range("A" & J).Value 'Código do departamento
            rsDepartamento.Fields(1) = .Range("B" & J).Value 'Nome do departamento
            rsDepartamento.Fields(2) = .Range("C" & J).Value 'descrição do departamento
            rsDepartamento.Fields(3) = "S" 'Status
            rsDepartamento.Fields(4) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de Departamentos do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "Atenção"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub ImportaDadosSetor()
'On Error GoTo TrataErro
    Dim J As Integer
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
    J = 2
    For X = 1 To 100000
        With Plan
            If .Range("A" & J).Value = "" Then Exit For
            J = J + 1
        End With
    Next
    Principal.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - SETORES
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Setores..."
    Principal.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To Principal.ProgressBar1.Max
        With Plan
            Principal.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsSetor.AddNew
            rsSetor.Fields(2) = .Range("A" & J).Value 'Código do departamento
            rsSetor.Fields(0) = .Range("B" & J).Value 'Código do setor
            rsSetor.Fields(1) = .Range("C" & J).Value 'Nome do setor
            rsSetor.Fields(3) = .Range("C" & J).Value 'Descrição do setor
            rsSetor.Fields(4) = "S" 'Status
            rsSetor.Fields(5) = vCodcoligada 'Codigo da coligada
            J = J + 1
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
TrataErro:
    mobjMsg.Abrir "Existem dados cadastrados na tabela de Setores do sistema. Para que a importação seja realizada ela deve estar vazia", Ok, critico, "IMRM"
    Legenda = "ERRO na importação de dados"
    Principal.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub chameleonButton4_Click()
    AlteraColigada
    SSTab4.Tab = 0
End Sub

Private Sub chameleonButton5_Click()
    mobjMsg.Abrir "Rotina em desenvolvimento", Ok, critico, "Atenção"
End Sub

Private Sub Check11_Click()
On Error GoTo Err
    Dim cnTestaConexão As ADODB.Connection
    Set cnTestaConexão = New ADODB.Connection
    Set cnTestaConexão = New ADODB.Connection
    If Check11.Value = 1 Then
        Frame26.Enabled = True
        SkinLabel25.Enabled = True
        SkinLabel26.Enabled = True
        SkinLabel27.Enabled = True
        SkinLabel28.Enabled = True
        SkinLabel29.Enabled = True
        SkinLabel30.Enabled = True
        txtIntegra(7).Enabled = True
        txtIntegra(6).Enabled = False
        txtIntegra(5).Enabled = True
        txtIntegra(4).Enabled = True
        Command1.Enabled = True
        Command2.Enabled = True
    Else
        Frame26.Enabled = False
        SkinLabel25.Enabled = False
        SkinLabel26.Enabled = False
        SkinLabel27.Enabled = False
        SkinLabel28.Enabled = False
        SkinLabel29.Enabled = False
        SkinLabel30.Enabled = False
        txtIntegra(7).Enabled = False
        txtIntegra(6).Enabled = False
        txtIntegra(5).Enabled = False
        txtIntegra(4).Enabled = False
        Command1.Enabled = False
        Command2.Enabled = False
    End If
    'SOMENTE TESTA SE O BANCO EXISTE
    'SE EXISTIR DESABILITA O BOTAO DE CRIAR BANCO
    cnTestaConexão.Open "Provider=SQLOLEDB.1;Password=" & vSenhaBancoOffline & ";Persist Security Info=True;User ID=" & vUsuBancoOffline & ";Initial Catalog=" & vBancoOffline & ";Data Source=" & vServerOffline
    Command2.Enabled = False
    cnTestaConexão.Close
    Set cnTestaConexão = Nothing
    Exit Sub

Err:
    If Err.Number = -2147467259 Then
        Command1.Enabled = False
        Command2.Enabled = False
        Resume Next
        Exit Sub
    ElseIf Err.Number = 3704 Then
        Command1.Enabled = False
        Command2.Enabled = False
        Resume Next
        Exit Sub
    End If
    Command2.Enabled = True
    Resume Next
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
        Text2 = "Informe o caminho do executável: AtualizaIMRM.exe"
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
            .Filter = "(AtualizaIMRM *.EXE)|*.exe"
            .ShowOpen
            Caminho3 = .FileName
        End With
        Text2 = Caminho3
    End Select
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        mobjMsg.Abrir "Deseja salvar os dados de parametrização?", YesNo, pergunta, "IMRM"
        If Tp = 1 Then
            GravaParametros
            'gravaLog "Mádia para aprovação: " & txtCadParametro(0), "Gerar introdutório: " & Check3.Value, "Aprovação com restrição: " & txtCadParametro(1)
            Pesquisa = 0
            'Unload Me
        End If
    Case 1
        mobjMsg.Abrir "Deseja sair da tela configurações do sistema?", YesNo, pergunta, "IMRM"
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
        ExcluirItemLV ListView1
    Case 8
        LimpaControles txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
    Case 9
        AlteraLV ListView1, txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
    Case 10
        IncluirLV ListView1, txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
        LimpaControles txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
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
    Case 11
        ExcluirItemLV ListView2
    Case 14
        LimpaControles Text4, Text5, Text6, Text4, Text4, Text4, Text4, Text4, Text4, Text4
    Case 17
        AlteraLV ListView2, Text6, Text4, Text5, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
    Case 18
        IncluirLV ListView2, Text6, Text4, Text5, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
        LimpaControles Text4, Text5, Text6, Text4, Text4, Text4, Text4, Text4, Text4, Text4
    End Select
End Sub

Private Sub IncluiTreeview()
    'If ValidaCampoTree = False Then Exit Sub
On Error GoTo Err
    Dim rsMenu As New ADODB.Recordset
    Dim SqlMenu As String
    
    cnBanco.BeginTrans
    
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
Err:
    mobjMsg.Abrir "Não é permitido duplicação de registros", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
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
End Sub

Private Sub DeletaTreeview()
    Dim rsDeleta As New ADODB.Recordset
    Dim SqlDeleta As String
    Dim llng_Contador As Long
    For llng_Contador = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(llng_Contador).Selected = True Then
            If Msgbox("Confirma Exclusão", vbQuestion + vbYesNo, "IMRM") = vbYes Then
                SqlDeleta = "Delete from tbMenuConf where id = '" & Val(SkinLabel13) & "' and codcoligada= '" & vCodcoligada & "'"
                rsDeleta.Open SqlDeleta, cnBanco, adOpenKeyset, adLockOptimistic
                
'                If Check7.Value = 1 Then
'                    SqlMenu = "Delete from tbConfGrupo where idmenu = '" & Val(Combo1) & "' and idsub = '" & Combo3 & Combo4 & "'"
'                Else
'                    SqlMenu = "Delete from tbConfGrupo where idmenu = '" & Val(Combo1) & "' and idsub = '" & Combo3 & "'"
'                End If
            
            End If
        End If
    Next
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
        .hWndOwner = Me.hwnd
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

Private Sub Command1_Click()
    mobjMsg.Abrir "Esse processo pode demorar vários minutos. Deseja continuar?", YesNo, pergunta, "IMRM"
    If Tp = 1 Then
        ValidaIDMOV "tbEmprestimo"
        ValidaIDMOV "tbDevolucao"
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbConfEmail", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbConfiguracoes", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbConfLV", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbDadosBanco", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbDadosEmpresa", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbDevolucao", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbDevolucaoItens", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbEmprestimo", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbEmprestimoItens", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbGrupo", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbintegracao", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbintegracaooffline", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbLog", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbMenu", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbMenuConf", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbMov", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbparametros", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbSenha", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbUsuarios", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbSincronizacao", ""
        ClonarDados vServerOffline, txtIntegra(6).Text, vUsuBancoOffline, vSenhaBancoOffline, "tbServerSincronizaDados", ""
   
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "GAUTOINC", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TMOV", "where CODTMV in('2.2.15','1.2.16')"
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TITMMOV", ",TMOV where TITMMOV.IDMOV = TMOV.IDMOV and TMOV.CODTMV in('2.2.15','1.2.16')"
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRDLOC", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TMOVRELAC", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TVEN", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TVENCOMPL", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "PFUNCAO", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "PSECAO", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRODUTO", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRODUTODEF", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "OFVENCPLANOMANUT", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "OFPLANOMANUT", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TLOC", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "GCOLIGADA", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "PFUNC", ""
        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TITMMOVRELAC", ""
        
        'Controle de Inserção de registros referente às tabelas do RM
        ControlaRegTabs "PFUNC", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TVEN", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TVENCOMPL", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TPRODUTO", "-", "DTCADASTRAMENTO", "", "CORPORE"
        ControlaRegTabs "TPRDLOC", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TPRODUTODEF", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "OFPLANOMANUT", "OFVENCPLANOMANUT", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TLOC", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TBUSUARIOS", "-", "CODIGO", "", "FERRAMENTARIA"
        ControlaRegTabs "TBSENHA", "-", "CODIGO", "", "FERRAMENTARIA"
        ControlaRegTabs "PFUNCAO", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "PSECAO", "-", "RECCREATEDON", "", "CORPORE"
        ControlaRegTabs "TBLOCALESTOQUE", "-", "ID", "", "FERRAMENTARIA"
        ControlaRegTabs "OFVENCPLANOMANUT", "-", "DATAATUALIZACAO", "", "CORPORE"
        
        mobjMsg.Abrir "Tabelas importadas com sucesso", Ok, informacao, "Ferramentaria"
    End If
    'mobjMsg.Abrir "Deseja importar dados de antigos empréstimos?", YesNo, pergunta, "Ferramentaria"
    'If Tp = 1 Then
    '    ConexaoSAP
    '    ImportarDadosSistemaAntigo
    '    cnBancoSAP.Close
    '    Set cnBancoSAP = Nothing
    'End If
End Sub

Private Sub Command2_Click()
    If TestaConexaoOffLine = True Then
        mobjMsg.Abrir "Bancos e Tabelas OFFLINE criadas com sucesso", , critico
    End If
End Sub

Private Sub Command3_Click()
        vServerOffline = txtIntegra(13).Text 'Servidor e porta
        vBancoOffline = txtIntegra(12).Text 'Banco
        vUsuBancoOffline = txtIntegra(11).Text 'Usuário
        vSenhaBancoOffline = txtIntegra(10).Text 'Senha

'        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "GAUTOINC", ""
'        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TMOV", "where CODTMV in('2.2.15','1.2.16')"
'        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TITMMOV", ",TMOV where TITMMOV.IDMOV = TMOV.IDMOV and TMOV.CODTMV in('2.2.15','1.2.16')"
'        ClonarDados vServerOffline, txtIntegra(8).Text, vUsuBancoOffline, vSenhaBancoOffline, "TMOVRELAC", ""

    
    If Check12.Value = 1 Then
        'Funcionários
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "PFUNC", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TVEN", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TVENCOMPL", "", "CORPORERM_OFF", "RECCREATEDON"
        Picture1.Visible = True
    End If
    If Check13.Value = 1 Then
        'Usuários
        ClonarDadosSvRemoto vServerOffline, txtIntegra(9).Text, vUsuBancoOffline, vSenhaBancoOffline, "TBUSUARIOS", "", "FERRAMENTARIA_OFF", "CODIGO"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(9).Text, vUsuBancoOffline, vSenhaBancoOffline, "TBSENHA", "", "FERRAMENTARIA_OFF", "CODIGO"
    End If
    If Check14.Value = 1 Then
        'Produto (1ª parte)
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRODUTO", "", "CORPORERM_OFF", "DTCADASTRAMENTO"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRDLOC", "", "CORPORERM_OFF", "RECCREATEDON"
'        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TPRODUTODEF", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "OFVENCPLANOMANUT", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "OFPLANOMANUT", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "TLOC", "", "CORPORERM_OFF", "RECCREATEDON"
        'Produto (2ª parte)
        'SincronizarDadosImportar txtIntegra(13).Text, txtIntegra(12).Text, txtIntegra(11).Text, txtIntegra(10).Text, "TPRDLOC", "where codcoligada = '" & vCodcoligada & "' and codloc = '" & vLocalEstoque & "'", "ORDER BY IDPRD"
    End If
    If Check15.Value = 1 Then
        '1ª Passo - Buscar IDMOV no servidor GNV (SERVIDOR REMOTO) e Atualizar no CORPORERM_OFF (SERVIDOR LOCAL)
        'e tambem na FERRAMENTARIA_OFF (SERVIDOR LOCAL)
        mobjMsg.Abrir "A velocidade da sincronização irá depender das conexões de internet", Ok, informacao, "Atenção"
        AtualizaGAutoIncIDMOV vServerOffline, vBancoOffline, vUsuBancoOffline, vSenhaBancoOffline
        
        '2º Passo - Sincronizar IDMOV (Identificadores de movimentos das tabelas:)
        'FERRAMENTARIA_OFF: tbEmprestimo, tbEmprestimoItens, tbDevolucao, tbDevolucaoItens
        'CORPORERM_OFF: GAUTOINC, TMOV, TITMMOV, TMOVRELAC
        
        'EM DESENVOLVIMENTO
        'TOTVS
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(12).Text, txtIntegra(11).Text, txtIntegra(10).Text, "TMOV", "where idmov = ", "", "TOTVS"
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(12).Text, txtIntegra(11).Text, txtIntegra(10).Text, "TITMMOV", "where idmov = ", "", "TOTVS"
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(12).Text, txtIntegra(11).Text, txtIntegra(10).Text, "TMOVRELAC", "where idmovorigem = ", "", "TOTVS"
        'FERRAMENTARIA
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(9).Text, txtIntegra(11).Text, txtIntegra(10).Text, "tbEmprestimo", "where idmov = ", "", "FERRAMENTARIA"
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(9).Text, txtIntegra(11).Text, txtIntegra(10).Text, "tbEmprestimoItens", "where idmov = ", "", "FERRAMENTARIA"
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(9).Text, txtIntegra(11).Text, txtIntegra(10).Text, "tbDevolucao", "where idmov = ", "", "FERRAMENTARIA"
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(9).Text, txtIntegra(11).Text, txtIntegra(10).Text, "tbDevolucaoItens", "where idmov = ", "", "FERRAMENTARIA"
        'EM DESENVOLVIMENTO
        SincronizarDadosExportar txtIntegra(13).Text, txtIntegra(9).Text, txtIntegra(11).Text, txtIntegra(10).Text, "tbSincronizacao", "where idmovsincronizado = ", "", "FERRAMENTARIA"
        ValidaIDMOV "tbEmprestimo"
        ValidaIDMOV "tbDevolucao"
        
    End If
    If Check16.Value = 1 Then
        'Funções/Seções
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "PFUNCAO", "", "CORPORERM_OFF", "RECCREATEDON"
        ClonarDadosSvRemoto vServerOffline, txtIntegra(12).Text, vUsuBancoOffline, vSenhaBancoOffline, "PSECAO", "", "CORPORERM_OFF", "RECCREATEDON"
        Picture7.Visible = True
    End If
    mobjMsg.Abrir "Sincronização realizada com sucesso", Ok, informacao, "Atenção"
End Sub

Private Sub Command4_Click()
    enviaEmail
End Sub

Private Sub Command5_Click()
    ControlaRegTabs "PFUNC", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TVEN", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TVENCOMPL", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TPRODUTO", "-", "DTCADASTRAMENTO", "", "CORPORE"
    ControlaRegTabs "TPRDLOC", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TPRODUTODEF", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "OFPLANOMANUT", "OFVENCPLANOMANUT", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TLOC", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TBUSUARIOS", "-", "CODIGO", "", "FERRAMENTARIA"
    ControlaRegTabs "TBSENHA", "-", "CODIGO", "", "FERRAMENTARIA"
    ControlaRegTabs "PFUNCAO", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "PSECAO", "-", "RECCREATEDON", "", "CORPORE"
    ControlaRegTabs "TBLOCALESTOQUE", "-", "ID", "", "FERRAMENTARIA"
    ControlaRegTabs "OFVENCPLANOMANUT", "-", "DATAATUALIZACAO", "", "CORPORE"
    Msgbox "Dados criados com sucesso!"
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
    Check11_Click
CompoeTAB

'    LimpaControlesAvaliacao
'    LimpaControlesABS
    'LimpaControlesColigada
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Email", ListView1.Width / 1.1
'    ListView1.ColumnHeaders.Add , , "Avaliar após", ListView1.Width / 4
'    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 3
'    ListView1.ColumnHeaders.Add , , "Modelo Ativo", ListView1.Width / 6
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Classificação", ListView2.Width / 2
    ListView2.ColumnHeaders.Add , , "De", ListView2.Width / 5
    ListView2.ColumnHeaders.Add , , "Até", ListView2.Width / 5
    
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
    ListView3.ColumnHeaders.Add , , "serie", ListView3.Width / 10000
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
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
    Dim rsEmailAvRec As New ADODB.Recordset
    Dim sqlEmailAvRec  As String
    Dim rsClassificacao As New ADODB.Recordset
    Dim sqlClassificacao  As String
    
    Dim rsABS As New ADODB.Recordset
    Dim sqlABS As String
    Dim rsColigadas As New ADODB.Recordset
    Dim sqlColigadas As String
    
    Dim ItemLst As ListItem
    Dim X As Integer
    
    ' Compoe Listview1
'    sqlEmailAvRec = "Select * from tbEmailAvRec"
'    rsEmailAvRec.Open sqlEmailAvRec, cnBanco, adOpenKeyset, adLockReadOnly
'    X = 0
'    While Not rsEmailAvRec.EOF
'        Set ItemLst = ListView1.ListItems.Add(, , rsEmailAvRec.Fields(0))
'        rsEmailAvRec.MoveNext
'        X = X + 1
'    Wend
'    Me.ListView1.Sorted = True
'    Me.ListView1.SortKey = 0
'    Me.ListView1.SortOrder = lvwDescending
'    rsEmailAvRec.Close
'    Set rsEmailAvRec = Nothing
    
    ' Compoe Listview2
'    sqlClassificacao = "Select * from tbClassificacao Order by idclassificacao"
'    rsClassificacao.Open sqlClassificacao, cnBanco, adOpenKeyset, adLockOptimistic
'    X = 0
'    While Not rsClassificacao.EOF
'        Set ItemLst = ListView2.ListItems.Add(, , Format(rsClassificacao.Fields(0), "00"))
'        ItemLst.SubItems(1) = "" & rsClassificacao.Fields(1)
'        ItemLst.SubItems(2) = "" & rsClassificacao.Fields(2)
'        rsClassificacao.MoveNext
'        X = X + 1
'    Wend
'    Me.ListView2.Sorted = True
'    Me.ListView2.SortKey = 0
'    Me.ListView2.SortOrder = lvwAscending
'    rsClassificacao.Close
'    Set rsClassificacao = Nothing

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
        ItemLst.SubItems(15) = "" & rsColigadas.Fields(15)
        rsColigadas.MoveNext
        X = X + 1
    Wend
    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 0
    Me.ListView3.SortOrder = lvwDescending
    rsColigadas.Close
    Set rsColigadas = Nothing
End Sub

Private Function GeraCodigo(LV As ListView)
    Dim X As Integer
    X = 1
    LV.SortOrder = lvwDescending
    LV.ListItems.Item(X).Selected = True
    GeraCodigo = LV.ListItems.Item(X) + 1
    LV.SortOrder = lvwAscending
    Exit Function
End Function

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
                ListView3.SelectedItem.ListSubItems.Item(15) = txtDadosEmpresa(12)
                
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
    ItemLst.SubItems(15) = txtDadosEmpresa(12)
    
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
    Me.txtDadosEmpresa(12).Text = ListView3.SelectedItem.ListSubItems.Item(15) 'Numero de serie
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
End Sub

Private Sub CarregaParametros()
    Dim rsParametros As New ADODB.Recordset
    Dim sqlParametros As String
    Dim rsEmpresa As New ADODB.Recordset
    Dim sqlEmpresa As String
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String
    Dim rsIntegracao As New ADODB.Recordset
    Dim sqlIntegracao As String
    Dim rsSincroniza As New ADODB.Recordset
    Dim sqlSincroniza As String
    
    If Text1.Text = "" Then cmdCad(1).Enabled = False
    
    'sqlSincroniza = "Select * from tbServerSincronizaDados where codcoligada = '" & vCodcoligada & "'"
    'rsSincroniza.Open sqlSincroniza, cnBanco, adOpenKeyset, adLockOptimistic
    'If rsSincroniza.RecordCount > 0 Then
    '    txtIntegra(13).Text = rsSincroniza.Fields(0)
    '    txtIntegra(12).Text = rsSincroniza.Fields(1)
    '    txtIntegra(9).Text = rsSincroniza.Fields(2)
    '    txtIntegra(11).Text = rsSincroniza.Fields(3)
    '    txtIntegra(10).Text = rsSincroniza.Fields(4)
    'End If
    
    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsParametros.EOF Then
        txtCadParametro(0) = rsParametros.Fields(0) 'Media Provação
        If Not IsNull(rsParametros.Fields(2)) Then
            Combo5.Text = rsParametros.Fields(2) 'Aprovado com restrição
        End If
        
        If Not IsNull(rsParametros.Fields(14)) Then DTPicker1.Value = rsParametros.Fields(14)
        
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
        txtDadosEmpresa(12) = rsEmpresa.Fields(15)
    
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
        If Not IsNull(rsConfEmail.Fields(4)) Then txtEmail(3) = rsConfEmail.Fields(4)
        If Not IsNull(rsConfEmail.Fields(5)) Then
            If rsConfEmail.Fields(5) = 0 Then
                Check10.Value = 0
            Else
                Check10.Value = 1
            End If
        End If
        If Not IsNull(rsConfEmail.Fields(6)) Then
            If rsConfEmail.Fields(6) = 0 Then Option2.Value = True
            If rsConfEmail.Fields(6) = 1 Then Option3.Value = True
            If rsConfEmail.Fields(6) = 2 Then Option4.Value = True
        End If
    
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
    rsIntegracao.Close
    Set rsIntegracao = Nothing
'*********************
    'sqlIntegracao = "Select * from tbIntegracaooffline"
    'rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockReadOnly
    'If Not rsIntegracao.EOF Then
    '    If rsIntegracao.Fields(0) = "S" Then Check11.Value = 1 Else Check11.Value = 0
    '    txtIntegra(7).Text = rsIntegracao.Fields(1)
    '    txtIntegra(6).Text = rsIntegracao.Fields(2)
    '    txtIntegra(5).Text = rsIntegracao.Fields(3)
    '    txtIntegra(4).Text = rsIntegracao.Fields(4)
    '    vServerOffline = txtIntegra(7).Text
    '    vBancoOffline = txtIntegra(6).Text
    '    vUsuBancoOffline = txtIntegra(5).Text
    '    vSenhaBancoOffline = txtIntegra(4).Text
    'End If
'*********************
    
    rsConfEmail.Close
    Set rsConfEmail = Nothing
    rsEmpresa.Close
    Set rsEmpresa = Nothing
'    rsIntegracao.Close
'    Set rsIntegracao = Nothing
    Exit Sub
TrataErro1:
    Label59.Visible = True
    Resume Next
End Sub

Private Sub GravaParametros()
'On Error Resume Next
    If ValidaCampo = False Then Exit Sub
    If Check6.Value = 1 And Text2 = "" Then
        mobjMsg.Abrir "Informe o caminho do executável: AtualizaIMRMH.exe", Ok, critico, "Atenção"
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
    
    Dim rsEmailAvRec As New ADODB.Recordset
    Dim sqlEmailAvRec As String
    
    Dim rsClassificacao As New ADODB.Recordset
    Dim sqlClassificacao As String
    Dim rsSincroniza As New ADODB.Recordset
    Dim sqlSincroniza As String
    
    cnBanco.BeginTrans

'    sqlSincroniza = "Delete from tbServerSincronizaDados where codcoligada = '" & vCodcoligada & "'"
'    rsSincroniza.Open sqlSincroniza, cnBanco
    
'    sqlSincroniza = "Select * from tbServerSincronizaDados where codcoligada = '" & vCodcoligada & "'"
'    rsSincroniza.Open sqlSincroniza, cnBanco, adOpenKeyset, adLockOptimistic
'    If txtIntegra(13).Text <> "" Then
'        rsSincroniza.AddNew
'        rsSincroniza.Fields(0) = txtIntegra(13).Text 'Nome/endereço do servidor
'        rsSincroniza.Fields(1) = txtIntegra(12).Text 'Nome do banco Totvs RM
'        rsSincroniza.Fields(2) = txtIntegra(9).Text 'Nome do banco da Ferramentaria
'        rsSincroniza.Fields(3) = txtIntegra(11).Text 'Nome do usuário do banco
'        rsSincroniza.Fields(4) = txtIntegra(10).Text 'Senha do usuário do banco
'        rsSincroniza.Fields(5) = vCodcoligada 'Código da coligada
'        rsSincroniza.Update
'    End If
'    rsSincroniza.Close
'    Set rsSincroniza = Nothing
    
    
    sqlDeletar = "Delete from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco

    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockOptimistic
    rsParametros.AddNew
    rsParametros.Fields(0) = 50
    rsParametros.Fields(2) = Combo5.Text
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
        GeraLog = "S"
    Else
        rsParametros.Fields(3) = "N"
        GeraLog = "N"
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
    If DTPicker1.Value <> "" Then
        rsParametros.Fields(14) = DTPicker1.Value
        vInicioAvOC = DTPicker1.Value
    End If
    rsParametros.Fields(6) = vCodcoligada 'Codigo da coligada
    If Check4.Value = 1 Then
'*************************
        'GRAVA DADOS DE INTEGRAÇÃO
        
        vServerSAP = txtIntegra(0).Text
        vBancoSAP = txtIntegra(1).Text
        vUsuBancoTovs = txtIntegra(2).Text
        vSenhaBancoSAP = txtIntegra(3).Text
    
        If testaParametros = False Then
            Check4.Value = 0
            mobjMsg.Abrir "Os dados informados para conexão não estão corretos", Ok, critico, "Conexão TOTVS CORPORERM"
            rsParametros.Fields(5) = "N"
        Else
            rsParametros.Fields(5) = "S"
            vIntegra = "S"
            
            mobjMsg.Abrir "Deseja importar dados de antigos empréstimos?", YesNo, pergunta, "IMRM"
            If Tp = 1 Then
                ConexaoSAP
                ImportarDadosSistemaAntigo
                ValidaIDMOV "tbEmprestimo"
                ValidaIDMOV "tbDevolucao"
                
                cnBancoSAP.Close
                Set cnBancoSAP = Nothing
            End If
            
            
        End If
            
        sqlIntegracao = "Select * from tbIntegracao Where codcoligada = '" & vCodcoligada & "'"
        rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockOptimistic
        If rsIntegracao.EOF Then rsIntegracao.AddNew
        ' 1 = SQL Server / 2 = Oracle
        If optIntegra(0).Value = True Then rsIntegracao.Fields(0) = 1 Else rsIntegracao.Fields(0) = 2
        If chkIntegra(0).Value = True Then rsIntegracao.Fields(1) = 1
        If chkIntegra(1).Value = True Then rsIntegracao.Fields(1) = 2
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
        
'*************************
        'GRAVA DADOS DE INTEGRAÇÃO OFFLINE
        If Check11.Value = 1 Then
            vServerOffline = txtIntegra(7).Text
            vBancoOffline = txtIntegra(6).Text
            vUsuBancoOffline = txtIntegra(5).Text
            vSenhaBancoOffline = txtIntegra(4).Text
        
            sqlIntegracao = "Select * from tbintegracaooffline"
            rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockOptimistic
            If rsIntegracao.EOF Then rsIntegracao.AddNew
            If Check11.Value = 1 Then
                rsIntegracao.Fields(0) = "S"
                vIntegraOffline = "S"
            Else
                rsIntegracao.Fields(0) = "N"
                vIntegraOffline = "N"
            End If
            rsIntegracao.Fields(1) = txtIntegra(7).Text
            rsIntegracao.Fields(2) = txtIntegra(6).Text
            rsIntegracao.Fields(3) = txtIntegra(5).Text
            rsIntegracao.Fields(4) = txtIntegra(4).Text
            rsIntegracao.Update
            rsIntegracao.Close
            Set rsIntegracao = Nothing
            
            If TestaConexaoOffLine = False Then
                vIntegraOffline = "N"
            Else
                vIntegraOffline = "S"
            End If
        Else
            vIntegraOffline = "N"
        End If
    Else
        rsParametros.Fields(5) = "N"
        vIntegra = "N"
    End If
'*************************
    rsParametros.Update
    rsParametros.Close
    MediaGlobal = 50
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
        rsEmpresa.Fields(15) = ListView3.SelectedItem.ListSubItems.Item(15) 'Serie
    Next
    
    sqlDeletar = "Delete from tbConfEmail where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    If Check10.Value = 1 Then 'SSL
        vPonte1 = 1
        vSSL = True
    Else
        vPonte1 = 0
        vSSL = False
    End If
    
    
    If Option2.Value = True Then
        vPonte2 = 0
        vSMTPAutentic = 0
    End If
    If Option3.Value = True Then
        vPonte2 = 1
        vSMTPAutentic = 1
    End If
    If Option4.Value = True Then
        vPonte2 = 2
        vSMTPAutentic = 2
    End If
    
'    sqlConfEmail = "Insert into tbConfEmail(smtp,usuario,senha,codcoligada,porta,ssl,smtpautentic) Values('" & txtEmail(0) & "','" & txtEmail(1) & "','" & txtEmail(2) & "','" & vCodcoligada & "','" & txtEmail(3) & "','" & Val(vPonte1) & "','" & Val(vPonte2) & "')"
'    rsConfEmail.Open sqlConfEmail, cnBanco
    
'    vSMTP = txtEmail(0)
'    vUsuEmail = txtEmail(1)
'    vSenhaEmail = txtEmail(2)
'    vPorta = txtEmail(3)
    
    rsEmpresa.Update
    rsEmpresa.Close
    Set rsEmpresa = Nothing
    
    
'-------------------------------
'    sqlEmailAvRec = "Delete from tbEmailAvRec"
'    rsEmailAvRec.Open sqlEmailAvRec, cnBanco
'
'    sqlEmailAvRec = "Select * from tbEmailAvRec"
'    rsEmailAvRec.Open sqlEmailAvRec, cnBanco, adOpenKeyset, adLockOptimistic
'    For X = 1 To ListView1.ListItems.Count
'        ListView1.ListItems.Item(X).Selected = True
'        rsEmailAvRec.AddNew
'        rsEmailAvRec.Fields(0) = ListView1.ListItems.Item(X) ' E-mail
'        If sEmailAvRec = "" Then
'            sEmailAvRec = ListView1.ListItems.Item(X)
'        Else
'            sEmailAvRec = sEmailAvRec & ";" & ListView1.ListItems.Item(X)
'        End If
'    Next
'    rsEmailAvRec.Update
'    rsEmailAvRec.Close
'    Set rsEmailAvRec = Nothing
'------------------------------
    
'-------------------------------
'    sqlClassificacao = "Delete from tbClassificacao"
'    rsClassificacao.Open sqlClassificacao, cnBanco
'
'    sqlClassificacao = "Select * from tbClassificacao"
'    rsClassificacao.Open sqlClassificacao, cnBanco, adOpenKeyset, adLockOptimistic
'    For X = 1 To ListView2.ListItems.Count
'        ListView2.ListItems.Item(X).Selected = True
'        rsClassificacao.AddNew
'        rsClassificacao.Fields(0) = ListView2.ListItems.Item(X) ' E-mail
'        rsClassificacao.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1) '
'        rsClassificacao.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(2) '
'    Next
'    rsClassificacao.Update
'    rsClassificacao.Close
'    Set rsClassificacao = Nothing
'------------------------------
    
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sLogoEmpresa", Label53 'Logo da empresa
    
    If Check11.Value = 1 Then
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\" & "sServerName", vServerOffline 'Chave com o nome do Servidor OFFLINE
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\" & "sDatabaseName", vBancoOffline 'Chave com o nome do Banco DA IMRM OFFLINE
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\" & "sUsuName", vUsuBancoOffline '
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\" & "sSenhaDB", vSenhaBancoOffline '
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\OFFLINE\" & "sDatabaseNameCORPORERM_OFF", txtIntegra(8).Text 'Chave com o nome do Banco CORPORERM_OFF
        Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\IMRM\" & "sOFFLINE", "S" '
    End If
    
    Set Reg = Nothing
    
    cnBanco.CommitTrans
    mobjMsg.Abrir "Os dados de configuração do sistema foram salvos com sucesso", Ok, informacao, "IMRM"
    Check11_Click
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Function testaParametros()
On Error GoTo Err
    testaParametros = False
    ConexaoSAP
    If vIntegra = "S" Then testaParametros = True Else testaParametros = False
    cnBancoSAP.Close
    Set cnBancoSAP = Nothing
    Exit Function
Err:
    testaParametros = False
End Function

'ABAIXO CONEXÃO COM O BANCO DE DADOS RM OFFLINE
Private Function TestaConexaoOffLine()
    Exit Function 'função não será utilizada no sistema
End Function

Private Function ValidaCampo()
    ValidaCampo = False
    If ListView3.ListItems.Count = 0 Then
        mobjMsg.Abrir "Nenhuma empresa/coligada cadastrada. Favor informar os dados da empresa/coligada", Ok, informacao, "IMRM"
        SSTab1.Tab = 2
        Exit Function
    End If
    
    If Check11.Value = 1 Then
        If txtIntegra(7) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtIntegra(7).Tag, Ok, informacao, "Atenção"
            Me.txtIntegra(7).SetFocus
            Exit Function
        End If
        If txtIntegra(6) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtIntegra(6).Tag, Ok, informacao, "Atenção"
            Me.txtIntegra(6).SetFocus
            Exit Function
        End If
        If txtIntegra(5) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtIntegra(5).Tag, Ok, informacao, "Atenção"
            Me.txtIntegra(5).SetFocus
            Exit Function
        End If
        If txtIntegra(4) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtIntegra(4).Tag, Ok, informacao, "Atenção"
            Me.txtIntegra(4).SetFocus
            Exit Function
        End If
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
        If txtDadosEmpresa(12) = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtDadosEmpresa(12).Tag, Ok, informacao, "Atenção"
            Me.txtDadosEmpresa(12).SetFocus
            Exit Function
        End If
        If Len(txtDadosEmpresa(12)) < 3 Then
            mobjMsg.Abrir "O campo SERIE deve conter 3 caracteres", Ok, informacao, "Atenção"
            Me.txtDadosEmpresa(12).SetFocus
            Exit Function
        End If
    End If
    ValidaDadosColigada = True
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
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

Private Sub ListView1_DblClick()
    AlteraLV ListView1, txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1), txtCadParametro(1)
End Sub

Private Sub ListView2_DblClick()
    AlteraLV ListView2, Text6, Text4, Text5, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6, Text6
End Sub

Private Sub ListView3_DblClick()
    AlteraColigada
    SSTab4.Tab = 0
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
        Var = Split(Linhas(i), ";")
        For X = 0 To 17
            colheDados(X) = Var(X)
        Next
        If ValidaDados = False Then
            mobjMsg.Abrir "Erro na linha: " & i + 1, Ok, critico, "Atenção"
            Exit Sub
        End If
        insertDados
    Next
    mobjMsg.Abrir "Dados importados com sucesso!", Ok, informacao, "IMRM"
End Sub

Private Function ValidaDados()
    ValidaDados = False
    Dim Y As Integer
    For Y = 0 To 3
        If colheDados(Y) = "" Then
            mobjMsg.Abrir "Erro de consistência na fonte de dados", Ok, informacao, "IMRM"
            Exit Function
        End If
    Next
    ValidaDados = True
End Function

Private Sub insertDados()
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
End Function

Private Sub ValidaIDMOV(vTab As String)
    'Rotina para validar (manter informado) para saber que os dados clonados não
    'precisam ser sincronizados posteriormente (tabelas de EMPRESTIMO e DEVOLUCAO)
On Error Resume Next
    Dim rsValidaIDMOV As New ADODB.Recordset
    Dim SqlValidaIDMOV As String
    Dim rsSincronizacao As New ADODB.Recordset
    Dim SqlSincronizacao As String
    SqlValidaIDMOV = "Select idmov from " & vTab & " as a left join tbSincronizacao as b on a.idmov = b.idmovsincronizado where b.idmovsincronizado is null"
    rsValidaIDMOV.Open SqlValidaIDMOV, cnBanco, adOpenKeyset, adLockReadOnly
    If rsValidaIDMOV.RecordCount > 0 Then
        SqlSincronizacao = "Select * from tbsincronizacao where idmovsincronizado = 1"
        rsSincronizacao.Open SqlSincronizacao, cnBanco, adOpenKeyset, adLockOptimistic
        While Not rsValidaIDMOV.EOF
            rsSincronizacao.AddNew
            rsSincronizacao.Fields(0) = rsValidaIDMOV.Fields(0)
            rsValidaIDMOV.MoveNext
        Wend
        rsSincronizacao.Update
        rsSincronizacao.Close
        Set rsSincronizacao = Nothing
        
        rsValidaIDMOV.Close
        Set rsValidaIDMOV = Nothing
    End If
End Sub

Private Function AtualizaGAutoIncIDMOV(vServerRemoto As String, vBancoRemoto As String, vUsuarioRemoto As String, vSenhaRemoto As String)
    Dim rsAtualizaGAutoIncIDMOV As New ADODB.Recordset
    Dim SqlAtualizaGAutoIncIDMOV As String
    Dim vGAutoIncRemoto As Double, vGAutoIncLocal As Double, vDifGAutoInc As Double
    Dim vGuardaMaiorIDMOV As Double, vUltimoIDSinc As Double
    
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=SQLOLEDB.1;Password=" & vSenhaRemoto & ";Persist Security Info=True;User ID=" & vUsuarioRemoto & ";Initial Catalog=" & vBancoRemoto & ";Data Source=" & vServerRemoto
'Pega valor do ultimo IDMOV na tabela GAUTOINC remota
    SqlAtualizaGAutoIncIDMOV = "select * from GAUTOINC as a where a.codautoinc like 'IDMOV' and a.codcoligada = '" & vCodColigadaRM & "'"
    rsAtualizaGAutoIncIDMOV.Open SqlAtualizaGAutoIncIDMOV, oConn, adOpenKeyset, adLockReadOnly
    If rsAtualizaGAutoIncIDMOV.RecordCount > 0 Then
        vGAutoIncRemoto = Val(rsAtualizaGAutoIncIDMOV.Fields(3)) + 1
    End If
    rsAtualizaGAutoIncIDMOV.Close
    Set rsAtualizaGAutoIncIDMOV = Nothing

'Pega valor do ultimo IDMOV na tabela GAUTOINC local
    SqlAtualizaGAutoIncIDMOV = "select * from GAUTOINC as a where a.codautoinc like 'IDMOV' and a.codcoligada = '" & vCodColigadaRM & "'"
    rsAtualizaGAutoIncIDMOV.Open SqlAtualizaGAutoIncIDMOV, cnBancoSAP, adOpenKeyset, adLockReadOnly
    If rsAtualizaGAutoIncIDMOV.RecordCount > 0 Then
        vGAutoIncLocal = Val(rsAtualizaGAutoIncIDMOV.Fields(3)) + 1
    End If
    rsAtualizaGAutoIncIDMOV.Close
    Set rsAtualizaGAutoIncIDMOV = Nothing
    
'Verifica qual dos 2 GAutoInc é maior
    If vGAutoIncRemoto > vGAutoIncLocal Then
        vGuardaMaiorIDMOV = vGAutoIncRemoto
    Else
        vGuardaMaiorIDMOV = vGAutoIncLocal
    End If
    
    
    
'Pega o ultimo valor sincronizado da tabela tbSincronizacao
    SqlAtualizaGAutoIncIDMOV = "select max(idmovsincronizado) from tbSincronizacao"
    rsAtualizaGAutoIncIDMOV.Open SqlAtualizaGAutoIncIDMOV, cnBanco, adOpenKeyset, adLockReadOnly

'    SqlAtualizaGAutoIncIDMOV = "SELECT MAX(idmov) AS idmov FROM (SELECT idmov FROM tbEmprestimo where codcoligada = '" & vCodColigadaRM & "' UNION ALL SELECT idmov FROM tbDevolucao where codcoligada = '" & vCodColigadaRM & "') as idmov"
'    rsAtualizaGAutoIncIDMOV.Open SqlAtualizaGAutoIncIDMOV, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaGAutoIncIDMOV.RecordCount > 0 Then
        vUltimoIDSinc = Val(rsAtualizaGAutoIncIDMOV.Fields(0))
    End If
    rsAtualizaGAutoIncIDMOV.Close
    Set rsAtualizaGAutoIncIDMOV = Nothing

'Diferença entre o IDMOV remoto e o primeiro IDMOV Local sem sincronização
    'If vGAutoIncRemoto > vGAutoIncLocal Then
        vDifGAutoInc = vGuardaMaiorIDMOV - vUltimoIDSinc
    'Else
    '    vDifGAutoInc = vGAutoIncLocal - vGAutoIncRemoto
    'End If
'Chama a rotina que irá percorrer todas as tabelas de FERRAMENTARIA_OFF e CORPORERM_OFF atualizando os dados
'dos campos ondem ficam gravados os IDMOV's
    vMaiorIDMOV = vGuardaMaiorIDMOV + vDifGAutoInc
    SincronizaIDMov vDifGAutoInc

'    SqlAtualizaGAutoIncIDMOV = "UPDATE GAUTOINC set VALAUTOINC = " & vIDMov & " where codautoinc like 'IDMOV' and codcoligada = '" & vCodColigadaRM & "'"
'    rsAtualizaGAutoIncIDMOV.Open SqlAtualizaGAutoIncIDMOV, cnBancoSAP
'    Set rsAtualizaGAutoIncIDMOV = Nothing

End Function

Private Function SincronizaIDMov(vDiferenca As Double)
    Dim rsAtualizaGeral As New ADODB.Recordset
    Dim SqlAtualizaGeral As String
    
    Dim rsAtualizaEmprestimo As New ADODB.Recordset
    Dim SqlAtualizaEmprestimo As String
    Dim rsAtualizaEmprestimoItens As New ADODB.Recordset
    Dim SqlAtualizaEmprestimoItens As String
   
    Dim rsAtualizaDevolucao As New ADODB.Recordset
    Dim SqlAtualizaDevolucao As String
    Dim rsAtualizaDevolucaoItens As New ADODB.Recordset
    Dim SqlAtualizaDevolucaoItens As String
   
    Dim rsAtualizaTMov As New ADODB.Recordset
    Dim SqlAtualizaTMov As String
    Dim rsAtualizaTitMMov As New ADODB.Recordset
    Dim SqlAtualizaTitMMov As String
    Dim vGuardaMaiorIDMOV As Double
   
    vGuardaMaiorIDMOV = 0
    'ATUALIZA a tabela tbEmprestimo
    SqlAtualizaEmprestimo = "SELECT A.idmov FROM tbEmprestimo AS A LEFT JOIN tbSincronizacao AS B ON A.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NULL AND A.codcoligada = '" & vCodColigadaRM & "' ORDER BY A.idmov"
    rsAtualizaEmprestimo.Open SqlAtualizaEmprestimo, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaEmprestimo.RecordCount > 0 Then
        While Not rsAtualizaEmprestimo.EOF
            SqlAtualizaGeral = "Update tbEmprestimo set idmov = '" & rsAtualizaEmprestimo.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaEmprestimo.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBanco
            Set rsAtualizaGeral = Nothing
            'rsAtualizaEmprestimo.Fields(0) = rsAtualizaEmprestimo.Fields(0) + vDiferenca
            If rsAtualizaEmprestimo.Fields(0) + vDiferenca > vGuardaMaiorIDMOV Then vGuardaMaiorIDMOV = rsAtualizaEmprestimo.Fields(0) + vDiferenca
            rsAtualizaEmprestimo.MoveNext
        Wend
        'rsAtualizaEmprestimo.Update
    End If
    rsAtualizaEmprestimo.Close
    Set rsAtualizaEmprestimo = Nothing
   
    'ATUALIZA a tabela tbEmprestimoItens
    SqlAtualizaEmprestimoItens = "SELECT A.idmov FROM tbEmprestimoItens AS A LEFT JOIN tbSincronizacao AS B ON A.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NULL AND A.codcoligada = '" & vCodColigadaRM & "' group by a.idmov ORDER BY A.idmov"
    rsAtualizaEmprestimoItens.Open SqlAtualizaEmprestimoItens, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaEmprestimoItens.RecordCount > 0 Then
        While Not rsAtualizaEmprestimoItens.EOF
            'ATUALIZA TbEmprestimoItens
            SqlAtualizaGeral = "Update tbEmprestimoItens set idmov = '" & rsAtualizaEmprestimoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaEmprestimoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBanco
            Set rsAtualizaGeral = Nothing
            
            'ATUALIZA A TABELA TITMMOV BASEADO NA TABELA TBEMPRESTIMO
            SqlAtualizaGeral = "Update TITMMOV set idmov = '" & rsAtualizaEmprestimoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaEmprestimoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBancoSAP
            Set rsAtualizaGeral = Nothing
            
            'ATUALIZA A TABELA TMOV BASEADO NA TABELA TBEMPRESTIMO
            SqlAtualizaGeral = "Update TMOV set idmov = '" & rsAtualizaEmprestimoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaEmprestimoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBancoSAP
            
'            rsAtualizaEmprestimoItens.Fields(0) = rsAtualizaEmprestimoItens.Fields(0) + vDiferenca
            If rsAtualizaEmprestimoItens.Fields(0) + vDiferenca > vGuardaMaiorIDMOV Then vGuardaMaiorIDMOV = rsAtualizaEmprestimoItens.Fields(0) + vDiferenca
            rsAtualizaEmprestimoItens.MoveNext
        Wend
        'rsAtualizaEmprestimoItens.Update
    End If
    rsAtualizaEmprestimoItens.Close
    Set rsAtualizaEmprestimoItens = Nothing
   
    
    
    'ATUALIZA a tabela tbDevolucao
    SqlAtualizaDevolucao = "SELECT A.idmov FROM tbDevolucao AS A LEFT JOIN tbSincronizacao AS B ON A.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NULL AND A.codcoligada = '" & vCodColigadaRM & "' ORDER BY A.idmov"
    rsAtualizaDevolucao.Open SqlAtualizaDevolucao, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaDevolucao.RecordCount > 0 Then
        While Not rsAtualizaDevolucao.EOF
            SqlAtualizaGeral = "Update tbDevolucao set idmov = '" & rsAtualizaDevolucao.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaDevolucao.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBanco
            Set rsAtualizaGeral = Nothing
'            rsAtualizaDevolucao.Fields(0) = rsAtualizaDevolucao.Fields(0) + vDiferenca
            If rsAtualizaDevolucao.Fields(0) + vDiferenca > vGuardaMaiorIDMOV Then vGuardaMaiorIDMOV = rsAtualizaDevolucao.Fields(0) + vDiferenca
            rsAtualizaDevolucao.MoveNext
        Wend
        'rsAtualizaDevolucao.Update
    End If
    rsAtualizaDevolucao.Close
    Set rsAtualizaDevolucao = Nothing
   
    'ATUALIZA a tabela tbDevolucaoItens
    SqlAtualizaDevolucaoItens = "SELECT A.idmov FROM tbDevolucaoItens AS A LEFT JOIN tbSincronizacao AS B ON A.idmov = B.idmovsincronizado WHERE B.idmovsincronizado IS NULL AND A.codcoligada = '" & vCodColigadaRM & "' group by a.idmov ORDER BY A.idmov"
    rsAtualizaDevolucaoItens.Open SqlAtualizaDevolucaoItens, cnBanco, adOpenKeyset, adLockReadOnly
    If rsAtualizaDevolucaoItens.RecordCount > 0 Then
        While Not rsAtualizaDevolucaoItens.EOF
            SqlAtualizaGeral = "Update tbDevolucaoItens set idmov = '" & rsAtualizaDevolucaoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaDevolucaoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBanco
            Set rsAtualizaGeral = Nothing

            'ATUALIZA A TABELA TMOV BASEADO NA TABELA TBDEVOLUCAO
            SqlAtualizaGeral = "Update TMOV set idmov = '" & rsAtualizaDevolucaoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaDevolucaoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBancoSAP
            Set rsAtualizaGeral = Nothing
    
            'ATUALIZA A TABELA TITMMOV BASEADO NA TABELA TBDEVOLUCAO
            SqlAtualizaGeral = "Update TITMMOV set idmov = '" & rsAtualizaDevolucaoItens.Fields(0) + vDiferenca & "' where idmov = '" & rsAtualizaDevolucaoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBancoSAP
            Set rsAtualizaGeral = Nothing


            'ATUALIZA A TABELA TMORELAC BASEADO NA TABELA TBDEVOLUCAO
            SqlAtualizaGeral = "Update TMOVRELAC set idmovorigem = '" & rsAtualizaDevolucaoItens.Fields(0) + vDiferenca & "',idmovdestino = idmovdestino+ " & vDiferenca & " where idmovorigem = '" & rsAtualizaDevolucaoItens.Fields(0) & "'"
            rsAtualizaGeral.Open SqlAtualizaGeral, cnBancoSAP
            Set rsAtualizaGeral = Nothing


'            rsAtualizaDevolucaoItens.Fields(0) = rsAtualizaDevolucaoItens.Fields(0) + vDiferenca
            If rsAtualizaDevolucaoItens.Fields(0) + vDiferenca > vGuardaMaiorIDMOV Then vGuardaMaiorIDMOV = rsAtualizaDevolucaoItens.Fields(0) + vDiferenca
            rsAtualizaDevolucaoItens.MoveNext
        Wend
        'rsAtualizaDevolucaoItens.Update
    End If
    rsAtualizaDevolucaoItens.Close
    Set rsAtualizaDevolucaoItens = Nothing
    
End Function
