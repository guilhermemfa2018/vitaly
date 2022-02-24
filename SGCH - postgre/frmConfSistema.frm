VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmConfSistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Importação"
      TabPicture(0)   =   "frmConfSistema.frx":37E04
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Parametrizações"
      TabPicture(1)   =   "frmConfSistema.frx":37E20
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Empresa/Coligadas"
      TabPicture(2)   =   "frmConfSistema.frx":37E3C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Servidor - email"
      TabPicture(3)   =   "frmConfSistema.frx":37E58
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(1)=   "Frame10"
      Tab(3).ControlCount=   2
      Begin TabDlg.SSTab SSTab4 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   85
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Empresa/coligada ativa"
         TabPicture(0)   =   "frmConfSistema.frx":37E74
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Empresas/Coligadas"
         TabPicture(1)   =   "frmConfSistema.frx":37E90
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ListView3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chameleonButton4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "chameleonButton5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chameleonButton6"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "imgColigada"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
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
                  Picture         =   "frmConfSistema.frx":37EAC
                  Key             =   "OK"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmConfSistema.frx":388BE
                  Key             =   "EXC"
               EndProperty
            EndProperty
         End
         Begin SGCH.chameleonButton chameleonButton6 
            Height          =   615
            Left            =   8160
            TabIndex        =   106
            Tag             =   "Ativar coligada"
            ToolTipText     =   "Ativar coligada"
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
            MICON           =   "frmConfSistema.frx":392D0
            PICN            =   "frmConfSistema.frx":392EC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton chameleonButton5 
            Height          =   615
            Left            =   720
            TabIndex        =   104
            Tag             =   "Excluir Empresa/Coligada"
            ToolTipText     =   "Excluir Empresa/Coligada"
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
            MICON           =   "frmConfSistema.frx":39FC6
            PICN            =   "frmConfSistema.frx":39FE2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton chameleonButton4 
            Height          =   615
            Left            =   120
            TabIndex        =   103
            Tag             =   "Editar Empresa/Coligada"
            ToolTipText     =   "Editar Empresa/Coligada"
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
            MICON           =   "frmConfSistema.frx":3ACBC
            PICN            =   "frmConfSistema.frx":3ACD8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   3495
            Left            =   120
            TabIndex        =   102
            Top             =   1200
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   6165
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
            NumItems        =   0
         End
         Begin VB.Frame Frame2 
            Caption         =   "Dados da empresa/coligada ativa"
            Height          =   4335
            Left            =   -74880
            TabIndex        =   86
            Top             =   360
            Width           =   8775
            Begin VB.TextBox txtDadosEmpresa 
               Enabled         =   0   'False
               Height          =   285
               Index           =   11
               Left            =   1200
               TabIndex        =   105
               Tag             =   "Código da coligada"
               ToolTipText     =   "Código da coligada"
               Top             =   360
               Width           =   735
            End
            Begin SGCH.chameleonButton cmdCadastro 
               Height          =   615
               Index           =   16
               Left            =   720
               TabIndex        =   123
               Tag             =   "Incluir Empresa/Coligada"
               ToolTipText     =   "Incluir Empresa/Coligada"
               Top             =   3600
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
               MICON           =   "frmConfSistema.frx":3B9B2
               PICN            =   "frmConfSistema.frx":3B9CE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin SGCH.chameleonButton cmdCadastro 
               Height          =   615
               Index           =   15
               Left            =   120
               TabIndex        =   122
               Tag             =   "Nova Empresa/Coligada"
               ToolTipText     =   "Nova Empresa/Coligada"
               Top             =   3600
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
               MICON           =   "frmConfSistema.frx":3C6A8
               PICN            =   "frmConfSistema.frx":3C6C4
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
               Index           =   0
               Left            =   2040
               TabIndex        =   107
               Tag             =   "Razão social"
               ToolTipText     =   "Razão social"
               Top             =   360
               Width           =   3615
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   1
               Left            =   1200
               TabIndex        =   108
               Top             =   720
               Width           =   4455
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   2
               Left            =   1200
               TabIndex        =   109
               Top             =   1080
               Width           =   4455
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   3
               Left            =   1200
               TabIndex        =   110
               Top             =   1440
               Width           =   4455
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   4
               Left            =   2640
               TabIndex        =   112
               Top             =   1800
               Width           =   1575
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   5
               Left            =   1200
               TabIndex        =   113
               Top             =   2160
               Width           =   4455
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   6
               Left            =   1200
               TabIndex        =   114
               Top             =   2520
               Width           =   4455
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   7
               Left            =   1200
               TabIndex        =   115
               Top             =   2880
               Width           =   2055
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   8
               Left            =   3720
               TabIndex        =   116
               Top             =   2880
               Width           =   1935
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   9
               Left            =   1200
               TabIndex        =   117
               Top             =   3240
               Width           =   2055
            End
            Begin VB.TextBox txtDadosEmpresa 
               Height          =   285
               Index           =   10
               Left            =   3720
               TabIndex        =   118
               Top             =   3240
               Width           =   1935
            End
            Begin VB.ComboBox cboDadosEmpresa 
               Height          =   315
               ItemData        =   "frmConfSistema.frx":3D39E
               Left            =   1200
               List            =   "frmConfSistema.frx":3D3F3
               TabIndex        =   111
               Top             =   1800
               Width           =   735
            End
            Begin VB.Frame Frame6 
               Caption         =   "Logo"
               Height          =   3855
               Index           =   0
               Left            =   5760
               TabIndex        =   87
               Top             =   240
               Width           =   2895
               Begin VB.PictureBox Picture2 
                  Height          =   2775
                  Left            =   120
                  ScaleHeight     =   2715
                  ScaleWidth      =   2595
                  TabIndex        =   119
                  Top             =   240
                  Width           =   2655
                  Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
                     Height          =   2655
                     Left            =   120
                     Top             =   0
                     Width           =   2415
                     _ExtentX        =   4260
                     _ExtentY        =   4683
                     Image           =   "frmConfSistema.frx":3D463
                  End
                  Begin VB.Label Label59 
                     Alignment       =   2  'Center
                     Caption         =   "A Imagem não se encontra no local especificado"
                     Height          =   495
                     Left            =   240
                     TabIndex        =   88
                     Top             =   1200
                     Visible         =   0   'False
                     Width           =   2055
                  End
               End
               Begin SGCH.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   13
                  Left            =   720
                  TabIndex        =   121
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
                  MICON           =   "frmConfSistema.frx":3D47B
                  PICN            =   "frmConfSistema.frx":3D497
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSComDlg.CommonDialog cdlFoto 
                  Left            =   1800
                  Top             =   3240
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin SGCH.chameleonButton cmdCadastro 
                  Height          =   615
                  Index           =   12
                  Left            =   120
                  TabIndex        =   120
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
                  MICON           =   "frmConfSistema.frx":3E171
                  PICN            =   "frmConfSistema.frx":3E18D
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
            Begin VB.Label Label2 
               Caption         =   "Razão social:"
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label3 
               Caption         =   "Endereço:"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label Label4 
               Caption         =   "Bairro:"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade:"
               Height          =   255
               Left            =   120
               TabIndex        =   98
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label6 
               Caption         =   "UF:"
               Height          =   255
               Left            =   120
               TabIndex        =   97
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "CEP:"
               Height          =   255
               Left            =   2160
               TabIndex        =   96
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label Label8 
               Caption         =   "Email:"
               Height          =   255
               Left            =   120
               TabIndex        =   95
               Top             =   2280
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "Site:"
               Height          =   255
               Left            =   120
               TabIndex        =   94
               Top             =   2640
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Telefone:"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   3000
               Width           =   975
            End
            Begin VB.Label Label11 
               Caption         =   "Fax:"
               Height          =   255
               Left            =   3360
               TabIndex        =   92
               Top             =   3000
               Width           =   735
            End
            Begin VB.Label Label12 
               Caption         =   "CNPJ:"
               Height          =   255
               Left            =   120
               TabIndex        =   91
               Top             =   3360
               Width           =   735
            End
            Begin VB.Label Label13 
               Caption         =   "IE:"
               Height          =   255
               Left            =   3360
               TabIndex        =   90
               Top             =   3360
               Width           =   375
            End
            Begin VB.Label Label53 
               BackColor       =   &H8000000C&
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   3840
               Visible         =   0   'False
               Width           =   5415
            End
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4815
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   8493
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Gerais"
         TabPicture(0)   =   "frmConfSistema.frx":3EE67
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame8"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Aval. de Desempenho"
         TabPicture(1)   =   "frmConfSistema.frx":3EE83
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdCadastro(10)"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(2)=   "Frame5"
         Tab(1).Control(3)=   "Frame7"
         Tab(1).Control(4)=   "ListView1"
         Tab(1).Control(5)=   "cmdCadastro(3)"
         Tab(1).Control(6)=   "cmdCadastro(2)"
         Tab(1).Control(7)=   "cmdCadastro(4)"
         Tab(1).Control(8)=   "cmdCadastro(5)"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Absenteísmo"
         TabPicture(2)   =   "frmConfSistema.frx":3EE9F
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdCadastro(9)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdCadastro(8)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdCadastro(7)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "cmdCadastro(6)"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "Frame11"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Frame12"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Frame13"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "ListView2"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "Frame14"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Integração"
         TabPicture(3)   =   "frmConfSistema.frx":3EEBB
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Check4"
         Tab(3).Control(1)=   "Frame15"
         Tab(3).Control(2)=   "Frame3"
         Tab(3).Control(3)=   "SSTab3"
         Tab(3).ControlCount=   4
         Begin VB.CheckBox Check4 
            Caption         =   "Integrar o SGCH à um dos sistemas abaixo relacionados."
            Height          =   255
            Left            =   -74880
            TabIndex        =   80
            Top             =   480
            Width           =   4695
         End
         Begin VB.Frame Frame15 
            Caption         =   "Sistema "
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
            ForeColor       =   &H8000000D&
            Height          =   1335
            Left            =   -70680
            TabIndex        =   76
            Top             =   840
            Width           =   4455
            Begin VB.OptionButton chkIntegra 
               Caption         =   "SAP"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   79
               Top             =   960
               Width           =   2535
            End
            Begin VB.OptionButton chkIntegra 
               Caption         =   "Microsiga"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   78
               Top             =   600
               Width           =   2415
            End
            Begin VB.OptionButton chkIntegra 
               Caption         =   "Totvs - RM Labore (11.40)"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "SGBD "
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
            ForeColor       =   &H8000000D&
            Height          =   1335
            Left            =   -74880
            TabIndex        =   73
            Top             =   840
            Width           =   4095
            Begin VB.OptionButton optIntegra 
               Caption         =   "SQL Server"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   240
               Value           =   -1  'True
               Width           =   2895
            End
            Begin VB.OptionButton optIntegra 
               Caption         =   "Oracle"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   74
               Top             =   600
               Width           =   2895
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   2415
            Left            =   -74880
            TabIndex        =   64
            Top             =   2280
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   4260
            _Version        =   393216
            TabHeight       =   520
            Enabled         =   0   'False
            TabCaption(0)   =   "RM Sistemas"
            TabPicture(0)   =   "frmConfSistema.frx":3EED7
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label22"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label23"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label24"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label25"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "txtIntegra(3)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "txtIntegra(2)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txtIntegra(1)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtIntegra(0)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Microsiga"
            TabPicture(1)   =   "frmConfSistema.frx":3EEF3
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "SAP"
            TabPicture(2)   =   "frmConfSistema.frx":3EF0F
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
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
               Left            =   120
               TabIndex        =   65
               Top             =   960
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
               Index           =   1
               Left            =   3000
               TabIndex        =   66
               Top             =   960
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
               Left            =   120
               TabIndex        =   67
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
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   3000
               PasswordChar    =   "*"
               TabIndex        =   68
               Top             =   1560
               Width           =   2655
            End
            Begin VB.Label Label25 
               Caption         =   "Nome do SERVIDOR:"
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   720
               Width           =   2175
            End
            Begin VB.Label Label24 
               Caption         =   "Nome do BANCO:"
               Height          =   255
               Left            =   3000
               TabIndex        =   71
               Top             =   720
               Width           =   2775
            End
            Begin VB.Label Label23 
               Caption         =   "Usuário:"
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   1320
               Width           =   2655
            End
            Begin VB.Label Label22 
               Caption         =   "Senha:"
               Height          =   255
               Left            =   3000
               TabIndex        =   69
               Top             =   1320
               Width           =   2655
            End
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   10
            Left            =   -67080
            TabIndex        =   63
            Tag             =   "Gerar Avaliações de Desempenho"
            ToolTipText     =   "Gerar Avaliações de Desempenho"
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
            MICON           =   "frmConfSistema.frx":3EF2B
            PICN            =   "frmConfSistema.frx":3EF47
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame4 
            Caption         =   "Avaliar após"
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
            Left            =   -74880
            TabIndex        =   55
            Top             =   420
            Width           =   1335
            Begin VB.TextBox txtCadParametro 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label17 
               Caption         =   "dias"
               Height          =   255
               Left            =   840
               TabIndex        =   57
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Tipo"
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
            Left            =   -73440
            TabIndex        =   52
            Top             =   420
            Width           =   3135
            Begin VB.OptionButton optCadParametro 
               Caption         =   "Periódico"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   54
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optCadParametro 
               Caption         =   "Experiência"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Identificador"
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
            Left            =   -70200
            TabIndex        =   50
            Top             =   420
            Width           =   1455
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Identificador"
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
            Left            =   6720
            TabIndex        =   48
            Top             =   420
            Width           =   1455
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   "ID"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   1215
            End
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2835
            Left            =   120
            TabIndex        =   47
            Top             =   1860
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5001
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
            NumItems        =   0
         End
         Begin VB.Frame Frame13 
            Caption         =   "Pontos (negativos)"
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
            Left            =   4560
            TabIndex        =   41
            Top             =   420
            Width           =   2055
            Begin VB.TextBox txtABS 
               Height          =   285
               Left            =   120
               TabIndex        =   42
               Tag             =   "Pontos"
               ToolTipText     =   "Pontos"
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Entre (ocorrências)"
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
            Left            =   2280
            TabIndex        =   37
            Top             =   420
            Width           =   2175
            Begin VB.ComboBox cboABS 
               Height          =   315
               Index           =   2
               ItemData        =   "frmConfSistema.frx":3FC21
               Left            =   1200
               List            =   "frmConfSistema.frx":3FC7F
               TabIndex        =   39
               Tag             =   "Ocorrência 2"
               ToolTipText     =   "Ocorrência 2"
               Top             =   240
               Width           =   735
            End
            Begin VB.ComboBox cboABS 
               Height          =   315
               Index           =   1
               ItemData        =   "frmConfSistema.frx":3FCFB
               Left            =   120
               List            =   "frmConfSistema.frx":3FD59
               TabIndex        =   38
               Tag             =   "Ocorrência 1"
               ToolTipText     =   "Ocorrência 1"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label20 
               Caption         =   "à"
               Height          =   255
               Left            =   960
               TabIndex        =   40
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Tipo"
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
            Left            =   120
            TabIndex        =   35
            Top             =   420
            Width           =   2055
            Begin VB.ComboBox cboABS 
               Height          =   315
               Index           =   0
               ItemData        =   "frmConfSistema.frx":3FDD5
               Left            =   120
               List            =   "frmConfSistema.frx":3FDDF
               TabIndex        =   36
               Tag             =   "Tipo"
               Text            =   "Ausência"
               ToolTipText     =   "Tipo"
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Gerais "
            Height          =   4095
            Left            =   -74880
            TabIndex        =   20
            Top             =   420
            Width           =   8775
            Begin VB.Frame Frame19 
               Caption         =   "Médias "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               TabIndex        =   127
               Top             =   2640
               Width           =   3495
               Begin VB.TextBox txtCadParametro 
                  Height          =   285
                  Index           =   0
                  Left            =   2280
                  TabIndex        =   29
                  Tag             =   "Média para aprovação"
                  ToolTipText     =   "Média para aprovação"
                  Top             =   240
                  Width           =   975
               End
               Begin VB.TextBox txtCadParametro 
                  Height          =   285
                  Index           =   1
                  Left            =   2280
                  TabIndex        =   30
                  Tag             =   "Média para aprovação com restrição"
                  ToolTipText     =   "Média para aprovação com restrição"
                  Top             =   660
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Caption         =   "Aprovação (%):"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   129
                  Top             =   360
                  Width           =   1935
               End
               Begin VB.Label Label15 
                  Caption         =   "Aprovação com restrição (%):"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   128
                  Top             =   720
                  Width           =   2535
               End
            End
            Begin VB.Frame Frame18 
               Height          =   1095
               Left            =   3720
               TabIndex        =   125
               Top             =   2640
               Width           =   4935
               Begin VB.TextBox txtCadParametro 
                  Height          =   285
                  Index           =   3
                  Left            =   240
                  TabIndex        =   32
                  Top             =   600
                  Width           =   615
               End
               Begin VB.CheckBox Check10 
                  Caption         =   "Gerar treinamentos obrigatórios"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   34
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.CheckBox Check9 
                  Caption         =   "Gerar treinamentos introdutórios"
                  Height          =   255
                  Left            =   2160
                  TabIndex        =   33
                  Top             =   240
                  Width           =   2655
               End
               Begin VB.CheckBox Check8 
                  Caption         =   "Reciclar colaboradores afastados"
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
                  TabIndex        =   31
                  Top             =   0
                  Width           =   3255
               End
               Begin VB.Label Label26 
                  Caption         =   "Período superior à:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   130
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.Label Label27 
                  Caption         =   "dias"
                  Height          =   255
                  Left            =   960
                  TabIndex        =   126
                  Top             =   600
                  Width           =   615
               End
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Calcular automaticamente tempo de experiência dos colaboradores"
               Height          =   615
               Left            =   120
               TabIndex        =   25
               Top             =   1800
               Width           =   6855
            End
            Begin VB.Frame Frame17 
               Height          =   855
               Left            =   3360
               TabIndex        =   124
               Top             =   240
               Width           =   5295
               Begin MSComDlg.CommonDialog cdlTXT2 
                  Left            =   4800
                  Top             =   240
                  _ExtentX        =   847
                  _ExtentY        =   847
                  _Version        =   393216
               End
               Begin SGCH.chameleonButton cmdCadastro 
                  Height          =   255
                  Index           =   17
                  Left            =   4680
                  TabIndex        =   28
                  Tag             =   "Localizar"
                  ToolTipText     =   "Localizar"
                  Top             =   360
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   450
                  BTYPE           =   2
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
                  MICON           =   "frmConfSistema.frx":3FDF5
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.CheckBox Check6 
                  Caption         =   " Atualizações automáticas"
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
                  TabIndex        =   26
                  Top             =   0
                  Width           =   2775
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
                  TabIndex        =   27
                  Text            =   "Informe o caminho do executável: AtualizaSGCH.exe"
                  Top             =   360
                  Width           =   4455
               End
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Exibir avisos ao logar"
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   1560
               Width           =   2175
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Gerar treinamentos obrigatórios"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   780
               Width           =   2535
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Ativar arquivo de LOG"
               Height          =   375
               Left            =   120
               TabIndex        =   23
               Top             =   1140
               Width           =   2775
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Gerar treinamentos introdutórios"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   300
               Width           =   2775
            End
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   6
            Left            =   1920
            TabIndex        =   43
            Tag             =   "Excluir Absenteísmo"
            ToolTipText     =   "Excluir Absenteísmo"
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
            MICON           =   "frmConfSistema.frx":3FE11
            PICN            =   "frmConfSistema.frx":3FE2D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   7
            Left            =   1320
            TabIndex        =   44
            Tag             =   "Editar Absenteísmo"
            ToolTipText     =   "Editar Absenteísmo"
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
            MICON           =   "frmConfSistema.frx":40B07
            PICN            =   "frmConfSistema.frx":40B23
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   8
            Left            =   720
            TabIndex        =   45
            Tag             =   "Novo Absenteísmo"
            ToolTipText     =   "Novo Absenteísmo"
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
            MICON           =   "frmConfSistema.frx":417FD
            PICN            =   "frmConfSistema.frx":41819
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   9
            Left            =   120
            TabIndex        =   46
            Tag             =   "Incluir Absenteísmo"
            ToolTipText     =   "Incluir Absenteísmo"
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
            MICON           =   "frmConfSistema.frx":424F3
            PICN            =   "frmConfSistema.frx":4250F
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
            Height          =   2835
            Left            =   -74880
            TabIndex        =   58
            Top             =   1860
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   5001
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
            NumItems        =   0
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   -73080
            TabIndex        =   59
            Tag             =   "Excluir ADP"
            ToolTipText     =   "Excluir ADP"
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
            MICON           =   "frmConfSistema.frx":431E9
            PICN            =   "frmConfSistema.frx":43205
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   2
            Left            =   -73680
            TabIndex        =   60
            Tag             =   "Editar ADP"
            ToolTipText     =   "Editar ADP"
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
            MICON           =   "frmConfSistema.frx":43EDF
            PICN            =   "frmConfSistema.frx":43EFB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   4
            Left            =   -74280
            TabIndex        =   61
            Tag             =   "Novo ADP"
            ToolTipText     =   "Novo ADP"
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
            MICON           =   "frmConfSistema.frx":44BD5
            PICN            =   "frmConfSistema.frx":44BF1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin SGCH.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   5
            Left            =   -74880
            TabIndex        =   62
            Tag             =   "Incluir ADP"
            ToolTipText     =   "Incluir ADP"
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
            MICON           =   "frmConfSistema.frx":458CB
            PICN            =   "frmConfSistema.frx":458E7
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
      Begin VB.Frame Frame10 
         Caption         =   "Autenticação de email "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   16
         Top             =   1680
         Width           =   8895
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
         Begin VB.Label Label19 
            Caption         =   "Senha:"
            Height          =   255
            Left            =   4920
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label18 
            Caption         =   "Usuário:"
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -74880
         TabIndex        =   14
         Top             =   600
         Width           =   8895
         Begin VB.TextBox txtEmail 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Tag             =   "Endereço do servidor de SMTP"
            ToolTipText     =   "Endereço do servidor de SMTP"
            Top             =   240
            Width           =   8655
         End
         Begin VB.Label Label16 
            Caption         =   "Ex: smtp.mail.yahoo.com.br ou 192.168.1.1"
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione a tabela a ser importada "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   -74880
         TabIndex        =   7
         Top             =   420
         Width           =   9015
         Begin VB.Frame Frame16 
            Caption         =   "Importar colaboradores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   5160
            TabIndex        =   81
            Top             =   240
            Width           =   3735
            Begin SGCH.chameleonButton cmdCadastro 
               Height          =   375
               Index           =   14
               Left            =   1560
               TabIndex        =   84
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   2
               TX              =   "Importar"
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
               MICON           =   "frmConfSistema.frx":465C1
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin SGCH.chameleonButton cmdCadastro 
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   83
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               BTYPE           =   2
               TX              =   "Localizar..."
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
               MICON           =   "frmConfSistema.frx":465DD
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   82
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
         Begin VB.OptionButton Option6 
            Caption         =   "Setores"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Departamentos"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2280
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Escolaridade"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1800
            Width           =   2175
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Avaliação do treinamento"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Habilidades funcionais"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cargos"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
         Begin SGCH.chameleonButton chameleonButton1 
            Height          =   735
            Left            =   240
            TabIndex        =   0
            Tag             =   "Importar dados"
            ToolTipText     =   "Importar dados"
            Top             =   3600
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
            MICON           =   "frmConfSistema.frx":465F9
            PICN            =   "frmConfSistema.frx":46615
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
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   5880
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
      MICON           =   "frmConfSistema.frx":47A6F
      PICN            =   "frmConfSistema.frx":47A8B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   5880
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
      MICON           =   "frmConfSistema.frx":48765
      PICN            =   "frmConfSistema.frx":48781
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

Private Sub chameleonButton1_Click()
On Error Resume Next
    If MsgBox("Deseja realmente importar os dados das tabelas selecionadas?", vbQuestion + vbYesNo, "SGCH") = vbNo Then
        Exit Sub
    End If
    
    If Option1.Value = True Then ImportaDadosCargo
    If Option2.Value = True Then ImportaDadosHabilidade
    If Option3.Value = True Then ImportaDadosAvaliacao
    If Option4.Value = True Then ImportaDadosEscolaridade
    If Option5.Value = True Then ImportaDadosDepartamento
    If Option6.Value = True Then ImportaDadosSetor
    
    If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False Then
        MsgBox "Nenhuma tabela selecionada. Marque a tabela a ser importada", vbInformation, "SGCH"
        Exit Sub
    End If
    
    'A ROTINA ABAIXO VC SELECIONA UM PROCESSO Q ESTA NA MEMORIA P SER REMOVIDO
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'EXCEL.EXE'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    '--------------------------------------------------------------------------
    MsgBox "Dados importados com sucesso. Para vizualisar os dados feche a tabela e abra novamente"
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de CARGOS..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
            If .Range("A" & J).Value = "" Then Exit For
            rsCargos.AddNew
            rsCargos.Fields(0) = .Range("A" & J).Value 'Código do CARGO
            rsCargos.Fields(1) = .Range("B" & J).Value 'Código do CBO
            rsCargos.Fields(2) = .Range("C" & J).Value 'Nome do CARGO
            rsCargos.Fields(5) = vCodcoligada 'Codigo da coligada
            J = J + 1
        End With
    Next
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de cargos do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - RAMO DE ATIVIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Habilidade..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de Habilidades do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - AVALIACAO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Avaliação do Treinamento..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de Avaliação do Treinamento do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - ESCOLARIDADE
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Escolaridade..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de Escolaridade do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - DEPARTAMENTO
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Departamento..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de Departamentos do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
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
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
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
    frmMenu2.ProgressBar1.Max = J
    
    'PREENCHE CÉLULAS DESEJADAS - SETORES
    '**********************************************************************
    Legenda = "Aguarde, importando tabela de Setores..."
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    
    J = 2
    For X = 1 To frmMenu2.ProgressBar1.Max
        With Plan
            frmMenu2.ProgressBar1.Value = X
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
    frmMenu2.ProgressBar1.Value = 0
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
    MsgBox "Existem dados cadastrados na tabela de Setores do sistema. Para que a importação seja realizada ela deve estar vazia", vbInformation, "Atenção"
    Legenda = "ERRO na importação de dados"
    frmMenu2.StatusBar1.Panels(3).Text = Legenda
    Exit Sub
End Sub

Private Sub chameleonButton4_Click()
    AlteraColigada
    SSTab4.Tab = 0
End Sub

Private Sub chameleonButton5_Click()
    MsgBox "Rotina em desenvolvimento", vbInformation, "SGCH"
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
        cmdCadastro(17).Enabled = True
        Text2 = ""
    Else
        Text2.Enabled = False
        cmdCadastro(17).Enabled = False
        Text2 = "Informe o caminho do executável: AtualizaSGCH.exe"
    End If
End Sub

Private Sub Check8_Click()
    If Check8.Value = 1 Then
        txtCadParametro(3).Enabled = True
        Check9.Enabled = True
        Check10.Enabled = True
    Else
        txtCadParametro(3).Enabled = False
        Check9.Enabled = False
        Check10.Enabled = False
        txtCadParametro(3).Text = ""
        Check9.Value = 0
        Check10.Value = 0
    End If
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        If MsgBox("Deseja salvar os dados de parametrização?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            GravaParametros
            gravaLog "Mádia para aprovação: " & txtCadParametro(0), "Gerar introdutório: " & Check3.Value, "Aprovação com restrição: " & txtCadParametro(1)
            Pesquisa = 0
            'Unload Me
        End If
    Case 1
        If MsgBox("Deseja sair da tela configurações do sistema?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = 0
            Unload Me
            Set frmConfSistema = Nothing
        End If
    Case 2
        AlteraAvaliacao
    Case 3
        If MsgBox("Deseja EXCLUIR esse período de avaliação?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            ExcluirItemLV ListView1
        '    LimpaControlesAprovadorReq
        End If
    Case 4
        LimpaControlesAvaliacao
    Case 5
        IncluirAvaliacao
    Case 6
        If MsgBox("Deseja EXCLUIR essa ocorrência?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            ExcluirItemLV ListView2
        End If
    Case 7
        AlteraABS
    Case 8
        LimpaControlesABS
    Case 9
        IncluirABS
    Case 10
        If ListView1.ListItems.Count > 0 Then
            carregaADP "TODOS"
            MsgBox "Rotina de Avaliação de Desempenho efetuada com sucesso!"
        Else
            MsgBox "É necessário cadastrar primeiramente os períodos de Avaliação de Desempenho Profissional"
        End If
    Case 11
        'carregar arquivo texto
        With cdlTXT
            .Filter = "(Arquivo *.TXT)|*.txt"
            .ShowOpen
            Caminho2 = .FileName
        End With
        Text1 = Caminho2
        If Text1.Text <> "" Then cmdCadastro(14).Enabled = True
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
        importaColaboradores
    Case 15
        LimpaControlesColigada
    Case 16
        IncluirColigada
        'criaUsuEMenu Val(txtDadosEmpresa(11) - 1)
    Case 17
'        carregaPasta
        With cdlTXT2
            .Filter = "(AtualizaSGCH *.EXE)|*.exe"
            .ShowOpen
            Caminho3 = .FileName
        End With
        Text2 = Caminho3
    End Select
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

Private Sub Form_Load()
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    SSTab4.Tab = 0
    
    SSTab3.TabEnabled(1) = False
    SSTab3.TabEnabled(2) = False
    
    CarregaParametros
    configControles
    listview_cabecalho
    Compoe_Listview
    LimpaControlesAvaliacao
    LimpaControlesABS
    'LimpaControlesColigada
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Avaliar após", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Tipo", ListView1.Width / 3
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 12
    ListView2.ColumnHeaders.Add , , "Tipo", ListView2.Width / 10
    ListView2.ColumnHeaders.Add , , "Ocorrência1", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Ocorrência2", ListView2.Width / 8
    ListView2.ColumnHeaders.Add , , "Pontos%", ListView2.Width / 10
    
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
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub LimpaControlesABS()
    Dim X As Integer
    cboABS(0).Text = ""
    cboABS(1).Text = ""
    cboABS(2).Text = ""
    txtABS = ""
    If ListView2.ListItems.Count > 0 Then
        Label21.Caption = Format(GeraCodigo(ListView2), "00")
    Else
        Label21.Caption = Format(Val(Label21) + 1, "00")
    End If
End Sub

Private Sub LimpaControlesAvaliacao()
    Dim X As Integer
    txtCadParametro(2) = ""
    optCadParametro(0).Value = True
    If ListView1.ListItems.Count > 0 Then
        Label14.Caption = Format(GeraCodigo(ListView1), "00")
    Else
        Label14.Caption = Format(Val(Label14) + 1, "00")
    End If
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
    Label53 = ""
    txtDadosEmpresa(0).SetFocus
End Sub

Private Sub Compoe_Listview()
    Dim rsAD As New ADODB.Recordset
    Dim sqlAD As String
    Dim rsABS As New ADODB.Recordset
    Dim sqlABS As String
    Dim rsColigadas As New ADODB.Recordset
    Dim sqlColigadas As String
    
    Dim ItemLst As ListItem
    Dim X As Integer
    
    ' Compoe Listview1
    sqlAD = "Select * from tbAvaliacaoDesempenho where codcoligada = '" & vCodcoligada & "' Order by id"
    rsAD.Open sqlAD, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsAD.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsAD.Fields(0), "00"))
        ItemLst.SubItems(1) = "" & rsAD.Fields(1)
        ItemLst.SubItems(2) = "" & rsAD.Fields(2)
        rsAD.MoveNext
        X = X + 1
    Wend
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
    rsAD.Close
    Set rsAD = Nothing
    
    ' Compoe Listview2
    sqlABS = "Select * from tbABS where codcoligada ='" & vCodcoligada & "' Order by id"
    rsABS.Open sqlABS, cnBanco, adOpenKeyset, adLockOptimistic
    X = 0
    While Not rsABS.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsABS.Fields(0), "00"))
        ItemLst.SubItems(1) = "" & rsABS.Fields(1)
        ItemLst.SubItems(2) = "" & rsABS.Fields(2)
        ItemLst.SubItems(3) = "" & rsABS.Fields(3)
        ItemLst.SubItems(4) = "" & rsABS.Fields(4)
        rsABS.MoveNext
        X = X + 1
    Wend
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwDescending
    rsABS.Close
    Set rsABS = Nothing

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

Private Function GeraCodigo(LV As ListView)
    Dim X As Integer
    X = 1
    LV.SortOrder = lvwDescending
    LV.ListItems.Item(X).Selected = True
    GeraCodigo = LV.ListItems.Item(X) + 1
    LV.SortOrder = lvwAscending
    Exit Function
End Function

Private Sub IncluirAvaliacao()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.ListItems.Item(X) = Me.Label14.Caption Then
                Label14.Caption = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadParametro(2).Text
                If optCadParametro(0).Value = True Then
                    ListView1.SelectedItem.ListSubItems.Item(2) = "Experiência"
                Else
                    ListView1.SelectedItem.ListSubItems.Item(2) = "Periódico"
                End If
                Y = ListView1.ListItems.Count
                Me.ListView1.Sorted = True
                Me.ListView1.SortKey = 0
                Me.ListView1.SortOrder = lvwAscending
                LimpaControlesAvaliacao
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , Label14)
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , Label14)
        Y = ListView1.ListItems.Count
        Me.ListView1.Sorted = True
        Me.ListView1.SortKey = 0
        Me.ListView1.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = txtCadParametro(2).Text
    If optCadParametro(0).Value = True Then
        ItemLst.SubItems(2) = "Experiência"
    Else
        ItemLst.SubItems(2) = "Periódico"
    End If
    Me.ListView1.SortOrder = lvwAscending
    txtCadParametro(2).SetFocus
    LimpaControlesAvaliacao
End Sub

Private Sub IncluirABS()
    If ValidaABS = False Then Exit Sub
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True
            If ListView2.ListItems.Item(X) = Me.Label21.Caption Then
                Label21.Caption = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = cboABS(0).Text
                ListView2.SelectedItem.ListSubItems.Item(2) = cboABS(1).Text
                ListView2.SelectedItem.ListSubItems.Item(3) = cboABS(2).Text
                ListView2.SelectedItem.ListSubItems.Item(4) = txtABS.Text
                Y = ListView2.ListItems.Count
                Me.ListView2.Sorted = True
                Me.ListView2.SortKey = 0
                Me.ListView2.SortOrder = lvwAscending
                LimpaControlesABS
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , Label21)
        Y = ListView2.ListItems.Count
    Else
        Set ItemLst = ListView2.ListItems.Add(, , Label21)
        Y = ListView2.ListItems.Count
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 0
        Me.ListView2.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = cboABS(0).Text
    ItemLst.SubItems(2) = cboABS(1).Text
    ItemLst.SubItems(3) = cboABS(2).Text
    ItemLst.SubItems(4) = txtABS.Text
    Me.ListView2.SortOrder = lvwAscending
    cboABS(0).SetFocus
    LimpaControlesABS
End Sub

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

Private Function ValidaABS()
    ValidaABS = False
    Dim X As Integer
    For X = 0 To 1
        If cboABS(X).Text = "" Then
            MsgBox "Favor informar o campo " & Me.cboABS(X).Tag, vbInformation, "Atenção"
            Me.cboABS(X).SetFocus
            Exit Function
        End If
    
    Next
    If txtABS.Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtABS.Tag, vbInformation, "Atenção"
        Me.txtABS.SetFocus
        Exit Function
    End If
    ValidaABS = True
End Function

Private Sub AlteraAvaliacao()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.Label14.Caption = ListView1.ListItems.Item(X)
    Me.txtCadParametro(2).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    If ListView1.SelectedItem.ListSubItems.Item(2) = "Experiência" Then
        Me.optCadParametro(0).Value = True
    Else
        Me.optCadParametro(1).Value = True
    End If
End Sub

Private Sub AlteraABS()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.Label21.Caption = ListView2.ListItems.Item(X)
    Me.cboABS(0).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.cboABS(1).Text = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.cboABS(2).Text = ListView2.SelectedItem.ListSubItems.Item(3)
    Me.txtABS.Text = ListView2.SelectedItem.ListSubItems.Item(4)
End Sub

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
    Dim rsParametros As New ADODB.Recordset
    Dim sqlParametros As String
    Dim rsEmpresa As New ADODB.Recordset
    Dim sqlEmpresa As String
    Dim rsConfEmail As New ADODB.Recordset
    Dim sqlConfEmail As String
    Dim rsIntegracao As New ADODB.Recordset
    Dim sqlIntegracao As String
    
    If Text1.Text = "" Then cmdCadastro(14).Enabled = False
    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsParametros.EOF Then
        txtCadParametro(0) = rsParametros.Fields(0)
        txtCadParametro(1) = rsParametros.Fields(2)
        If Not IsNull(rsParametros.Fields(11)) And rsParametros.Fields(11) <> 0 Then
            Check8.Value = 1
            txtCadParametro(3) = rsParametros.Fields(11)
        Else
            txtCadParametro(3).Enabled = False
            Check9.Enabled = False
            Check10.Enabled = False
        End If
        If rsParametros.Fields(1) = "S" Then
            Check3.Value = 1
        Else
            Check3.Value = 0
        End If
        
        If rsParametros.Fields(7) = "S" Then
            Check5.Value = 1
        Else
            Check5.Value = 0
        End If
        If rsParametros.Fields(10) = "S" Then
            Check7.Value = 1
        Else
            Check7.Value = 0
        End If
        
        If rsParametros.Fields(8) = "S" Then
            Check6.Value = 1
            Text2.Text = rsParametros.Fields(9)
        Else
            Check6.Value = 0
        End If
        
        If rsParametros.Fields(5) = "S" Then
            Check4.Value = 1
        Else
            Check4.Value = 0
        End If
        
        If rsParametros.Fields(4) = "S" Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        
        If rsParametros.Fields(3) = "S" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        
        If Check8.Value = 1 Then
            If rsParametros.Fields(12) = "S" Then
                Check9.Value = 1
            Else
                Check9.Value = 0
            End If
            If rsParametros.Fields(13) = "S" Then
                Check10.Value = 1
            Else
                Check10.Value = 0
            End If
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
    rsIntegracao.Open sqlIntegracao, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsIntegracao.EOF Then
        If rsIntegracao.Fields(0) = 1 Then optIntegra(0).Value = True
        If rsIntegracao.Fields(0) = 2 Then optIntegra(1).Value = True
        If rsIntegracao.Fields(1) = 1 Then chkIntegra(0).Value = True
        If rsIntegracao.Fields(1) = 2 Then chkIntegra(1).Value = True
        If rsIntegracao.Fields(1) = 3 Then chkIntegra(2).Value = True
        
        txtIntegra(0).Text = rsIntegracao.Fields(3)
        txtIntegra(1).Text = rsIntegracao.Fields(4)
        txtIntegra(2).Text = rsIntegracao.Fields(5)
        txtIntegra(3).Text = rsIntegracao.Fields(6)
    End If
'*********************
    
    rsConfEmail.Close
    Set rsConfEmail = Nothing
    rsEmpresa.Close
    Set rsEmpresa = Nothing
    rsIntegracao.Close
    Set rsIntegracao = Nothing
    Exit Sub
TrataErro1:
    Label59.Visible = True
    Resume Next
End Sub

Private Sub GravaParametros()
'On Error Resume Next
    If ValidaCampo = False Then Exit Sub
    If Check6.Value = 1 And Text2 = "" Then
        MsgBox "Informe o caminho do executável: AtualizaSGCH.exe", vbCritical, "SGCH"
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
    
    cnBanco.BeginTrans
    
    sqlDeletar = "Delete from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco

    sqlParametros = "Select * from tbparametros where codcoligada = '" & vCodcoligada & "'"
    rsParametros.Open sqlParametros, cnBanco, adOpenKeyset, adLockOptimistic
    rsParametros.AddNew
    rsParametros.Fields(0) = txtCadParametro(0)
    rsParametros.Fields(2) = txtCadParametro(1)
    If Check8.Value = 1 Then
        rsParametros.Fields(11) = txtCadParametro(3)
    Else
        rsParametros.Fields(11) = 0
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
        If Check10.Value = 1 Then
            rsParametros.Fields(13) = "S"
        Else
            rsParametros.Fields(13) = "N"
        End If
    Else
        rsParametros.Fields(12) = "N"
        rsParametros.Fields(13) = "N"
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
            MsgBox "Os dados informados para conexão não estão corretos", vbCritical, "Conexão TOTVS"
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
        If chkIntegra(1).Value = True Then rsIntegracao.Fields(1) = 2
        If chkIntegra(2).Value = True Then rsIntegracao.Fields(1) = 3
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
    MediaGlobal = txtCadParametro(0)
    vAprovadoRest = txtCadParametro(1)
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
        vAfastDias = txtCadParametro(3)
        If Check9.Value = 1 Then
            vAfastTreiInt = "S"
        Else
            vAfastTreiInt = "N"
        End If
        If Check10.Value = 1 Then
            vAfastTreiObr = "S"
        Else
            vAfastTreiObr = "N"
        End If
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
    
    Dim Reg As Object
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\SGCH\" & "sLogoEmpresa", Label53 'Logo da empresa
    Set Reg = Nothing
    
    sqlDeletar = "Delete from tbAvaliacaoDesempenho where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    sqlConfAD = "select * from tbAvaliacaoDesempenho"
    rsConfAD.Open sqlConfAD, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsConfAD.AddNew
        rsConfAD.Fields(0) = ListView1.ListItems.Item(X)
        rsConfAD.Fields(1) = Val(ListView1.SelectedItem.ListSubItems.Item(1))
        rsConfAD.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsConfAD.Fields(3) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsConfAD.EOF Then rsConfAD.Update
    
    sqlDeletar = "Delete from tbABS where codcoligada = '" & vCodcoligada & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    sqlConfABS = "select * from tbABS"
    rsConfABS.Open sqlConfABS, cnBanco, adOpenKeyset, adLockOptimistic
    For X = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(X).Selected = True
        rsConfABS.AddNew
        rsConfABS.Fields(0) = ListView2.ListItems.Item(X)
        rsConfABS.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
        rsConfABS.Fields(2) = Val(ListView2.SelectedItem.ListSubItems.Item(2))
        If ListView2.SelectedItem.ListSubItems.Item(3) = "" Or ListView2.SelectedItem.ListSubItems.Item(3) = 0 Then
            rsConfABS.Fields(3) = 365
        Else
            rsConfABS.Fields(3) = Val(ListView2.SelectedItem.ListSubItems.Item(3))
        End If
        rsConfABS.Fields(4) = Val(ListView2.SelectedItem.ListSubItems.Item(4))
        rsConfABS.Fields(5) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsConfABS.EOF Then rsConfABS.Update
    cnBanco.CommitTrans
    MsgBox "Os dados de configuração do sistema foram salvos com sucesso", vbInformation, "SGCH"
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
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
        MsgBox "Nenhum empresa/coligada cadastrada. Favor informar os dados da empresa/coligada", vbInformation, "Atenção"
        SSTab1.Tab = 2
        Exit Function
    End If
    
    If txtCadParametro(0) = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadParametro(0).Tag, vbInformation, "Atenção"
        Me.txtCadParametro(0).SetFocus
        Exit Function
    End If
    If txtCadParametro(1) = "" Then
        MsgBox "Favor informar o campo " & Me.txtCadParametro(1).Tag, vbInformation, "Atenção"
        Me.txtCadParametro(1).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function ValidaDadosColigada()
    ValidaDadosColigada = False
    If ListView3.ListItems.Count = 0 Then
        If txtDadosEmpresa(0) = "" Then
            MsgBox "Favor informar o campo " & Me.txtDadosEmpresa(0).Tag, vbInformation, "Atenção"
            Me.txtDadosEmpresa(0).SetFocus
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
    If vInc = "N" Then
        cmdCadastro(12).UseGreyscale = True
        cmdCadastro(12).DragMode = 1
        cmdCadastro(12).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(13).UseGreyscale = True
        cmdCadastro(13).DragMode = 1
        cmdCadastro(13).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    End If
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

Private Sub ListView1_DblClick()
    AlteraAvaliacao
End Sub

Private Sub ListView2_DblClick()
    AlteraABS
End Sub

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
            MsgBox "Erro na linha: " & i + 1, vbCritical, "SGCH"
            Exit Sub
        End If
        insertDados
    Next
    MsgBox "Dados importados com sucesso!", vbInformation, "SGCH"
End Sub

Private Function ValidaDados()
    ValidaDados = False
    Dim Y As Integer
    For Y = 0 To 3
        If colheDados(Y) = "" Then
            MsgBox "Erro de consistência na fonte de dados", vbCritical, "SGCH"
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
