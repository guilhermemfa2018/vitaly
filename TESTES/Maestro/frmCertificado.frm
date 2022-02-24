VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCertificado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Certificados"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   Icon            =   "frmCertificado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7680
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
      MPTR            =   0
      MICON           =   "frmCertificado.frx":3469A
      PICN            =   "frmCertificado.frx":346B6
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
      Left            =   720
      TabIndex        =   4
      Tag             =   "Imprimir"
      ToolTipText     =   "Imprimir"
      Top             =   7680
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
      MPTR            =   0
      MICON           =   "frmCertificado.frx":35390
      PICN            =   "frmCertificado.frx":353AC
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
      Left            =   120
      TabIndex        =   3
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7680
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
      MPTR            =   0
      MICON           =   "frmCertificado.frx":36086
      PICN            =   "frmCertificado.frx":360A2
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
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Colaboradores"
      TabPicture(0)   =   "frmCertificado.frx":36D7C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Configurações"
      TabPicture(1)   =   "frmCertificado.frx":36D98
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Texto"
      TabPicture(2)   =   "frmCertificado.frx":36DB4
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Instruções"
      TabPicture(3)   =   "frmCertificado.frx":36DD0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Texto "
         Height          =   6495
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   10935
         Begin VB.TextBox Text5 
            Height          =   6135
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   10695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Composição do certificado "
         Height          =   6975
         Left            =   -74880
         TabIndex        =   8
         Top             =   360
         Width           =   10935
         Begin VB.CheckBox Check5 
            Caption         =   "Imprimir conteúdo programático do treinamento"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   6480
            Width           =   3855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3720
            TabIndex        =   66
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Frame Frame19 
            Caption         =   "Alinhamento Certificadora"
            Height          =   1095
            Left            =   4920
            TabIndex        =   57
            Top             =   1560
            Width           =   2415
            Begin VB.OptionButton Option8 
               Height          =   255
               Left            =   120
               TabIndex        =   61
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option7 
               Height          =   255
               Left            =   840
               TabIndex        =   60
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option6 
               Height          =   255
               Left            =   1440
               TabIndex        =   59
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option5 
               Height          =   255
               Left            =   2040
               TabIndex        =   58
               Top             =   720
               Width           =   255
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage9 
               Height          =   315
               Left            =   120
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":36DEC
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage8 
               Height          =   315
               Left            =   780
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":37ACA
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage7 
               Height          =   315
               Left            =   1380
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":387A8
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage6 
               Height          =   315
               Left            =   1980
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":39486
               Props           =   5
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Alinhamento Cabeçalho "
            Height          =   1095
            Left            =   4920
            TabIndex        =   52
            Top             =   2760
            Width           =   2415
            Begin VB.OptionButton Option12 
               Height          =   255
               Left            =   2040
               TabIndex        =   56
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option11 
               Height          =   255
               Left            =   1440
               TabIndex        =   55
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option10 
               Height          =   255
               Left            =   840
               TabIndex        =   54
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option9 
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Width           =   255
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage13 
               Height          =   315
               Left            =   1980
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3A164
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage12 
               Height          =   315
               Left            =   1380
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3AE42
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage11 
               Height          =   315
               Left            =   780
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3BB20
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage10 
               Height          =   315
               Left            =   120
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3C7FE
               Props           =   5
            End
         End
         Begin VB.Frame Frame21 
            Caption         =   "Alinhamento Rodapé "
            Height          =   1095
            Left            =   4920
            TabIndex        =   47
            Top             =   5160
            Width           =   2415
            Begin VB.OptionButton Option16 
               Height          =   255
               Left            =   2040
               TabIndex        =   51
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option15 
               Height          =   255
               Left            =   1440
               TabIndex        =   50
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option14 
               Height          =   255
               Left            =   840
               TabIndex        =   49
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option13 
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   255
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage17 
               Height          =   315
               Left            =   1980
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3D4DC
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage16 
               Height          =   315
               Left            =   1380
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3E1BA
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage15 
               Height          =   315
               Left            =   780
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3EE98
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage14 
               Height          =   315
               Left            =   120
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":3FB76
               Props           =   5
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Alinhamento Corpo "
            Height          =   1095
            Left            =   4920
            TabIndex        =   42
            Top             =   3960
            Width           =   2415
            Begin VB.OptionButton Option4 
               Height          =   255
               Left            =   2040
               TabIndex        =   46
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option3 
               Height          =   255
               Left            =   1440
               TabIndex        =   45
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option2 
               Height          =   255
               Left            =   840
               TabIndex        =   44
               Top             =   720
               Width           =   255
            End
            Begin VB.OptionButton Option1 
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   255
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage5 
               Height          =   315
               Left            =   1980
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":40854
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage4 
               Height          =   315
               Left            =   1380
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":41532
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage3 
               Height          =   315
               Left            =   780
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":42210
               Props           =   5
            End
            Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
               Height          =   315
               Left            =   120
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               Image           =   "frmCertificado.frx":42EEE
               Props           =   5
            End
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   3375
         End
         Begin VB.Frame Frame6 
            Caption         =   "Fundo "
            Height          =   3375
            Index           =   0
            Left            =   7440
            TabIndex        =   32
            Top             =   120
            Width           =   3375
            Begin MAESTRO.chameleonButton cmdCadastro 
               Height          =   615
               Index           =   13
               Left            =   720
               TabIndex        =   69
               Tag             =   "Excluir foto"
               ToolTipText     =   "Excluir foto"
               Top             =   2640
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
               MPTR            =   0
               MICON           =   "frmCertificado.frx":43BCC
               PICN            =   "frmCertificado.frx":43BE8
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
               TabIndex        =   68
               Tag             =   "Adicionar foto"
               ToolTipText     =   "Adicionar foto"
               Top             =   2640
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
               MPTR            =   0
               MICON           =   "frmCertificado.frx":448C2
               PICN            =   "frmCertificado.frx":448DE
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
               Height          =   2295
               Left            =   120
               ScaleHeight     =   2235
               ScaleWidth      =   3075
               TabIndex        =   33
               Top             =   240
               Width           =   3135
               Begin VB.Label Label59 
                  Alignment       =   2  'Center
                  Caption         =   "A Imagem não se encontra no local especificado"
                  Height          =   615
                  Left            =   840
                  TabIndex        =   34
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
                  Height          =   2055
                  Left            =   45
                  Top             =   75
                  Width           =   2985
                  _ExtentX        =   5265
                  _ExtentY        =   3625
                  Image           =   "frmCertificado.frx":455B8
               End
            End
            Begin MSComDlg.CommonDialog cdlFoto 
               Left            =   1680
               Top             =   2640
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Borda"
            Height          =   255
            Left            =   7560
            TabIndex        =   31
            Top             =   3840
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Logo"
            Height          =   255
            Left            =   7560
            TabIndex        =   30
            Top             =   4200
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Fundo"
            Height          =   255
            Left            =   7560
            TabIndex        =   29
            Top             =   4560
            Width           =   855
         End
         Begin VB.Frame Frame5 
            Caption         =   "Fonte Cabeçalho"
            Height          =   1095
            Left            =   120
            TabIndex        =   24
            Top             =   2760
            Width           =   4695
            Begin VB.Frame Frame7 
               Caption         =   "Nome"
               Height          =   735
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   3375
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   28
                  Top             =   240
                  Width           =   3135
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Tamanho"
               Height          =   735
               Left            =   3600
               TabIndex        =   25
               Top             =   240
               Width           =   975
               Begin VB.ComboBox Combo2 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   26
                  Top             =   240
                  Width           =   735
               End
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Fonte Corpo"
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   3960
            Width           =   4695
            Begin VB.Frame Frame12 
               Caption         =   "Nome"
               Height          =   735
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   3375
               Begin VB.ComboBox Combo3 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   23
                  Top             =   240
                  Width           =   3135
               End
            End
            Begin VB.Frame Frame13 
               Caption         =   "Tamanho"
               Height          =   735
               Left            =   3600
               TabIndex        =   20
               Top             =   240
               Width           =   975
               Begin VB.ComboBox Combo4 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   21
                  Top             =   240
                  Width           =   735
               End
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Fonte Rodapé"
            Height          =   1095
            Left            =   120
            TabIndex        =   14
            Top             =   5160
            Width           =   4695
            Begin VB.Frame Frame14 
               Caption         =   "Nome"
               Height          =   735
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   3375
               Begin VB.ComboBox Combo5 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   18
                  Top             =   240
                  Width           =   3135
               End
            End
            Begin VB.Frame Frame15 
               Caption         =   "Tamanho"
               Height          =   735
               Left            =   3600
               TabIndex        =   15
               Top             =   240
               Width           =   975
               Begin VB.ComboBox Combo6 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   16
                  Top             =   240
                  Width           =   735
               End
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Fonte Certificadora "
            Height          =   1095
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   4695
            Begin VB.Frame Frame16 
               Caption         =   "Nome"
               Height          =   735
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   3375
               Begin VB.ComboBox Combo7 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   13
                  Top             =   240
                  Width           =   3135
               End
            End
            Begin VB.Frame Frame17 
               Caption         =   "Tamanho"
               Height          =   735
               Left            =   3600
               TabIndex        =   10
               Top             =   240
               Width           =   975
               Begin VB.ComboBox Combo8 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   11
                  Top             =   240
                  Width           =   735
               End
            End
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   3720
            TabIndex        =   62
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   223019009
            CurrentDate     =   40742
         End
         Begin VB.Label Label1 
            Caption         =   "Título do responsável:"
            Height          =   255
            Left            =   3720
            TabIndex        =   65
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Data de emissão:"
            Height          =   255
            Left            =   3720
            TabIndex        =   63
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Empresa certificadora:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Título cabeçalho:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label53 
            BackColor       =   &H8000000C&
            Height          =   255
            Left            =   7440
            TabIndex        =   35
            Top             =   3480
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione os colaboradores os quais irão ser impressos os certificados "
         Height          =   6495
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   10455
         Begin VB.CheckBox Check4 
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   240
            Width           =   1455
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   5775
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   10186
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
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Instruções"
         Height          =   6375
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   10935
         Begin VB.TextBox Text6 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   6015
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   5
            Text            =   "frmCertificado.frx":455D0
            Top             =   240
            Width           =   10695
         End
      End
   End
End
Attribute VB_Name = "frmCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public rsEdText As New ADODB.Recordset
Public sqlEdText As String
Public Caminho1 As String
Public contaSelecionado As Integer

Private Sub Check4_Click()
    MarcaDesmarca ListView1
End Sub

Private Sub MarcaDesmarca(LV As ListView)
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    
    Y = LV.ListItems.Count
    For X = 1 To Y
        LV.ListItems(X).Selected = True
        If LV.ListItems.Item(X).Checked = True Then
            LV.ListItems.Item(X).Checked = False
        Else
            LV.ListItems.Item(X).Checked = True
        End If
    Next
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        contaSelecao
        If contaSelecionado = 0 Then
            Msgbox "Nenhum colaborador selecionado para emissão do certificado !", vbCritical, "SGCH"
        Else
            gravaDados
            Msgbox "Registro gravado com sucesso !", vbInformation, "SGCH"
        End If
    Case 1
        Unload Me
    Case 2
        contaSelecao
        If contaSelecionado = 0 Then
            Msgbox "Nenhum colaborador selecionado para emissão do certificado !", vbCritical, "SGCH"
        Else
            gravaDados
            FCRCertificado.Show 1
        End If
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
    End Select
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    listview_cabecalho
    preencheComboFontes
    preencheComboTamanhoFontes
    CompoeColabs
    restauraDados
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Nome", ListView1.Width / 4
    ListView1.ColumnHeaders.Add , , "Nota", ListView1.Width / 10
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub preencheComboFontes()
    'preenche a combo box com as fontes disponíveis
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(i)
        Combo3.AddItem Screen.Fonts(i)
        Combo5.AddItem Screen.Fonts(i)
        Combo7.AddItem Screen.Fonts(i)
    Next i
    Combo1.Text = "Arial"
    Combo3.Text = "Arial"
    Combo5.Text = "Arial"
    Combo7.Text = "Arial"
End Sub

Private Sub preencheComboTamanhoFontes()
    'preenche a combo box com os tamanhos das fontes
    Dim i As Integer
    For i = 8 To 24 Step 2
        Combo2.AddItem i
        Combo4.AddItem i
        Combo6.AddItem i
        Combo8.AddItem i
    Next i
    Combo2.ListIndex = 0
    Combo4.ListIndex = 0
    Combo6.ListIndex = 0
    Combo8.ListIndex = 0
End Sub

Private Sub CompoeColabs()
    Dim Y As Integer, X As Integer
    Dim ItemLst As ListItem
    Y = chamaForm.ListView1.ListItems.Count
    For X = 1 To Y
        chamaForm.ListView1.ListItems.Item(X).Selected = True
        Set ItemLst = ListView1.ListItems.Add(, , chamaForm.ListView1.SelectedItem.ListSubItems.Item(1))
        ItemLst.SubItems(1) = chamaForm.ListView1.SelectedItem.ListSubItems.Item(3)
    Next
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 1
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub restauraDados()
    sqlEdText = "Select * from tbConfCertificado where codcoligada = '" & vCodcoligada & "'"
    rsEdText.Open sqlEdText, cnBanco, adOpenKeyset, adLockOptimistic
    
    If Not rsEdText.EOF Then
    Text5.Text = rsEdText.Fields(1)
    Text2.Text = rsEdText.Fields(2)
    Text4.Text = rsEdText.Fields(3)
    If Not IsNull(rsEdText.Fields(5)) And rsEdText.Fields(5) = "S" Then Check1.Value = 1 Else Check1.Value = 0
    If Not IsNull(rsEdText.Fields(4)) And rsEdText.Fields(4) = "S" Then Check2.Value = 1 Else Check2.Value = 0
    If Not IsNull(rsEdText.Fields(6)) And rsEdText.Fields(6) = "S" Then Check3.Value = 1 Else Check3.Value = 0
    If rsEdText.Fields(7) <> "Null" Then
        On Error GoTo TrataErro1
        Label53.Caption = rsEdText.Fields(7)
        aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    End If
    'Fonte
    Combo1.Text = rsEdText.Fields(8)
    Combo3.Text = rsEdText.Fields(9)
    Combo5.Text = rsEdText.Fields(10)
    Combo7.Text = rsEdText.Fields(11)
    'Tamanho Fonte
    Combo2.Text = rsEdText.Fields(12)
    Combo4.Text = rsEdText.Fields(13)
    Combo6.Text = rsEdText.Fields(14)
    Combo8.Text = rsEdText.Fields(15)
    'Alinhamento Fonte Corpo
    If rsEdText.Fields(16) = 1 Then Option1.Value = True
    If rsEdText.Fields(16) = 2 Then Option2.Value = True
    If rsEdText.Fields(16) = 3 Then Option3.Value = True
    If rsEdText.Fields(16) = 4 Then Option4.Value = True
    
    'Alinhamento Fonte Rodapé
    If rsEdText.Fields(17) = 1 Then Option13.Value = True
    If rsEdText.Fields(17) = 2 Then Option14.Value = True
    If rsEdText.Fields(17) = 3 Then Option15.Value = True
    If rsEdText.Fields(17) = 4 Then Option16.Value = True
    
    'Alinhamento Fonte Cabeçalho
    If rsEdText.Fields(18) = 1 Then Option9.Value = True
    If rsEdText.Fields(18) = 2 Then Option10.Value = True
    If rsEdText.Fields(18) = 3 Then Option11.Value = True
    If rsEdText.Fields(18) = 4 Then Option12.Value = True
    
    'Alinhamento Fonte Certificadora
    If rsEdText.Fields(19) = 1 Then Option8.Value = True
    If rsEdText.Fields(19) = 2 Then Option7.Value = True
    If rsEdText.Fields(19) = 3 Then Option6.Value = True
    If rsEdText.Fields(19) = 4 Then Option5.Value = True
    DTPicker1.Value = rsEdText.Fields(25)
    If Not IsNull(rsEdText.Fields(27)) Then Text1.Text = rsEdText.Fields(27)
    End If
    rsEdText.Close
    Set rsEdText = Nothing
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub gravaDados()
    Dim idCertificado As Integer
    Dim Y As Integer, X As Integer
    sqlEdText = "Select * from tbConfCertificado where codcoligada = '" & vCodcoligada & "'"
    rsEdText.Open sqlEdText, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsEdText.RecordCount = 0 Then rsEdText.AddNew 'Identificador
    rsEdText.Fields(1) = Text5.Text 'Texto do certificado
    rsEdText.Fields(2) = Text2.Text 'Empresa certificadora
    rsEdText.Fields(3) = Text4.Text 'Titulo
    If Check2.Value = 1 Then rsEdText.Fields(4) = "S" Else rsEdText.Fields(4) = "N" 'Usar logo?
    If Check1.Value = 1 Then rsEdText.Fields(5) = "S" Else rsEdText.Fields(5) = "N" 'Usar borda?
    If Check3.Value = 1 Then rsEdText.Fields(6) = "S" Else rsEdText.Fields(6) = "N" 'Usar fundo?
    rsEdText.Fields(7) = Label53.Caption 'caminho do fundo
    rsEdText.Fields(26) = frmSplash.aicAlphaImage1.Tag 'caminho da logo

    Text5 = rsEdText.Fields(1)
    
    rsEdText.Fields(8) = Combo1.Text 'Fonte do cabeçalho
    rsEdText.Fields(9) = Combo3.Text
    rsEdText.Fields(10) = Combo5.Text
    rsEdText.Fields(11) = Combo7.Text
    
    'Fonte
    rsEdText.Fields(8) = Combo1.Text
    rsEdText.Fields(9) = Combo3.Text
    rsEdText.Fields(10) = Combo5.Text
    rsEdText.Fields(11) = Combo7.Text
    'Tamanho Fonte
    rsEdText.Fields(12) = Combo2.Text
    rsEdText.Fields(13) = Combo4.Text
    rsEdText.Fields(14) = Combo6.Text
    rsEdText.Fields(15) = Combo8.Text
    'Alinhamento da fonte do Corpo
    rsEdText.Fields(16) = 0
    If Option1.Value = True Then rsEdText.Fields(16) = 1
    If Option2.Value = True Then rsEdText.Fields(16) = 2
    If Option3.Value = True Then rsEdText.Fields(16) = 3
    If Option4.Value = True Then rsEdText.Fields(16) = 4
    
    'Alinhamento da fonte do Rodapé
    rsEdText.Fields(17) = 0
    If Option13.Value = True Then rsEdText.Fields(17) = 1
    If Option14.Value = True Then rsEdText.Fields(17) = 2
    If Option15.Value = True Then rsEdText.Fields(17) = 3
    If Option16.Value = True Then rsEdText.Fields(17) = 4
    
    'Alinhamento da fonte do Cabeçalho
    rsEdText.Fields(18) = 0
    If Option9.Value = True Then rsEdText.Fields(18) = 1
    If Option10.Value = True Then rsEdText.Fields(18) = 2
    If Option11.Value = True Then rsEdText.Fields(18) = 3
    If Option12.Value = True Then rsEdText.Fields(18) = 4
    
    'Alinhamento Fonte Certificadora
    rsEdText.Fields(19) = 0
    If Option8.Value = True Then rsEdText.Fields(19) = 1
    If Option7.Value = True Then rsEdText.Fields(19) = 2
    If Option6.Value = True Then rsEdText.Fields(19) = 3
    If Option5.Value = True Then rsEdText.Fields(19) = 4
    
    
    rsEdText.Fields(20) = chamaForm.Text6
    rsEdText.Fields(21) = chamaForm.DTPicker2
    rsEdText.Fields(22) = chamaForm.DTPicker3
    rsEdText.Fields(23) = chamaForm.Text9
    chamaForm.ListView2.ListItems.Item(1).Selected = True
    rsEdText.Fields(24) = chamaForm.ListView2.SelectedItem.ListSubItems.Item(3)
    rsEdText.Fields(25) = DTPicker1.Value
    rsEdText.Fields(27) = Text1.Text 'Titulo responsável
    rsEdText.Fields(28) = vCodcoligada 'Codigo da coligada
    rsEdText.Update
    idCertificado = rsEdText.Fields(0)
    rsEdText.Close
    Set rsEdText = Nothing

    
    sqlEdText = "Delete from tbColabCertificado where codcoligada ='" & vCodcoligada & "'"
    rsEdText.Open sqlEdText, cnBanco
      
    sqlEdText = "select * from tbColabCertificado where codcoligada ='" & vCodcoligada & "'"
    rsEdText.Open sqlEdText, cnBanco, adOpenKeyset, adLockOptimistic
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            rsEdText.AddNew
            rsEdText.Fields(0) = idCertificado
            rsEdText.Fields(1) = ListView1.ListItems.Item(X)
            rsEdText.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(1)
            rsEdText.Fields(3) = vCodcoligada 'Codigo da coligada
        End If
    Next
    rsEdText.Update
    rsEdText.Close
    Set rsEdText = Nothing
End Sub

Private Sub contaSelecao()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    contaSelecionado = 0
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.ListItems.Item(X).Checked = True Then
            contaSelecionado = contaSelecionado + 1
        End If
    Next
End Sub
