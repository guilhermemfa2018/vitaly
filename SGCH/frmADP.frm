VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADP - Avaliação de Desempenho Profissional"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11985
   Icon            =   "frmADP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      TabIndex        =   72
      Top             =   8640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   10800
      TabIndex        =   69
      Top             =   8760
      Width           =   1095
      Begin VB.CheckBox Check7 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Tag             =   "Status do curso/treinamento"
         ToolTipText     =   "Status do curso/treinamento"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Nota "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8040
      TabIndex        =   59
      Top             =   120
      Width           =   1815
      Begin SGCH.chameleonButton chameleonButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "Calcular"
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
         MICON           =   "frmADP.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "NOTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   55
      Top             =   4440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Avaliações"
      TabPicture(0)   =   "frmADP.frx":0CE6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label15"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdINTD(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdINTD(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdINTD(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdINTD(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtADP(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtADP(11)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Combo1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtADP(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtADP(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdINTD(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Observações"
      TabPicture(1)   =   "frmADP.frx":0D02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "txtADP(13)"
      Tab(1).ControlCount=   2
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   615
         Index           =   2
         Left            =   10920
         TabIndex        =   77
         Tag             =   "Definir MODELO de Avaliação de Desempenho Profissional"
         ToolTipText     =   "Definir MODELO de Avaliação de Desempenho Profissional"
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
         MICON           =   "frmADP.frx":0D1E
         PICN            =   "frmADP.frx":0D3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   16
         Left            =   9240
         TabIndex        =   73
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   615
         Index           =   15
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   6375
      End
      Begin VB.Frame Frame8 
         Caption         =   "Indicações de treinamentos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   -67560
         TabIndex        =   65
         Top             =   360
         Width           =   4215
         Begin VB.Frame Frame10 
            Caption         =   "Modalidade "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   120
            TabIndex        =   67
            Top             =   960
            Width           =   3975
            Begin VB.Frame Frame11 
               Caption         =   "Modalidade outro "
               Enabled         =   0   'False
               Height          =   615
               Left            =   120
               TabIndex        =   68
               Top             =   1800
               Width           =   3735
               Begin VB.TextBox txtADP 
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   14
                  Left            =   120
                  TabIndex        =   33
                  Top             =   240
                  Width           =   3495
               End
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Outros"
               Height          =   255
               Index           =   5
               Left            =   1560
               TabIndex        =   32
               Top             =   1320
               Width           =   2055
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Segurança"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   31
               Top             =   1320
               Width           =   1335
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Funcional"
               Height          =   255
               Index           =   3
               Left            =   1560
               TabIndex        =   30
               Top             =   840
               Width           =   1215
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Gerencial"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   840
               Width           =   1335
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Administrativo"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   28
               Top             =   360
               Width           =   1695
            End
            Begin VB.CheckBox chkADP 
               Caption         =   "Informática"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Tipo "
            Height          =   615
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   3975
            Begin VB.CheckBox Check2 
               Caption         =   "Externo"
               Height          =   255
               Left            =   1560
               TabIndex        =   75
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Interno"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.TextBox txtADP 
         Height          =   3615
         Index           =   13
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   7215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmADP.frx":1A14
         Left            =   9240
         List            =   "frmADP.frx":1A21
         TabIndex        =   22
         Text            =   "Institucional"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   17
         Tag             =   "Código do treinamento"
         ToolTipText     =   "Código do treinamento"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   18
         Tag             =   "Nome do treinamento"
         ToolTipText     =   "Nome do treinamento"
         Top             =   720
         Width           =   6375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2235
         Left            =   120
         TabIndex        =   25
         Tag             =   "Características avaliadas na ADP"
         ToolTipText     =   "Características avaliadas na ADP"
         Top             =   1800
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3942
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
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   615
         Index           =   5
         Left            =   1320
         TabIndex        =   24
         Tag             =   "Excluir avaliação"
         ToolTipText     =   "Excluir avaliação"
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
         MICON           =   "frmADP.frx":1A4B
         PICN            =   "frmADP.frx":1A67
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   615
         Index           =   7
         Left            =   720
         TabIndex        =   23
         Tag             =   "Nova avaliação"
         ToolTipText     =   "Nova avaliação"
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
         MICON           =   "frmADP.frx":2741
         PICN            =   "frmADP.frx":275D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Tag             =   "Incluir Avaliação"
         ToolTipText     =   "Incluir Avaliação"
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
         MICON           =   "frmADP.frx":3437
         PICN            =   "frmADP.frx":3453
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   255
         Index           =   4
         Left            =   8640
         TabIndex        =   21
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
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
         MICON           =   "frmADP.frx":412D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label16 
         Caption         =   "Dimensão:"
         Height          =   255
         Left            =   9240
         TabIndex        =   63
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Avaliação:"
         Height          =   255
         Left            =   2160
         TabIndex        =   62
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Absenteísmo/Responsável pela informação "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5160
      TabIndex        =   52
      Top             =   2640
      Width           =   6735
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Height          =   1335
         Left            =   1680
         TabIndex        =   56
         Top             =   240
         Width           =   4935
         Begin VB.TextBox txtADP 
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   13
            Tag             =   "Registro do colaborador em treinamento"
            ToolTipText     =   "Registro do colaborador em treinamento"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtADP 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   14
            Tag             =   "Nome do colaborador em treinamento"
            ToolTipText     =   "Nome do colaborador em treinamento"
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtADP 
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   4695
         End
         Begin SGCH.chameleonButton cmdINTD 
            Height          =   255
            Index           =   1
            Left            =   4440
            TabIndex        =   15
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "..."
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmADP.frx":4149
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label14 
            Caption         =   "Registro:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Nome:"
            Height          =   255
            Left            =   1320
            TabIndex        =   57
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Label Label13 
         Caption         =   "Atrasos no ano:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Ausências no ano:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Responsável pela avaliação "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   48
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Tag             =   "Registro do colaborador em treinamento"
         ToolTipText     =   "Registro do colaborador em treinamento"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   8
         Tag             =   "Nome do colaborador em treinamento"
         ToolTipText     =   "Nome do colaborador em treinamento"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   4695
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   9
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "..."
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmADP.frx":4165
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label12 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Matriz/Cargo"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Foto "
      Height          =   2415
      Index           =   0
      Left            =   9960
      TabIndex        =   37
      Top             =   120
      Width           =   1935
      Begin VB.PictureBox Picture2 
         Height          =   2055
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   1635
         TabIndex        =   38
         Top             =   240
         Width           =   1695
         Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
            Height          =   2175
            Left            =   0
            Top             =   -120
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   3836
            Image           =   "frmADP.frx":4181
         End
      End
      Begin MSComDlg.CommonDialog cdlFoto 
         Left            =   1080
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da ADP "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtADP 
         Height          =   285
         Index           =   17
         Left            =   3480
         TabIndex        =   79
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   115474433
         CurrentDate     =   40784
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   285
         Left            =   6360
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   115474433
         CurrentDate     =   40784
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4920
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   115474433
         CurrentDate     =   40784
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo/Período "
         Height          =   615
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Width           =   7575
         Begin VB.Label Label2 
            Caption         =   "Tipo/Período"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   7335
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   115474433
         CurrentDate     =   40784
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Registro do colaborador em treinamento"
         ToolTipText     =   "Registro do colaborador em treinamento"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Tag             =   "Nome do colaborador em treinamento"
         ToolTipText     =   "Nome do colaborador em treinamento"
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtADP 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label18 
         Caption         =   "ID Colaborador"
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
         Left            =   3240
         TabIndex        =   78
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Data devolução:"
         Height          =   255
         Left            =   6360
         TabIndex        =   47
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Data vencimento:"
         Height          =   255
         Left            =   6360
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data avaliação:"
         Height          =   255
         Left            =   4920
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Data admissão:"
         Height          =   255
         Left            =   4920
         TabIndex        =   42
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   1680
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Matriz 
         Caption         =   "Matriz/Cargo:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   1095
      End
   End
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   12
      Left            =   720
      TabIndex        =   35
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
      MICON           =   "frmADP.frx":4199
      PICN            =   "frmADP.frx":41B5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   11
      Left            =   120
      TabIndex        =   34
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
      MICON           =   "frmADP.frx":4E8F
      PICN            =   "frmADP.frx":4EAB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   9
      Left            =   10080
      TabIndex        =   76
      Tag             =   "Concluir ADP - Avaliação de Desempenho Profissional"
      ToolTipText     =   "Concluir ADP - Avaliação de Desempenho Profissional"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmADP.frx":549C
      PICN            =   "frmADP.frx":54B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label53 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   1440
      TabIndex        =   71
      Top             =   8640
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "frmADP"
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
'Acima - usado poder editar o listview --------------------

Private rsADP As New ADODB.Recordset
Private sqlADP As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsTreiADP As New ADODB.Recordset
Private sqlTreiADP As String

Private rsLocal As New ADODB.Recordset

Private Sub chameleonButton1_Click()
    calculaNotaADP
End Sub

Private Sub chkADP_Click(Index As Integer)
    Select Case Index
    Case 5
        If chkADP(5).Value = 1 Then
            Frame11.Enabled = True
            txtADP(14).Enabled = True
        Else
            Frame11.Enabled = False
            txtADP(14).Enabled = False
        End If
    End Select
End Sub

Private Sub cmdINTD_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridColaborador 0
        CarregaColaborador 3
    Case 1
        ChamaGridColaborador 1
        CarregaColaborador 8
    Case 2
        frmADPModelo.Show 1
    Case 4
        ChamaGridAvaliacaoADP
        CarregaAvaliacao
    Case 5
        If MsgBox("Deseja EXCLUIR Avaliação da ADP?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            ExcluirItemLV ListView1
            LimpaControlesAvaliacao
        End If
    Case 7
        LimpaControlesAvaliacao
    Case 8
        IncluirAvaliacao
        LimpaControlesAvaliacao
    Case 11
        If MsgBox("Deseja salvar os dados da ADP?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            calculaNotaADP
            gravaDadosADP
'            gravaLog "Código req: " & txtCadReq(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
            Unload Me
        End If
    Case 9
        'CONCLUSÃO DA ADP
        If MsgBox("Deseja iniciar a conclusão da ADP - Avaliação de Desempenho Profissional?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            calculaNotaADP
            gravaDadosADP
            If checkListADP = False Then Exit Sub
            'gravaLog "Código PS: " & txtProcesso(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
            
            carregaADP txtADP(17)
            'carregaADP "TODOS"
            Unload Me
        End If
        Unload Me
    Case 12
        If MsgBox("Deseja sair dessa Avaliação de Desempenho Profissional?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = "0"
            Unload Me
            Set frmADP = Nothing
        End If
    End Select
End Sub

Private Sub Form_Activate()
    If MeuLV.ListView1.ListItems.Count = 0 Then Unload Me
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    DTPicker2 = Date
    listview_cabecalho
    CompoeControles
    calculaNotaADP
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Avaliação", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Peso", ListView1.Width / 10000
    ListView1.ColumnHeaders.Add , , "Nota", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Dimensão", ListView1.Width / 6
    ListView1.ColumnHeaders.Add , , "Descrição", ListView1.Width / 10000
                
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub LimpaControlesAvaliacao()
    Dim X As Integer
    txtADP(11) = ""
    txtADP(12) = ""
    txtADP(15) = ""
    txtADP(11).SetFocus
End Sub

Private Sub IncluirAvaliacao()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaAvaliacao = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtADP(11) Then
                Me.txtADP(8) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtADP(12)
                ListView1.SelectedItem.ListSubItems.Item(2) = txtADP(16)
                If ListView1.SelectedItem.ListSubItems.Item(3) = "" Then ListView1.SelectedItem.ListSubItems.Item(3) = "0"
                ListView1.SelectedItem.ListSubItems.Item(4) = Combo1
                ListView1.SelectedItem.ListSubItems.Item(5) = txtADP(15)
                Y = ListView1.ListItems.Count
                
                Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
                Me.ListView1.SelectedItem.ListSubItems.Item(3).Bold = True

                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtADP(11))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtADP(11))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtADP(12)
    ItemLst.SubItems(2) = txtADP(16)
    If ItemLst.SubItems(3) = "" Then ItemLst.SubItems(3) = "0"
    ItemLst.SubItems(4) = Combo1
    ItemLst.SubItems(5) = txtADP(15)
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    ItemLst.ListSubItems(3).Bold = True
    txtADP(8).SetFocus
End Sub

Private Function ValidaAvaliacao()
    ValidaAvaliacao = False
    If txtADP(11).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtADP(11).Tag, vbInformation, "Atenção"
        Me.txtADP(11).SetFocus
        Exit Function
    End If
    ValidaAvaliacao = True
End Function

Private Sub CompoeControles()
On Error GoTo TrataErro1
    Dim vidADP As Integer
    txtADP(0).Text = varGlobal
    txtADP(1).Text = MeuLV.ListView1.SelectedItem.ListSubItems.Item(1)
    Label2 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) & " - " & MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) & " DIAS"
    If MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) <> "-" Then
        DTPicker2 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
    Else
        DTPicker2 = Date
    End If
    DTPicker3 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
    DTPicker4 = MeuLV.ListView1.SelectedItem.ListSubItems.Item(5)
    
    sqlADP = "select a.codcolaborador,a.nota,b.cpf,c.codmatriz,e.nomecargo,c.data,b.foto,a.codrespADP,a.nomerespADP,a.ausenciaano,a.atrasoano,a.codrespABS,a.nomerespABS,a.observacao,a.indicacaotipo,a.indicacaomod1,a.indicacaomod2,a.indicacaomod3,a.indicacaomod4," & _
             "a.indicacaomod5,a.indicacaomod6,a.indicacaooutros,a.statusimpressao,a.statusavaliacao,a.ativo,a.id,a.cargoADP from tblistaadp as a inner join tbcolaboradores as b on a.codcoligada = '" & vCodcoligada & "' and a.codcolaborador = b.id inner join tbcolaboradoreshist as c " & _
             "on b.cpf = c.cpf and c.ativo = 'S' inner join tbmatriz as d on c.codmatriz = d.codmatriz inner join tbcargos as e on d.codcargo = e.codcargo where a.id = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(11)) & "'"
    
    rsADP.Open sqlADP, cnBanco, adOpenKeyset, adLockReadOnly
    Label18 = rsADP.Fields(0)
    DTPicker1 = Format(rsADP.Fields(5), "dd/mm/yyyy")
    If Not IsNull(rsADP.Fields(26)) Then txtADP(2) = rsADP.Fields(26) Else txtADP(2) = Format(rsADP.Fields(3), "000000") & " - " & rsADP.Fields(4)
    If Not IsNull(rsADP.Fields(1)) Then Label17 = rsADP.Fields(1) Else Label17 = "-"
    Label53 = rsADP.Fields(6)
    aicAlphaImage1.LoadImage_FromFile (Label53.Caption)
    
    If Not IsNull(rsADP.Fields(2)) Then txtADP(17) = rsADP.Fields(2)
    If Not IsNull(rsADP.Fields(7)) Then txtADP(3) = rsADP.Fields(7)
    If Not IsNull(rsADP.Fields(9)) Then txtADP(6) = rsADP.Fields(9)
    If Not IsNull(rsADP.Fields(10)) Then txtADP(7) = rsADP.Fields(10)
    If Not IsNull(rsADP.Fields(11)) Then txtADP(8) = rsADP.Fields(11)
    If Not IsNull(rsADP.Fields(13)) Then txtADP(13) = rsADP.Fields(13)
    
    If rsADP.Fields(14) = 1 Then
        Check1.Value = 1
        Check2.Value = 0
    ElseIf rsADP.Fields(14) = 2 Then
        Check1.Value = 0
        Check2.Value = 1
    ElseIf rsADP.Fields(14) = 3 Then
        Check1.Value = 1
        Check2.Value = 1
    End If
    If rsADP.Fields(15) = 1 Then chkADP(0).Value = 1 Else chkADP(0).Value = 0
    If rsADP.Fields(16) = 1 Then chkADP(1).Value = 1 Else chkADP(1).Value = 0
    If rsADP.Fields(17) = 1 Then chkADP(2).Value = 1 Else chkADP(2).Value = 0
    If rsADP.Fields(18) = 1 Then chkADP(3).Value = 1 Else chkADP(3).Value = 0
    If rsADP.Fields(19) = 1 Then chkADP(4).Value = 1 Else chkADP(4).Value = 0
    If rsADP.Fields(20) = 1 Then chkADP(5).Value = 1 Else chkADP(5).Value = 0
    If Not IsNull(rsADP.Fields(21)) Then txtADP(14) = rsADP.Fields(21)
    If Not IsNull(rsADP.Fields(1)) Then Label17 = rsADP.Fields(1)
    CarregaColaborador 3
    CarregaColaborador 8
    'txtADP(4) = rsADP.Fields(9)
    If Not IsNull(rsADP.Fields(12)) Then txtADP(9) = rsADP.Fields(12)
    If rsADP.Fields(24) = "S" Then 'ativo
        Check7.Value = 1
    Else
        Check7.Value = 0
    End If
    vidADP = rsADP.Fields(25)
    rsADP.Close
    Set rsADP = Nothing
    
    'Compoe Listview1
    Dim ItemLst As ListItem
    sqlADP = "select a.codavaliacao,b.nomeavaliacao,b.peso,a.nota,a.dimensao,b.descricao from tbListaADPItens as a inner join tbavaliacao as b on a.codcoligada = '" & vCodcoligada & "' and a.codavaliacao = b.codavaliacao where a.idADP = '" & vidADP & "'"
    rsADP.Open sqlADP, cnBanco, adOpenKeyset, adLockReadOnly
    ListView1.ListItems.Clear
    While Not rsADP.EOF
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsADP.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsADP.Fields(1)
        ItemLst.SubItems(2) = "" & rsADP.Fields(2)
        ItemLst.SubItems(3) = "" & rsADP.Fields(3)
        ItemLst.SubItems(4) = "" & rsADP.Fields(4)
        ItemLst.SubItems(5) = "" & rsADP.Fields(5)
        ItemLst.ListSubItems(3).Bold = True
        rsADP.MoveNext
    Wend
    rsADP.Close
    Set rsADP = Nothing
    Me.ListView1.ColumnHeaders(4).Alignment = lvwColumnRight
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending
    If MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = "Concluido" Then BloqueiaControles
    Exit Sub
TrataErro1:
    Resume Next
End Sub

Private Sub CarregaColaborador(indice As Integer)
    Dim X As Integer
    SqlColaborador = "select a.codcolaborador,a.nomecolaborador,c.codmatriz,f.nomecargo,a.id from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c  on b.codmatriz = c.codmatriz inner join " & _
                     "tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor inner join tbcargos as f on c.codcargo = f.codcargo where a.ativo = 'S' and a.codcolaborador = '" & txtADP(indice) & "'"
    rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockReadOnly
    If rsColaborador.RecordCount <= 0 Then
        If indice = 1 Then
            If txtADP(3).Text <> "000000" And txtADP(3).Text <> "" Then MsgBox "Colaborador não cadastrado", vbInformation, "SGCH"
            txtADP(4) = ""
            txtADP(5) = ""
        Else
            If txtADP(8).Text <> "000000" And txtADP(8).Text <> "" Then MsgBox "Colaborador não cadastrado", vbInformation, "SGCH"
            txtADP(9) = ""
            txtADP(10) = ""
        End If
    Else
        If indice = 3 Then
            txtADP(3).Text = rsColaborador.Fields(0)
            txtADP(4).Text = rsColaborador.Fields(1)
            txtADP(5).Text = Format(rsColaborador.Fields(2), "000000") & " - " & rsColaborador.Fields(3)
        Else
            txtADP(8).Text = rsColaborador.Fields(0)
            txtADP(9).Text = rsColaborador.Fields(1)
            txtADP(10).Text = Format(rsColaborador.Fields(2), "000000") & " - " & rsColaborador.Fields(3)
        End If
    End If
    rsColaborador.Close
    Set rsColaborador = Nothing
End Sub

Private Sub CarregaAvaliacao()
    Dim X As Integer
    sqlTreiADP = "select a.codavaliacao,a.nomeavaliacao,a.descricao,a.tipo,a.peso from tbavaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.tipo = 'AD' and a.codavaliacao = '" & Val(txtADP(11)) & "' and a.ativo = 'S' order by a.codavaliacao"
    rsTreiADP.Open sqlTreiADP, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsTreiADP.EOF Then rsTreiADP.MoveFirst
    If rsTreiADP.EOF Then
        txtADP(11).Text = Format(txtADP(11), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            MsgBox "Avaliação não cadastrada", vbInformation, "SGCH"
            txtADP(12) = ""
            txtADP(15) = ""
            txtADP(16) = ""
        End If
    Else
        txtADP(11).Text = Format(rsTreiADP.Fields(0), "000000") & ""
        txtADP(12).Text = rsTreiADP.Fields(1)
        txtADP(15).Text = rsTreiADP.Fields(2)
        txtADP(16).Text = rsTreiADP.Fields(4)
    End If
    rsTreiADP.Close
    Set rsTreiADP = Nothing
End Sub

Private Sub txtADP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 3
        End If
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 8
        End If
    Case 11
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaAvaliacao
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub ChamaGridColaborador(indice As Integer)
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and tipo = 'colaborador' and ativo = 'S' order by nomecolaborador"
    procnom = "nomecolaborador"
    campo = 3
    Campo1 = 1
    Load F
    F.Caption = "Pesquisa de Colaborador"
    Pesquisa = frmADP.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If indice = 0 Then
                txtADP(3).Text = rsLocal.Fields(1)
            Else
                txtADP(8).Text = rsLocal.Fields(1)
            End If
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub ChamaGridAvaliacaoADP()
'On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select a.codavaliacao,a.nomeavaliacao,a.descricao,a.tipo from tbavaliacao as a where a.codcoligada = '" & vCodcoligada & "' and a.tipo = 'AD' and a.ativo = 'S' order by a.codavaliacao"
    procnom = "codavaliacao"
    procnom1 = "nomeavaliacao"
    campo = 1
    Campo1 = 0
    Pesquisa = frmADP.Tag
    Load F
    F.Caption = "Pesquisa de avaliação"
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "codavaliacao=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtADP(11).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub gravaDadosADP()
'On Error GoTo TrataErro
    'If ValidaCampos = False Then Exit Sub
    Dim rsSalvarADP As New ADODB.Recordset
    Dim SqlSalvarADP As String
    Dim rsSalvarADPItens As New ADODB.Recordset
    Dim SqlSalvarADPItens As String
    Dim vidADP As Integer
    
    cnBanco.BeginTrans
    SqlSalvarADP = "Select * from tblistaADP as a where a.codcoligada = '" & vCodcoligada & "' and a.id = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(11)) & "'"
    rsSalvarADP.Open SqlSalvarADP, cnBanco, adOpenKeyset, adLockOptimistic
    
    'Gravar na tabela tbListaADP
    vidADP = rsSalvarADP.Fields(0)
    rsSalvarADP.Fields(4) = DTPicker2 'data de avaliacaçao
    rsSalvarADP.Fields(6) = DTPicker4 'data devolução
    rsSalvarADP.Fields(7) = txtADP(3) 'codigo do responsavel pela avaliação
    rsSalvarADP.Fields(8) = txtADP(4) 'nome do responsavel pela avaliação
    If txtADP(6) <> "" Then rsSalvarADP.Fields(9) = txtADP(6) Else rsSalvarADP.Fields(9) = 0 'numero de ausencias no ano
    If txtADP(7) <> "" Then rsSalvarADP.Fields(10) = txtADP(7) Else rsSalvarADP.Fields(10) = 0 'numero de atrasos no ano
    rsSalvarADP.Fields(11) = txtADP(8) 'codigo do responsavel pelo ABS
    rsSalvarADP.Fields(12) = txtADP(9) 'nome do responsavel pelo ABS
    rsSalvarADP.Fields(13) = txtADP(13) 'observacao
    
    If Check1.Value = 0 And Check2.Value = 0 Then rsSalvarADP.Fields(14) = 0
    If Check1.Value = 1 And Check2.Value = 0 Then rsSalvarADP.Fields(14) = 1
    If Check1.Value = 0 And Check2.Value = 1 Then rsSalvarADP.Fields(14) = 2
    If Check1.Value = 1 And Check2.Value = 1 Then rsSalvarADP.Fields(14) = 3
    
    If chkADP(0).Value = 1 Then rsSalvarADP.Fields(15) = 1 Else rsSalvarADP.Fields(15) = 0
    If chkADP(1).Value = 1 Then rsSalvarADP.Fields(16) = 1 Else rsSalvarADP.Fields(16) = 0
    If chkADP(2).Value = 1 Then rsSalvarADP.Fields(17) = 1 Else rsSalvarADP.Fields(17) = 0
    If chkADP(3).Value = 1 Then rsSalvarADP.Fields(18) = 1 Else rsSalvarADP.Fields(18) = 0
    If chkADP(4).Value = 1 Then rsSalvarADP.Fields(19) = 1 Else rsSalvarADP.Fields(19) = 0
    If chkADP(5).Value = 1 Then
        rsSalvarADP.Fields(20) = 1
        rsSalvarADP.Fields(21) = txtADP(14) 'indicacao outros
    Else
        rsSalvarADP.Fields(20) = 0
    End If
    
    rsSalvarADP.Fields(23) = "Avaliando" 'Status da ADP
    rsSalvarADP.Fields(25) = RemoveMask(Label17) 'Media Geral da ADP
    rsSalvarADP.Fields(26) = vCodcoligada 'Codigo da coligada
    rsSalvarADP.Fields(27) = txtADP(2).Text 'Cargo atual que esta sendo avaliado na ADP do colaborador
    If Not rsSalvarADP.EOF Then rsSalvarADP.Update
    rsSalvarADP.Close
    Set rsSalvarADP = Nothing
    
    'Gravar na tabela tbListaADPItens
    SqlSalvarADPItens = "Delete from tbListaADPItens where tbListaADPItens.codcoligada = '" & vCodcoligada & "' and tbListaADPItens.idADP = '" & vidADP & "'"
    rsSalvarADPItens.Open SqlSalvarADPItens, cnBanco
    
    SqlSalvarADPItens = "Select * from tbListaADPItens"
    rsSalvarADPItens.Open SqlSalvarADPItens, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvarADPItens.AddNew
        rsSalvarADPItens.Fields(0) = vidADP 'Identificador da avaliação
        rsSalvarADPItens.Fields(1) = ListView1.ListItems.Item(X) 'código do item da avaliação
        rsSalvarADPItens.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(3) 'nota do item avaliado
        rsSalvarADPItens.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(4) ' dimensão avaliada
        rsSalvarADPItens.Fields(4) = vCodcoligada ' Código da coligada
    Next
    If Not rsSalvarADPItens.EOF Then rsSalvarADPItens.Update
    rsSalvarADPItens.Close
    Set rsSalvarADPItens = Nothing
    
    cnBanco.CommitTrans
    
    MsgBox "Os dados da ADP foram salvos com sucesso", vbInformation, "SGCH"
    AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub calculaNotaADP()
    Dim rsABS As New ADODB.Recordset
    Dim sqlABS As String
    Dim vValor1 As Double, vPontosAusencia As Double, vPontosAtraso As Double
    Dim X As Integer, Y As Integer
    
    If ListView1.ListItems.Count > 0 Then
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If Val(ListView1.SelectedItem.ListSubItems.Item(3)) <> 0 Then
                vValor1 = vValor1 + Val(ListView1.SelectedItem.ListSubItems.Item(3))
            End If
        Next
        vValor1 = vValor1
    Else
        vValor1 = 0
    End If
    
    'Acha pontos de ausencia na tabela tbABS
    sqlABS = "select a.pontos from tbABS as a where a.codcoligada = '" & vCodcoligada & "' and tipo = 'Ausência' and '" & Val(txtADP(6)) & "'>= oc1 and '" & Val(txtADP(6)) & "' <= oc2"
    rsABS.Open sqlABS, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsABS.EOF Then
        vPontosAusencia = rsABS.Fields(0)
    Else
        vPontosAusencia = 0
    End If
    rsABS.Close
    Set rsABS = Nothing
    
    'Acha pontos de atraso na tabela tbABS
    sqlABS = "select a.pontos from tbABS as a where a.codcoligada = '" & vCodcoligada & "' and tipo = 'Atraso' and '" & Val(txtADP(7)) & "'>= oc1 and '" & Val(txtADP(7)) & "' <= oc2"
    rsABS.Open sqlABS, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsABS.EOF Then
        vPontosAtraso = rsABS.Fields(0)
    Else
        vPontosAtraso = 0
    End If
    rsABS.Close
    Set rsABS = Nothing
    
    If vValor1 > 0 Then
        vMediaGeral = (vValor1 - (vPontosAusencia + vPontosAtraso)) / Y
    Else
        vMediaGeral = 0
    End If
    Label17 = Format(vMediaGeral, "#,##0.00;(#,##0.00)") & " %"
    
    If vMediaGeral < vAprovadoRest Then
        Label17.ForeColor = &HC0&
    End If
    If vMediaGeral < MediaGlobal And vMediaGeral >= vAprovadoRest Then
        Label17.ForeColor = &H80FF&
    End If
    If vMediaGeral >= MediaGlobal Then
        Label17.ForeColor = &H8000&
    End If
End Sub

Private Function checkListADP()
    checkListADP = True
    '1º passo - Verificar no LV de "Avaliações" se todos os itens foram avaliados
    Dim X As Integer, Y As Integer
    
    If Date < DTPicker3 - 10 Then
        MsgBox "Essa ADP não pode ser CONCLUIDA fora do período de vencimento", vbCritical, "SGCH"
        checkListADP = False
        Exit Function
    End If
    
    If ListView1.ListItems.Count > 0 Then
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If Val(ListView1.SelectedItem.ListSubItems.Item(3)) = 0 Then
                MsgBox "Existem Itens da guia AVALIAÇÕES não avaliados", vbCritical, "SGCH"
                checkListADP = False
                Exit Function
            End If
        Next
    Else
        MsgBox "A guia AVALIAÇÕES não possui itens a serem avaliados", vbCritical, "SGCH"
        checkListADP = False
        Exit Function
    End If
    sqlADP = "Update tbListaADP set statusavaliacao = 'Concluido' Where codcoligada = '" & vCodcoligada & "' and codcolaborador = '" & Label18 & "'"
    rsADP.Open sqlADP, cnBanco
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = "Concluido"
    MsgBox "O processo de conclusão da ADP foi realizado com sucesso!"
End Function

Private Sub BloqueiaControles()
    Dim X As Integer
    For X = 0 To 15
        txtADP(X).Enabled = False
    Next
    txtLvw.Enabled = False
    
    cmdINTD(0).DragMode = 1
    cmdINTD(0).UseGreyscale = True
    cmdINTD(1).DragMode = 1
    cmdINTD(1).UseGreyscale = True
    cmdINTD(2).DragMode = 1
    cmdINTD(2).UseGreyscale = True
    cmdINTD(4).DragMode = 1
    cmdINTD(4).UseGreyscale = True
    cmdINTD(5).DragMode = 1
    cmdINTD(5).UseGreyscale = True
    cmdINTD(7).DragMode = 1
    cmdINTD(7).UseGreyscale = True
    cmdINTD(8).DragMode = 1
    cmdINTD(8).UseGreyscale = True
    cmdINTD(9).DragMode = 1
    cmdINTD(9).UseGreyscale = True
    cmdINTD(11).DragMode = 1
    cmdINTD(11).UseGreyscale = True
    
    DTPicker2.Enabled = False
    DTPicker4.Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    For X = 0 To 5
        chkADP(X).Enabled = False
    Next
    Combo1.Enabled = False
    chameleonButton1.Enabled = False
End Sub

'**********************************************
'**********************************************
'**********************************************
'**********************************************
'**********************************************

'----EDITA LISTVIEW DAKI P BAIXO------
'-------------------------------------
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, leftPos As Single 'the left pos of the column
Dim dx As Single, lvwX As Single  'the x in relation to listview coordinate

If Button = vbLeftButton Then
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.LabelEdit = lvwManual
        dx = GetLvwDeltaX
        lvwX = X + dx
        For i = 4 To 4
            leftPos = ListView1.Left + ListView1.ColumnHeaders(i).Left
            If lvwX > leftPos And lvwX < leftPos + ListView1.ColumnHeaders(i).Width Then 'we found the column
                m_RowIndex = ListView1.SelectedItem.Index 'row
                m_ColIndex = i 'column
                MoveTxtLvw dx 'move and size the edit box over the selected item
                With txtLvw 'turn on edit box
                    If i = 1 Then 'copy the text of the selected item to txtlvw
                        .Text = ListView1.SelectedItem.Text
                    Else
                        .Text = ListView1.SelectedItem.SubItems(i - 1)
                    End If
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    If txtLvw.Enabled = True Then .SetFocus Else cmdINTD(12).SetFocus
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
   
    Set lvwCol = ListView1.ColumnHeaders(ListView1.ColumnHeaders.Count)
    actualLvwWidth = lvwCol.Left + lvwCol.Width
    
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_ALL
    GetScrollInfo ListView1.HWnd, SB_HORZ, si
    maxScrollPos = si.nMax - si.nPage + 1 'formula from SDK, 0 if scroll bar is invinsible
    If maxScrollPos <> 0 Then GetLvwDeltaX = si.nPos / maxScrollPos * (actualLvwWidth - ListView1.Width + 58)
End Function

Sub MoveTxtLvw(Optional ByVal dx As Single = -1)
    Dim txtLeft As Single, txtWidth As Single, txtRight As Single, lvwCol As ColumnHeader
    Dim txtRightMax As Single, txtTop As Single, txtTopMin As Single, txtTopMax As Single
    
    
    If m_ColIndex Then
        If dx = -1 Then dx = GetLvwDeltaX 'called from subclass event
        Set lvwCol = ListView1.ColumnHeaders(m_ColIndex)
        
        txtLeft = ListView1.Left + lvwCol.Left + 48 - dx
        If txtLeft < ListView1.Left Then txtLeft = ListView1.Left + 48
    
        txtRightMax = ListView1.Left + ListView1.Width - 48
        If ScrollBarVisible(SB_VERT) Then txtRightMax = txtRightMax - 240
    
        If m_ColIndex = ListView1.ColumnHeaders.Count Then
            txtRight = txtRightMax
        Else
            txtRight = ListView1.Left + ListView1.ColumnHeaders(m_ColIndex + 1).Left - 8 - dx
            If txtRight > txtRightMax Then txtRight = txtRightMax
        End If
    
        txtWidth = txtRight - txtLeft
        If txtWidth < 0 Then txtWidth = 0: txtLeft = -1000
    
        txtTopMin = ListView1.Top
        If Not ListView1.HideColumnHeaders Then txtTopMin = txtTopMin + 210 'add height of header
        txtTopMax = ListView1.Top + ListView1.Height
        If ScrollBarVisible(SB_HORZ) Then txtTopMax = txtTopMax - 420 'minus height of scrollbar
    
        txtTop = ListView1.Top + ListView1.SelectedItem.Top + 54
        If txtTop < txtTopMin Or txtTop > txtTopMax Then txtTop = -1000 'move it out of view
    
        With txtLvw '.move produces runtime error with -ve values
            .Left = 7265
            .Top = txtTop + 4450
            .Width = 615
            .Height = ListView1.SelectedItem.Height - 9
        End With
    End If
End Sub

Private Sub txtADP_LostFocus(Index As Integer)
    Select Case Index
    Case 6
        calculaNotaADP
    Case 7
        calculaNotaADP
    End Select
End Sub

Private Sub txtLvw_GotFocus()
    If txtLvw.Text = "" Then txtLvw.Text = " "
End Sub

Private Sub txtLvw_KeyPress(KeyAscii As Integer)
    txtLvw.Tag = True 'ListView1 is edited
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
On Error GoTo TrataErro
    'AKI - desenvolver rotina para verificar qtd digitada
    If txtLvw.Text = " " Then txtLvw.Text = ""
    If Not IsNumeric(txtLvw.Text) And txtLvw.Text <> "" And Len(txtLvw) = 1 Then txtLvw.Text = "0"
    If m_ColIndex = 1 Then
        'Verifica com qual Listview vc esta trabalhando
        ListView1.ListItems(m_RowIndex).Text = Trim(txtLvw.Text) 'put in the text
        'add text entry to the last row
        'If ListView1.ListItems(ListView1.ListItems.Count) <> c_EntryTxt Then ListView1.ListItems.Add , , c_EntryTxt
    ElseIf m_ColIndex Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = Trim(txtLvw.Text)
    End If
    If ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) = "-" And ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2) < ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
        Exit Sub
    End If
    
    'A qtd do txtLvw nao pode ser maior q a qtd da coluna anterior
    If IsNumeric(txtLvw.Text) And Val(txtLvw.Text) > Val(ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 2)) Then
        ListView1.ListItems(m_RowIndex).SubItems(m_ColIndex - 1) = "0"
    End If
    
    txtLvw.Visible = False 'hide edit box
    m_RowIndex = 0
    m_ColIndex = 0
    txtADP(11).SetFocus
    'ListView1.SetFocus
TrataErro:
    Exit Sub
End Sub

Private Function ScrollBarVisible(ByVal fnBar As Long) As Boolean
'returns true if ListView1's vertical scrollbar is visible
Dim si As SCROLLINFO
    si.cbSize = 28 '7 long vars x 4 bytes
    si.fMask = SIF_PAGE Or SIF_RANGE 'retrieve page and range info only
    GetScrollInfo ListView1.HWnd, fnBar, si
    ScrollBarVisible = si.nPage <> si.nMax + 1 'maxScrollPos=0 if scrollbar is invinsible
End Function

Private Sub AtualizaListview()
'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(7) = Label17
    If RemoveMask(MeuLV.ListView1.SelectedItem.ListSubItems.Item(7)) >= MediaGlobal Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H8000&
    ElseIf RemoveMask(MeuLV.ListView1.SelectedItem.ListSubItems.Item(7)) < MediaGlobal And RemoveMask(MeuLV.ListView1.SelectedItem.ListSubItems.Item(7)) >= vAprovadoRest Then
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &H80FF&
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(7).ForeColor = &HC0&
    End If
    MeuLV.ListView1.SelectedItem.ListSubItems.Item(9) = "Avaliando"
    Exit Sub
Err:
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

