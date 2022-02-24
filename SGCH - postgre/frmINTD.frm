VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6E41052A-1C6B-4B1D-BE99-3928E843A6D8}#1.0#0"; "aicalphaimage.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmINTD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INTD - Identificação das Necessidades de Treinamento e Desenvolvimento"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13845
   Icon            =   "frmINTD.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   13
      Left            =   7800
      TabIndex        =   77
      Top             =   7440
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
      MICON           =   "frmINTD.frx":0CCA
      PICN            =   "frmINTD.frx":0CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame10 
      Caption         =   "Avaliação Escolaridade "
      Height          =   615
      Left            =   8880
      TabIndex        =   65
      Top             =   2280
      Width           =   2055
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Avaliação da INTD"
      Height          =   615
      Left            =   11040
      TabIndex        =   58
      Top             =   2280
      Width           =   1815
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   1575
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
      Left            =   1560
      TabIndex        =   56
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   3
      Left            =   6960
      TabIndex        =   18
      Tag             =   "Filtrar cursos/treinamentos da matriz em evidência"
      ToolTipText     =   "Filtrar cursos/treinamentos da matriz em evidência"
      Top             =   2280
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmINTD.frx":19C0
      PICN            =   "frmINTD.frx":19DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Enabled         =   0   'False
      Height          =   615
      Left            =   12600
      TabIndex        =   48
      Top             =   7440
      Width           =   1095
      Begin VB.CheckBox Check1 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Tag             =   "Status do curso/treinamento"
         ToolTipText     =   "Status do curso/treinamento"
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Objetivos"
      TabPicture(0)   =   "frmINTD.frx":26B6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtINTD(7)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cursos/Treinamentos"
      TabPicture(1)   =   "frmINTD.frx":26D2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label36"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdINTD(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdINTD(8)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdINTD(7)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdINTD(5)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdINTD(4)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ListView1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtINTD(8)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtINTD(9)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cboINTD(5)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Habilidades"
      TabPicture(2)   =   "frmINTD.frx":26EE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Resultado "
      TabPicture(3)   =   "frmINTD.frx":270A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame12"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(2)=   "Frame9"
      Tab(3).Control(3)=   "Frame7"
      Tab(3).Control(4)=   "aicAlphaImage1"
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame12 
         Caption         =   "ATENÇÃO "
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
         Height          =   855
         Left            =   -71280
         TabIndex        =   73
         Top             =   480
         Width           =   4815
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   74
            Text            =   "frmINTD.frx":2726
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Integração Totvs"
         Height          =   855
         Left            =   -65880
         TabIndex        =   68
         Tag             =   "Verifique se os dados de integração no cadastro do colaborador estão corretamente preenchidos"
         ToolTipText     =   "Verifique se os dados de integração no cadastro do colaborador estão corretamente preenchidos"
         Top             =   480
         Width           =   4335
         Begin VB.TextBox txtCons 
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   71
            Tag             =   "Função"
            ToolTipText     =   "Função"
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox Combo 
            Height          =   315
            Index           =   9
            Left            =   960
            TabIndex        =   70
            Tag             =   "Função"
            ToolTipText     =   "Função"
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label lblCons 
            Caption         =   "Função:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Resultado da INTD "
         Height          =   735
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   2775
         Begin VB.ComboBox cboINTD 
            Height          =   315
            Index           =   1
            ItemData        =   "frmINTD.frx":2787
            Left            =   120
            List            =   "frmINTD.frx":2791
            TabIndex        =   30
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Observações "
         Height          =   2775
         Left            =   -74880
         TabIndex        =   62
         Top             =   1320
         Width           =   13335
         Begin VB.TextBox txtINTD 
            Height          =   2415
            Index           =   13
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   13095
         End
      End
      Begin VB.ComboBox cboINTD 
         Height          =   315
         Index           =   5
         Left            =   7440
         TabIndex        =   24
         Top             =   630
         Width           =   3375
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   22
         Tag             =   "Nome do treinamento"
         ToolTipText     =   "Nome do treinamento"
         Top             =   660
         Width           =   5295
      End
      Begin VB.TextBox txtINTD 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Tag             =   "Código do treinamento"
         ToolTipText     =   "Código do treinamento"
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtINTD 
         Height          =   3735
         Index           =   7
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   20
         Tag             =   "Objetivo do INTD"
         ToolTipText     =   "Objetivo do INTD"
         Top             =   480
         Width           =   13335
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   29
         Tag             =   "habilidades a serem avaliadas na INTD"
         ToolTipText     =   "habilidades a serem avaliadas na INTD"
         Top             =   480
         Width           =   13335
         _ExtentX        =   23521
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2475
         Left            =   120
         TabIndex        =   28
         Tag             =   "Cursos/treinamentos do INTD"
         ToolTipText     =   "Cursos/treinamentos do INTD"
         Top             =   1740
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   4366
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
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   23
         Top             =   660
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
         MICON           =   "frmINTD.frx":27A5
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
         Index           =   5
         Left            =   1320
         TabIndex        =   27
         Tag             =   "Excluir treinamento"
         ToolTipText     =   "Excluir treinamento"
         Top             =   1020
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
         MICON           =   "frmINTD.frx":27C1
         PICN            =   "frmINTD.frx":27DD
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
         TabIndex        =   26
         Tag             =   "Novo treinamento"
         ToolTipText     =   "Novo treinamento"
         Top             =   1020
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
         MICON           =   "frmINTD.frx":34B7
         PICN            =   "frmINTD.frx":34D3
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
         TabIndex        =   25
         Tag             =   "Incluir treinamento"
         ToolTipText     =   "Incluir treinamento"
         Top             =   1020
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
         MICON           =   "frmINTD.frx":41AD
         PICN            =   "frmINTD.frx":41C9
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
         Index           =   6
         Left            =   12600
         TabIndex        =   75
         Tag             =   "Cursos/treinamentos exigidos pela matriz"
         ToolTipText     =   "Cursos/treinamentos exigidos pela matriz"
         Top             =   1020
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
         MICON           =   "frmINTD.frx":4EA3
         PICN            =   "frmINTD.frx":4EBF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
         Height          =   600
         Left            =   -66480
         Top             =   650
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         Image           =   "frmINTD.frx":5B99
         Props           =   5
      End
      Begin VB.Label Label36 
         Caption         =   "Nível:"
         Height          =   255
         Left            =   7440
         TabIndex        =   52
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Nome do curso/treinamento:"
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   420
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Colaborador "
      Height          =   1695
      Left            =   120
      TabIndex        =   38
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   4455
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   11
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
         MICON           =   "frmINTD.frx":20D98
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   10
         Tag             =   "Nome do colaborador em treinamento"
         ToolTipText     =   "Nome do colaborador em treinamento"
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtINTD 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Tag             =   "Registro do colaborador em treinamento"
         ToolTipText     =   "Registro do colaborador em treinamento"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label17 
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
         Left            =   4800
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "CPF:"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Matriz 
         Caption         =   "Matriz/Cargo"
         Height          =   255
         Left            =   1680
         TabIndex        =   54
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Treinar para o cargo de "
      Height          =   975
      Left            =   6960
      TabIndex        =   37
      Top             =   1200
      Width           =   6735
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   1680
         TabIndex        =   16
         Tag             =   "Código do cargo"
         ToolTipText     =   "Código do cargo"
         Top             =   480
         Width           =   855
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   15
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
         MICON           =   "frmINTD.frx":20DB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   17
         Tag             =   "Nome do cargo"
         ToolTipText     =   "Nome do cargo"
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox txtINTD 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Tag             =   "nº da matriz de capacitação"
         ToolTipText     =   "nº da matriz de capacitação"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Cód. cargo:"
         Height          =   255
         Left            =   1680
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Cargo/Nível:"
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Matriz nº:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Solicitante "
      Height          =   975
      Left            =   6960
      TabIndex        =   36
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox cboINTD 
         Height          =   315
         Index           =   0
         ItemData        =   "frmINTD.frx":20DD0
         Left            =   120
         List            =   "frmINTD.frx":20DE0
         TabIndex        =   5
         Text            =   "Colaborador"
         Top             =   480
         Width           =   1695
      End
      Begin SGCH.chameleonButton cmdINTD 
         Height          =   255
         Index           =   1
         Left            =   6255
         TabIndex        =   8
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
         MICON           =   "frmINTD.frx":20E1A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   7
         Tag             =   "Nome do solicitante"
         ToolTipText     =   "Nome do solicitante"
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtINTD 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Tag             =   "Registro do solicitante"
         ToolTipText     =   "Registro do solicitante"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   1920
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo "
      Height          =   975
      Left            =   4680
      TabIndex        =   35
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton optINTD 
         Caption         =   "Alteração funcional"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optINTD 
         Caption         =   "Capacitação funcional"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados do INTD "
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Tag             =   "Data de término do período da INTD"
         ToolTipText     =   "Data de término do período da INTD"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   102367233
         CurrentDate     =   40721
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Tag             =   "Data de início do período da INTD"
         ToolTipText     =   "Data de início do período da INTD"
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   102367233
         CurrentDate     =   40721
      End
      Begin VB.TextBox txtINTD 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data término:"
         Height          =   255
         Left            =   3000
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Data início:"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   615
      End
   End
   Begin SGCH.chameleonButton cmdINTD 
      Height          =   615
      Index           =   12
      Left            =   720
      TabIndex        =   33
      Tag             =   "Sair"
      ToolTipText     =   "Sair"
      Top             =   7440
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
      MICON           =   "frmINTD.frx":20E36
      PICN            =   "frmINTD.frx":20E52
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
      TabIndex        =   32
      Tag             =   "Salvar dados"
      ToolTipText     =   "Salvar dados"
      Top             =   7440
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
      MICON           =   "frmINTD.frx":21B2C
      PICN            =   "frmINTD.frx":21B48
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
      Left            =   12960
      TabIndex        =   69
      Tag             =   "Concluir INTD"
      ToolTipText     =   "Concluir INTD"
      Top             =   2280
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
      MICON           =   "frmINTD.frx":22139
      PICN            =   "frmINTD.frx":22155
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
      Index           =   10
      Left            =   8160
      TabIndex        =   76
      Tag             =   "Graduação exigida pela matriz"
      ToolTipText     =   "Graduação exigida pela matriz"
      Top             =   2280
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
      MICON           =   "frmINTD.frx":22E2F
      PICN            =   "frmINTD.frx":22E4B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblStatusINTD 
      BackColor       =   &H80000018&
      Height          =   255
      Left            =   11160
      TabIndex        =   64
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Foi atingido a data de término da INTD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   61
      Top             =   7680
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000018&
      Caption         =   "Status"
      Height          =   255
      Left            =   11160
      TabIndex        =   60
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmINTD"
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

Private rsINTD As New ADODB.Recordset
Private SqlINTD As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsCargoINTD As New ADODB.Recordset
Private sqlCargoINTD As String

Private rsLocal As New ADODB.Recordset
Private periodoEmMeses As Single

Private Sub cboINTD_Click(Index As Integer)
    Select Case Index
    Case 0
        MontaMascara 0
    End Select
End Sub

Private Sub cmdINTD_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridColaborador 0
        CarregaColaborador 3
    Case 1
        ChamaGridColaborador 1
        CarregaColaborador 1
    Case 2
        ChamaGridCargoINTD
        CarregaCargoINTD
    Case 3
        If ListView1.ListItems.Count > 0 Then
            If MsgBox("Todos os treinamento da guia Cursos/treinamentos serão apagados. Deseja continuar?", vbQuestion + vbYesNo, "ATENÇÃO") = vbNo Then Exit Sub
        End If
        'If optINTD(0).Value = True Then
        '    filtraLVTrei Val(Mid$(txtINTD(10), 1, 6))
        '    CompoeLVHab Val(Mid$(txtINTD(10), 1, 6))
        'Else
            filtraLVTrei Val(txtINTD(5))
            CompoeLVHab Val(txtINTD(5))
        'End If
        avaliaEscolar
    Case 4
        ChamaGridCurso
        CarregaCurso
        CompoeComboNivel cboINTD(5), txtINTD(8)
    Case 5
        If MsgBox("Deseja EXCLUIR curso/treinamento da INTD?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            ExcluirItemLV ListView1
            LimpaControlesTreinamento
        End If
    Case 6
        Campo4 = 2
        frmAvisos.Show 1
    Case 7
        LimpaControlesTreinamento
    Case 8
        IncluirTreinamento
        LimpaControlesTreinamento
    Case 9
        'CONCLUSÃO DA INTD
        If MsgBox("Deseja iniciar a conclusão da INTD?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            calculaINTD
            gravaDadosINTD
            If checkListINTD = False Then Exit Sub
            'gravaLog "Código PS: " & txtProcesso(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
            carregaADP txtINTD(12)
            Unload Me
        End If
        Unload Me
    Case 10
        Campo4 = 3
        frmAvisos.Show 1
    Case 11
        If MsgBox("Deseja salvar os dados da INTD?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            calculaINTD
            gravaDadosINTD
'            gravaLog "Código req: " & txtCadReq(0), "Requisitante" & txtCadReq(1) & "-" & txtCadReq(2), ""
            Pesquisa = "0"
        End If
    Case 12
        If MsgBox("Deseja sair da tela de cadastro de INTD?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
            Pesquisa = "0"
            Unload Me
            Set frmINTD = Nothing
        End If
    Case 13
        FCRAvaHab.Show 1
    End Select
End Sub

Private Sub Form_Activate()
    Status = Pesquisa
    If lblStatusINTD.Caption = "" Then lblStatusINTD.Caption = Status
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    If lblStatusINTD.Caption = "" Then lblStatusINTD.Caption = Status
    listview_cabecalho
    SSTab1.Tab = 0
    optINTD(0).Value = True
    optINTD_Click (0)
    If Status = "novo" Then
        txtINTD(0) = "-"
        LimpaControles
        Label14 = "Aberto"
    ElseIf Status = "editar" Then
        ResultPesq
    End If
    If vIntegra = "S" Then
        Frame11.Visible = True
        Frame12.Visible = True
        aicAlphaImage1.Visible = True
    Else
        Frame11.Visible = False
        Frame12.Visible = False
        aicAlphaImage1.Visible = False
    End If
    'configControles
    If vIntegra = "S" Then ConexaoTotvs
    If vIntegra = "S" Then comporCombosTotvs
    If vIntegra = "S" Then
        comporControlesTotvs
    End If
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome curso/treinamento", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Nível", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Status", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Pontuação", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Programação", ListView1.Width / 11
    
    ListView2.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView2.ColumnHeaders.Add , , "Habilidade", ListView1.Width / 2
    ListView2.ColumnHeaders.Add , , "Peso", ListView1.Width / 12
    ListView2.ColumnHeaders.Add , , "Avaliado", ListView1.Width / 12
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub ResultPesq()
'    SqlINTD = "select a.codINTD,a.datainicio,a.datafim,a.tipoINTD,a.tiposolicitante,a.codsolicitante,a.nomesolicitante,a.codcolaborador,a.codmatriz,a.status,a.ativo,a.objetivo,a.mediageral,a.resultado,a.observacao from tbINTD as a where a.codINTD = '" & Val(varGlobal) & "' and a.ativo='S'"
    SqlINTD = "select a.codINTD,a.datainicio,a.datafim,a.tipoINTD,a.tiposolicitante,a.codsolicitante,a.nomesolicitante,a.codcolaborador,a.codmatriz,a.status,a.ativo,a.objetivo,a.mediageral,a.resultado,a.observacao,a.mediaescolar,a.cargoorigem from tbINTD as a where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(varGlobal) & "'"
    rsINTD.Open SqlINTD, cnBanco, adOpenKeyset, adLockReadOnly
    If rsINTD.RecordCount > 0 Then
        CompoeControles
        CarregaColaborador 3
        'If optINTD(1).Value = True Then
            CarregaCargoINTD
        'End If
'--------------
        'If optINTD(0).Value = True Then
        '    CompoeLVTrei Val(Mid$(txtINTD(10), 1, 6))
        '    CompoeLVHab Val(Mid$(txtINTD(10), 1, 6))
        'Else
            CompoeLVTrei Val(txtINTD(5))
            CompoeLVHab Val(txtINTD(5))
        'End If
'--------------
        RestauraLVHab
        RestauraLVTrei
        calculaINTD
        If rsINTD.Fields(9) = "Fechado" Or rsINTD.Fields(9) = "Cancelada" Then
            If Not IsNull(rsINTD.Fields(16)) Then txtINTD(10).Text = rsINTD.Fields(16)
            BloqueiaControles
        Else
            lembrete
        End If
    Else
        MsgBox "INTD não encontrada"
    End If
    avaliaEscolar
    rsINTD.Close
    Set rsINTD = Nothing
End Sub

Private Sub lembrete()
    If Date > DTPicker2 Then
        Label15.Visible = True
    End If
End Sub

Private Sub BloqueiaControles()
    Dim X As Integer
    For X = 1 To 13
        txtINTD(X).Enabled = False
    Next
    txtLvw.Enabled = False
    For X = 0 To 5
        cmdINTD(X).Enabled = False
    Next
    For X = 7 To 9
        cmdINTD(X).Enabled = False
    Next
    cmdINTD(11).Enabled = False
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    optINTD(0).Enabled = False
    optINTD(1).Enabled = False
    cboINTD(0).Enabled = False
    cboINTD(1).Enabled = False
    cboINTD(5).Enabled = False
'    Check1.Enabled = False
'    Check2.Enabled = False
'    Check3.Enabled = False
'    Check4.Enabled = False
'    Check5.Enabled = False
'    Combo1.Enabled = False
'    Combo2.Enabled = False
'    cmdCadastro(0).Enabled = False
'    cmdCadastro(1).Enabled = False
'    cmdCadastro(2).Enabled = False
'    cmdCadastro(3).Enabled = False
'    cmdCadastro(4).Enabled = False
'    cmdCadastro(5).Enabled = False
'    cmdCadastro(6).Enabled = False
'    cmdCadastro(9).Enabled = False
'    cmdCadastro(11).Enabled = False
End Sub


Private Sub CompoeControles()
    Dim X As Integer
    txtINTD(0).Text = Format(rsINTD.Fields(0), "000000") 'código da requisição
    DTPicker1 = rsINTD.Fields(1) 'Data da Início
    DTPicker2 = rsINTD.Fields(2) 'Data da Término
    If rsINTD.Fields(3) = 0 Then
        optINTD(0).Value = True
    Else
        optINTD(1).Value = True
    End If
    cboINTD(0) = rsINTD.Fields(4)
    txtINTD(1) = rsINTD.Fields(5)
    txtINTD(2) = rsINTD.Fields(6)
    If Not IsNull(rsINTD.Fields(15)) Then Label16 = rsINTD.Fields(15) & "%"
'------
    Dim rsRegColab As New ADODB.Recordset
    Dim sqlRegColab As String
    sqlRegColab = "Select codcolaborador from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and id = '" & rsINTD.Fields(7) & "'"
    rsRegColab.Open sqlRegColab, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsRegColab.EOF Then
        Label17 = rsINTD.Fields(7)
        txtINTD(3) = rsRegColab.Fields(0)
    End If
    rsRegColab.Close
    Set rsRegColab = Nothing
'------
    
    'If optINTD(1).Value = True Then
        txtINTD(5) = rsINTD.Fields(8)
    'End If
    txtINTD(7) = rsINTD.Fields(11)
    If rsINTD.Fields(10) = "S" Then Check1.Value = 1 Else Check1.Value = 0  'Informa se a requisição esta ativa ou nao
    If Not IsNull(rsINTD.Fields(13)) Then cboINTD(1) = rsINTD.Fields(13)
    If Not IsNull(rsINTD.Fields(14)) Then txtINTD(13) = rsINTD.Fields(14)
End Sub

Private Sub filtraLVTrei(indice As Integer)
    If indice = 0 Then Exit Sub
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    Dim ItemLst As ListItem
    Dim X As Integer
    Dim vValidador As Boolean
    X = 0
    'Ao filtrar, essa rotina verifica primeiro se algum treinamento ja foi
    'programado e somente excluir os pendentes de programação
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If X > Y Then Exit For
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(3) = "Pendente" Or ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
            ListView1.ListItems.Remove (X)
            Y = ListView1.ListItems.Count
            X = 0
        End If
    Next
    '--------------------------------------------------------------
    '1º Monta os treinamentos da Matriz
    sqlTrei = "Select a.*, b.nometreinamento, c.codnivel, c.nomenivel from tbMatrizCur as a left join tbTreinamentos as b on a.codtreinamento=b.codtreinamento left join tbTreinamentosNiv as c on b.codtreinamento = c.codtreinamento and a.codnivel = c.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & indice & "'"
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockOptimistic
    
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        vValidador = True
        While Not rsTrei.EOF
            For X = 1 To Y
                ListView1.ListItems.Item(X).Selected = True
                If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(1) Then
                    vValidador = False
                End If
            Next
            If vValidador = True Then
                    Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                    ItemLst.SubItems(1) = "" & rsTrei.Fields(4)
                    If Not IsNull(rsTrei.Fields(5)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(5), "00") & " - " & rsTrei.Fields(6) Else ItemLst.SubItems(2) = "-"
                    ItemLst.SubItems(3) = "-"
                    ItemLst.SubItems(4) = "-"
                    ItemLst.SubItems(5) = "-"
            End If
            rsTrei.MoveNext
            vValidador = True
            Y = ListView1.ListItems.Count
        Wend
    Else
        While Not rsTrei.EOF
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
            ItemLst.SubItems(1) = "" & rsTrei.Fields(4)
            If Not IsNull(rsTrei.Fields(5)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(5), "00") & " - " & rsTrei.Fields(6) Else ItemLst.SubItems(2) = "-"
            ItemLst.SubItems(3) = "-"
            ItemLst.SubItems(4) = "-"
            ItemLst.SubItems(5) = "-"
            rsTrei.MoveNext
            'X = X + 1
        Wend
    End If
    
    ''--------------------------------------------------------------
    ''2º Monta os treinamentos INTRODUTORIOS da Matriz (EM Fase de Teste)
    rsTrei.Close
    
    sqlTrei = "select c.codmatriz,a.codtreinamento,c.nivel,b.nometreinamento,'-' from tbTreinamentosint as a inner join tbTreinamentos as b on a.codtreinamento = b.codtreinamento " & _
    "left join tbmatriz as c on a.codsetor = c.codsetor Where a.codsetor = 0 or c.codmatriz = '" & indice & "'"
    
'    sqlTrei = "select c.codmatriz,a.codtreinamento,c.nivel,b.nometreinamento,d.nomenivel from tbTreinamentosint as a inner join tbTreinamentos as b on a.codtreinamento = b.codtreinamento " & _
'    "left join tbmatriz as c on a.codsetor = c.codsetor left join tbTreinamentosNiv as d on cast(c.nivel as int) = d.codnivel Where a.codsetor = 0 or c.codmatriz = '" & indice & "'"
    
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockReadOnly
    
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        vValidador = True
        While Not rsTrei.EOF
            For X = 1 To Y
                ListView1.ListItems.Item(X).Selected = True
                If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(1) Then
                    vValidador = False
                End If
            Next
            If vValidador = True Then
                    Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
                    ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
'                    If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
                    If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") Else ItemLst.SubItems(2) = "-"
                    ItemLst.SubItems(3) = "-"
                    ItemLst.SubItems(4) = "-"
                    ItemLst.SubItems(5) = "-"
            End If
            rsTrei.MoveNext
            vValidador = True
            Y = ListView1.ListItems.Count
        Wend
    Else
        While Not rsTrei.EOF
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
            ItemLst.SubItems(1) = "" & rsTrei.Fields(3)
'            If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") & " - " & rsTrei.Fields(5) Else ItemLst.SubItems(2) = "-"
            If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(4), "00") Else ItemLst.SubItems(2) = "-"
            ItemLst.SubItems(3) = "-"
            ItemLst.SubItems(4) = "-"
            ItemLst.SubItems(5) = "-"
            rsTrei.MoveNext
            'X = X + 1
        Wend
    End If
    '--------------------------------------------------------------
    
    rsTrei.Close
    Set rsTrei = Nothing
End Sub

Private Sub CompoeLVTrei(indice As Integer)
    If indice = 0 Then Exit Sub
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    
    
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    
    '2º Monta os treinamentos alem da Matriz
    
    Dim vNivel As Integer, vValidador As Boolean
    sqlTrei = "Select a.*,b.nometreinamento,c.codnivel,c.nomenivel from tbINTDcur as a left join tbTreinamentos as b on a.codtreinamento=b.codtreinamento left join tbTreinamentosNiv as c on b.codtreinamento = c.codtreinamento and a.codnivel = c.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(txtINTD(0)) & "'"
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockReadOnly
    vValidador = True
    While Not rsTrei.EOF
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            '----
            If ListView1.SelectedItem.ListSubItems.Item(2) = "-" Then
                vNivel = 0
            Else
                vnilvel = Val(ListView1.SelectedItem.ListSubItems.Item(2))
            End If
            '----
            If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(1) And vNivel = rsTrei.Fields(2) Then
                vValidador = False
            End If
        Next
        If vValidador = True Then
            Set ItemLst = ListView1.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
            ItemLst.SubItems(1) = "" & rsTrei.Fields(4)
            If Not IsNull(rsTrei.Fields(4)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(5), "00") & " - " & rsTrei.Fields(6) Else ItemLst.SubItems(2) = "-"
            ItemLst.SubItems(3) = "-"
            ItemLst.SubItems(4) = "-"
            ItemLst.SubItems(5) = "-"
        End If
        rsTrei.MoveNext
        vValidador = True
        Y = ListView1.ListItems.Count
    Wend
    rsTrei.Close
    Set rsTrei = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub RestauraLVTrei()
    'Apos listar todas as habilidades da matriz. O sistema restaura a
    'pontuação avaliada pelo usuário
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    sqlTrei = "Select a.codtreinamento,a.codnivel,a.status,a.nota,a.codprogramacao,b.nota from tbPendentesCur as a inner join tbprogramacao as b on a.codprogramacao = b.codprogramacao where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(txtINTD(0)) & "'"
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    While Not rsTrei.EOF
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
'            If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(0) And Val(ListView1.SelectedItem.ListSubItems.Item(2)) = rsTrei.Fields(1) Then
            If Val(ListView1.ListItems.Item(X)) = rsTrei.Fields(0) Then
                ListView1.SelectedItem.ListSubItems.Item(3) = rsTrei.Fields(2)
                If Not IsNull(rsTrei.Fields(3)) And rsTrei.Fields(3) <> 0 Then
                    ListView1.SelectedItem.ListSubItems.Item(4) = Format(rsTrei.Fields(3), "#,##0.00;(#,##0.00)") & "%"
                Else
                    ListView1.SelectedItem.ListSubItems.Item(4) = Format(rsTrei.Fields(5), "#,##0.00;(#,##0.00)") & "%"
                End If
                
                If Not IsNull(rsTrei.Fields(4)) Then ListView1.SelectedItem.ListSubItems.Item(5) = Format(rsTrei.Fields(4), "000000")
                Exit For
            End If
        Next
        rsTrei.MoveNext
    Wend
    Me.ListView1.ColumnHeaders(5).Alignment = lvwColumnRight
    rsTrei.Close
    Set rsTrei = Nothing
End Sub

Private Sub CompoeLVHab(indice As Integer)
    If indice = 0 Then Exit Sub
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidades As String
    sqlHabilidades = "Select tbMatrizHab.*, tbhabilidades.nomehabilidade, tbhabilidades.peso from tbMatrizHab, tbhabilidades where tbMatrizHab.codcoligada = '" & vCodcoligada & "' and tbMatrizHab.codhabilidade = tbhabilidades.codhabilidade and tbMatrizHab.codmatriz = '" & indice & "'order by tbMatrizHab.codhabilidade"
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
    Me.ListView2.ColumnHeaders(3).Alignment = lvwColumnRight
    Me.ListView2.ColumnHeaders(4).Alignment = lvwColumnRight
    rsHabilidade.Close
    Set rsHabilidade = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub RestauraLVHab()
    'Apos listar todas as habilidades da matriz. O sistema restaura a
    'pontuação avaliada pelo usuário
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidades As String
    sqlHabilidades = "Select a.codINTD,a.codHabilidade,a.pontuacao from tbINTDHab as a where a.codcoligada = '" & vCodcoligada & "' and a.codINTD = '" & Val(txtINTD(0)) & "'"
    rsHabilidade.Open sqlHabilidades, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    While Not rsHabilidade.EOF
        Y = ListView2.ListItems.Count
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True
            If Val(ListView2.ListItems.Item(X)) = rsHabilidade.Fields(1) Then
                ListView2.SelectedItem.ListSubItems.Item(3) = rsHabilidade.Fields(2)
                Exit For
            End If
        Next
        rsHabilidade.MoveNext
    Wend
    rsHabilidade.Close
    Set rsHabilidade = Nothing
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    For X = 1 To txtINTD.Count - 1
        txtINTD(X) = ""
    Next
    cboINTD(0).Text = "Colaborador"
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    'txtCadReq(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub optINTD_Click(Index As Integer)
    'Select Case Index
    'Case 0
    '    ControlaOpt False
    'Case 1
        ControlaOpt True
    'End Select
End Sub

Private Sub ControlaOpt(VT As Boolean)
    Frame4.Enabled = VT
    txtINTD(5).Enabled = VT
    cmdINTD(2).Enabled = VT
    Label4.Enabled = VT
    Label5.Enabled = VT
    Label25.Enabled = VT
    txtINTD(5) = ""
    txtINTD(6) = ""
    txtINTD(11) = ""
    ListView2.Enabled = VT
End Sub

Private Sub MontaMascara(indice As Integer)
    If indice = 0 Then
        If cboINTD(0) <> "Colaborador" Then
            txtINTD(1) = Format(0, "000000")
            txtINTD(1).Enabled = False
            txtINTD(2).Enabled = True
            txtINTD(2).BackColor = &H80000018
            If txtINTD(2) = "" Then txtINTD(2).Text = "Digite o nome do solicitante"
            cmdINTD(1).Enabled = False
        ElseIf cboINTD(0) = "Colaborador" Then
            txtINTD(1).Enabled = True
            txtINTD(2).Enabled = False
            txtINTD(2).BackColor = &H80000005
            If txtINTD(1) <> "000000" Or txtINTD(1) = "" Then
                txtINTD(1).Text = ""
                txtINTD(2).Text = ""
            End If
            cmdINTD(1).Enabled = True
            CarregaColaborador 1
        End If
    End If
End Sub

Private Sub CarregaColaborador(indice As Integer)
    Dim X As Integer
    SqlColaborador = "select a.codcolaborador,a.nomecolaborador,d.nomedepartamento,e.nomesetor,c.codmatriz,f.codcargo,f.nomecargo,c.nivel,a.cpf,a.id from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d " & _
    "on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor inner join tbcargos as f on c.codcargo = f.codcargo where a.codcolaborador = '" & txtINTD(indice) & "'"
    rsColaborador.Open SqlColaborador, cnBanco, adOpenKeyset, adLockReadOnly
    If rsColaborador.RecordCount <= 0 Then
        If indice = 1 Then
            If txtINTD(1).Text <> "000000" And txtINTD(1).Text <> "" Then MsgBox "Colaborador não cadastrado", vbInformation, "SGCH"
            txtINTD(2) = ""
        Else
            If txtINTD(3).Text <> "000000" And txtINTD(1).Text <> "" Then MsgBox "Colaborador não cadastrado", vbInformation, "SGCH"
            txtINTD(4) = ""
        End If
    Else
        If indice = 1 Then
            txtINTD(1).Text = rsColaborador.Fields(0)
            txtINTD(2).Text = rsColaborador.Fields(1)
            Label17 = rsColaborador.Fields(9)
        Else
            txtINTD(3).Text = rsColaborador.Fields(0)
            txtINTD(4).Text = rsColaborador.Fields(1)
            txtINTD(10) = Format(rsColaborador.Fields(4), "000000") & " - " & rsColaborador.Fields(6) & " (" & rsColaborador.Fields(7) & ")"
            txtINTD(12) = rsColaborador.Fields(8)
            Label17 = rsColaborador.Fields(9)
        End If
    End If
    rsColaborador.Close
    Set rsColaborador = Nothing
End Sub

Private Sub CarregaCurso()
    Dim X As Integer
    Dim SqlCursos As String
    Dim rsCursos As New ADODB.Recordset
    SqlCursos = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbTreinamentos.codtreinamento"
    rsCursos.Open SqlCursos, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsCursos.EOF Then rsCursos.MoveFirst
    rsCursos.Find "codtreinamento=" & "'" & Val(Me.txtINTD(8)) & "'"
    If rsCursos.EOF Then
        txtINTD(8).Text = Format(txtINTD(8), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            MsgBox "Curso/Treinamento não cadastrado", vbInformation, "SGCH"
            txtINTD(9) = ""
        End If
    Else
        txtINTD(8).Text = Format(rsCursos.Fields(0), "000000") & ""
        txtINTD(9).Text = rsCursos.Fields(1)
    End If
    rsCursos.Close
    Set rsCursos = Nothing
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
    Pesquisa = frmINTD.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If indice = 1 Then
                txtINTD(1).Text = rsLocal.Fields(1)
            Else
                txtINTD(3).Text = rsLocal.Fields(1)
                Label17 = rsLocal.Fields(29)
            End If
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub ChamaGridCargoINTD()
On Error GoTo Err
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "select tbmatriz.codmatriz,tbcargos.nomecargo,tbmatriz.nivel,tbdepartamentos.nomedepartamento,tbsetores.nomesetor from tbmatriz,tbdepartamentos,tbsetores,tbcargos where tbmatriz.codcoligada = '" & vCodcoligada & "' and tbmatriz.coddepartamento = tbdepartamentos.coddepartamento and tbmatriz.codsetor = tbsetores.codsetor and tbmatriz.codcargo = tbcargos.codcargo and tbmatriz.ativo = 'S' order by tbcargos.nomecargo,tbMatriz.nivel"
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
            txtINTD(5).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
    Exit Sub
Err:
    Exit Sub
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
    Pesquisa = frmINTD.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nometreinamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtINTD(8).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub txtINTD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 1
        End If
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 3
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCargoINTD
        End If
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCurso
            CompoeComboNivel cboINTD(5), txtINTD(8)
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaCargoINTD()
    Dim X As Integer
    sqlCargoINTD = "Select tbMatriz.codmatriz,tbMatriz.codcargo,tbMatriz.nivel,tbcargos.nomecargo from tbMatriz,tbcargos where tbMatriz.codcoligada = '" & vCodcoligada & "' and tbMatriz.codcargo = tbCargos.codcargo and tbmatriz.ativo = 'S'  and tbmatriz.codmatriz = '" & Val(txtINTD(5)) & "' order by tbMatriz.codmatriz"
    rsCargoINTD.Open sqlCargoINTD, cnBanco, adOpenKeyset, adLockOptimistic
    If rsCargoINTD.RecordCount <= 0 Then
        txtINTD(5).Text = Format(txtINTD(5), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            MsgBox "Matriz não cadastrada", vbInformation, "SGCH"
            txtINTD(5) = ""
            txtINTD(11) = ""
            txtINTD(6) = ""
        End If
    Else
        txtINTD(5).Text = Format(rsCargoINTD.Fields(0), "000000") & ""
        txtINTD(11).Text = Format(rsCargoINTD.Fields(1), "000000") & ""
        txtINTD(6).Text = rsCargoINTD.Fields(3) & " (" & rsCargoINTD.Fields(2) & ")"
    End If
    rsCargoINTD.Close
    Set rsCargoINTD = Nothing
End Sub

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE CURSOS/TREINAMENTOS <<<<<<<<<<
Private Sub IncluirTreinamento()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaTreinamento = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtINTD(8) Then
                Me.txtINTD(8) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtINTD(9)
                ListView1.SelectedItem.ListSubItems.Item(2) = cboINTD(5)
                ListView1.SelectedItem.ListSubItems.Item(3) = "-"
                ListView1.SelectedItem.ListSubItems.Item(4) = "-"
                ListView1.SelectedItem.ListSubItems.Item(5) = "-"
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtINTD(8))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtINTD(8))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtINTD(9)
    ItemLst.SubItems(2) = cboINTD(5)
    ItemLst.SubItems(3) = "-"
    ItemLst.SubItems(4) = "-"
    ItemLst.SubItems(5) = "-"
    txtINTD(8).SetFocus
End Sub

Private Function ValidaTreinamento()
    ValidaTreinamento = False
    If txtINTD(9).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtINTD(9).Tag, vbInformation, "Atenção"
        Me.txtINTD(8).SetFocus
        Exit Function
    End If
    ValidaTreinamento = True
End Function

Private Sub LimpaControlesTreinamento()
    Dim X As Integer
    txtINTD(8).Enabled = True
    cmdINTD(4).Enabled = True
    
    For X = 8 To 9
        txtINTD(X) = ""
    Next
    txtINTD(8).SetFocus
End Sub

Private Sub gravaDadosINTD()
'On Error GoTo TrataErro
    If ValidaCampos = False Then Exit Sub
    Dim rsSalvarINTD As New ADODB.Recordset
    Dim SqlSalvarINTD As String
    
    Dim rsSalvarColaboradorTotvs As New ADODB.Recordset
    Dim SqlSalvarColaboradorTotvs As String
    
    cnBanco.BeginTrans
    If txtINTD(0) <> "-" Then
        SqlSalvarINTD = "select * from tbINTD where codcoligada = '" & vCodcoligada & "' and codintd = '" & Val(txtINTD(0)) & "'"
    Else
        SqlSalvarINTD = "select * from tbINTD where codcoligada = '" & vCodcoligada & "' and codintd = '" & GeraCodigo & "'"
    End If
    rsSalvarINTD.Open SqlSalvarINTD, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarINTD.EOF Then rsSalvarINTD.AddNew
    rsSalvarINTD.Fields(0) = txtINTD(0) 'Codigo INTD
    rsSalvarINTD.Fields(1) = DTPicker1 'Data inicio
    rsSalvarINTD.Fields(2) = DTPicker2 'Data fim
    If optINTD(0).Value = True Then rsSalvarINTD.Fields(3) = 0 Else rsSalvarINTD.Fields(3) = 1 ' Tipo INTD
    rsSalvarINTD.Fields(4) = cboINTD(0) 'Tipo solicitante
    rsSalvarINTD.Fields(5) = txtINTD(1) 'Codigo do solicitante
    rsSalvarINTD.Fields(6) = txtINTD(2) 'Nome do solicitante
    rsSalvarINTD.Fields(7) = Label17 'txtINTD(3) 'Codigo do colaborador a ser treinado
    'If optINTD(0).Value = True Then
    '    rsSalvarINTD.Fields(8) = Val(Mid$(txtINTD(10), 1, 6)) 'Codigo da matriz
    'Else
        rsSalvarINTD.Fields(8) = Val(txtINTD(5)) 'Codigo da matriz
    'End If
    If lblStatusINTD.Caption = "novo" Then
        rsSalvarINTD.Fields(9) = "Aberto" 'Status
        Label14 = "Aberto"
        rsSalvarINTD.Fields(17) = txtINTD(10).Text
        'O BLOCO DE GRAVAÇÃO DE TREINAMENTO INTRODUTÓRIOS/OBRIGATÓRIOS FOI DESLOCADO PARA ESSE PONTO
        'QUANDO SE GRAVA UMA NOVA INTD
        GravaTreiPen txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        'If GeraIntr = "S" Then GravaTreiIntrodutorio txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        'If GeraObri = "S" Then GravaTreiObrigatorio txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        
    Else
        rsSalvarINTD.Fields(9) = "Aberto" 'Status
        Label14 = "Aberto"
        If ListView1.ListItems.Count > 0 Then
            Dim X As Integer, Y As Integer
            Y = ListView1.ListItems.Count
            For X = 1 To Y
                ListView1.ListItems.Item(X).Selected = True
                If ListView1.SelectedItem.ListSubItems.Item(3) <> "Pendente" Then 'And ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then
                    rsSalvarINTD.Fields(9) = "Em Andamento" 'Status
                    Label14 = "Em Andamento"
                End If
            Next
        End If
    End If
    If Check1.Value = 1 Then rsSalvarINTD.Fields(10) = "S" Else rsSalvarINTD.Fields(10) = "N" 'ativo
    rsSalvarINTD.Fields(11) = txtINTD(7) 'Objetivo
    rsSalvarINTD.Fields(12) = RemoveMask(Label13) 'Media Geral da INTD
    rsSalvarINTD.Fields(13) = cboINTD(1) 'Resultado
    rsSalvarINTD.Fields(14) = txtINTD(13) 'Observação
    rsSalvarINTD.Fields(15) = vCodcoligada 'Codigo da coligada
    rsSalvarINTD.Fields(16) = RemoveMask(Label16) 'Média escolar
    
    SqlSalvarColaboradorTotvs = "Update tbColaboradoresIntTotvs set funcao = '" & txtCons(8) & "' Where codcoligada = '" & vCodcoligada & "' and id = '" & Val(Label17) & "'"
    rsSalvarColaboradorTotvs.Open SqlSalvarColaboradorTotvs, cnBanco
    
    rsSalvarINTD.Update
    rsSalvarINTD.Close
    Set rsSalvarINTD = Nothing
    
    gravaDadosINTDPendCur
    gravaDadosINTDCur
    gravaDadosINTDHab
    'gravaDadosINTDPendCur
    cnBanco.CommitTrans
    
    MsgBox "Os dados da INTD foram salvos com sucesso", vbInformation, "SGCH"
    AtualizaListview
    Exit Sub
TrataErro:
    MsgBox "Ocorreu um erro, as alterções nos registros serão desfeitas!", vbInformation, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub
    
Private Sub calculaINTD()
    Dim vValor1 As Double, vValor2 As Double, vValor3 As Double, vMediaGeral As Double
    Dim X As Integer, Y As Integer
    If ListView1.ListItems.Count > 0 Then
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.SelectedItem.ListSubItems.Item(4) <> "-" Then
                vValor1 = vValor1 + RemoveMask(ListView1.SelectedItem.ListSubItems.Item(4))
            End If
        Next
        vValor1 = vValor1 / Y
    Else
        vValor1 = 100
    End If

    If ListView2.ListItems.Count > 0 Then
        Y = ListView2.ListItems.Count
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True
            If Val(ListView2.SelectedItem.ListSubItems.Item(3)) <> 0 Then
                vValor2 = vValor2 + Val(ListView2.SelectedItem.ListSubItems.Item(3))
            End If
        Next
        vValor2 = vValor2 / Y
    Else
        vValor2 = 100
    End If
    vValor3 = Val(RemoveMask(Label16))
    vMediaGeral = (vValor1 + vValor2 + vValor3) / 3
    Label13 = Format(vMediaGeral, "#,##0.00;(#,##0.00)") & " %"
    
    If vMediaGeral < vAprovadoRest Then
        Label13.ForeColor = &HC0&
    End If
    If vMediaGeral < MediaGlobal And vMediaGeral >= vAprovadoRest Then
        Label13.ForeColor = &H80FF&
    End If
    If vMediaGeral >= MediaGlobal Then
        Label13.ForeColor = &H8000&
    End If
End Sub
    
Private Sub gravaDadosINTDCur()
    Dim rsSalvarINTDCur As New ADODB.Recordset
    Dim SqlSalvarINTDCur As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    sqlDeletar = "Delete from tbINTDCur where codcoligada = '" & vCodcoligada & "' and tbINTDCur.codINTD = '" & Val(txtINTD(0)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    If ListView1.ListItems.Count > 0 Then
        SqlSalvarINTDCur = "Select * from tbINTDCur where codcoligada = '" & vCodcoligada & "'"
        rsSalvarINTDCur.Open SqlSalvarINTDCur, cnBanco, adOpenKeyset, adLockOptimistic
        For X = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(X).Selected = True
            rsSalvarINTDCur.AddNew
            rsSalvarINTDCur.Fields(0) = txtINTD(0)
            rsSalvarINTDCur.Fields(1) = Val(ListView1.ListItems.Item(X))
            If ListView1.SelectedItem.ListSubItems.Item(2) <> "-" Then
                rsSalvarINTDCur.Fields(2) = Val(Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 2))
            Else
                rsSalvarINTDCur.Fields(2) = 0
            End If
            rsSalvarINTDCur.Fields(3) = vCodcoligada ' Codigo da coligada
        Next
        If Not rsSalvarINTDCur.EOF Then rsSalvarINTDCur.Update
        rsSalvarINTDCur.Close
        Set rsSalvarINTDCur = Nothing
    End If
End Sub

Private Sub gravaDadosINTDHab()
    Dim rsSalvarINTDHab As New ADODB.Recordset
    Dim SqlSalvarINTDHab As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    
    If ListView2.ListItems.Count > 0 Then
        SqlSalvarINTDHab = "Select * from tbINTDHab where tbINTDHab.codcoligada = '" & vCodcoligada & "' and tbINTDHab.codINTD = '" & Val(txtINTD(0)) & "'"
        rsSalvarINTDHab.Open SqlSalvarINTDHab, cnBanco, adOpenKeyset, adLockOptimistic

        'If rsSalvarINTDHab.RecordCount <> ListView2.ListItems.Count Then
        If rsSalvarINTDHab.RecordCount > 0 Then '<> ListView2.ListItems.Count Then
            sqlDeletar = "Delete from tbINTDHab where tbINTDHab.codcoligada = '" & vCodcoligada & "' and tbINTDHab.codINTD = '" & Val(txtINTD(0)) & "'"
            rsDeletar.Open sqlDeletar, cnBanco
        End If

        For X = 1 To ListView2.ListItems.Count
            ListView2.ListItems.Item(X).Selected = True
            If ListView2.ListItems.Item(X).Checked = True Then
                rsSalvarINTDHab.Find "codhabilidade=" & "'" & Val(ListView2.ListItems.Item(X)) & "'"
                If rsSalvarINTDHab.EOF Then rsSalvarINTDHab.AddNew
                rsSalvarINTDHab.Fields(0) = txtINTD(0)
                rsSalvarINTDHab.Fields(1) = Val(ListView2.ListItems.Item(X))
                rsSalvarINTDHab.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(3)
                rsSalvarINTDHab.Fields(3) = vCodcoligada 'Codigo da coligada
            End If
        Next
        If Not rsSalvarINTDHab.EOF Then rsSalvarINTDHab.Update
        rsSalvarINTDHab.Close
        Set rsSalvarINTDHab = Nothing
    End If
End Sub

Private Sub gravaDadosINTDPendCur()
    'Grava todos os treinamentos listador no form frmINTD na tabela
    'tbPendentesCur
    Dim rsTreiPen As New ADODB.Recordset
    Dim sqlTreiPen As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim contaID As Integer
    
    'Rotina de deletar se INTD estiver com status de Pendente
    sqlDeletar = "Delete from tbPendentesCur where tbPendentesCur.codcoligada = '" & vCodcoligada & "' and tbPendentesCur.codINTD = '" & Val(txtINTD(0)) & "' and status = 'Pendente'"
    rsDeletar.Open sqlDeletar, cnBanco
    '--------------------------------------------------------
    
    sqlTreiPen = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id"
    rsTreiPen.Open sqlTreiPen, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsTreiPen.EOF Then
        rsTreiPen.MoveLast
        contaID = rsTreiPen.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsTreiPen.Close
    Set rsTreiPen = Nothing
    
    sqlTreiPen = "Select * from tbPendentesCur as a where codcoligada = '" & vCodcoligada & "' and codINTD = '" & Val(txtINTD(0)) & "'"
    rsTreiPen.Open sqlTreiPen, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        ListView1.ListItems.Item(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(3) = "Pendente" Or ListView1.SelectedItem.ListSubItems.Item(3) = "-" Then
            rsTreiPen.AddNew
            rsTreiPen.Fields(0) = txtINTD(12)
            rsTreiPen.Fields(1) = Val(Mid$(txtINTD(10), 1, 6)) 'Codigo da matriz
            rsTreiPen.Fields(2) = Val(ListView1.ListItems.Item(X))
            rsTreiPen.Fields(4) = "S"
            rsTreiPen.Fields(5) = contaID
            rsTreiPen.Fields(6) = "Pendente"
            rsTreiPen.Fields(7) = 0
            'rsTreiPen.Fields(9) = ListView1.SelectedItem.ListSubItems.Item(4)
            If ListView1.SelectedItem.ListSubItems.Item(2) <> "-" Then
                rsTreiPen.Fields(12) = Val(ListView1.SelectedItem.ListSubItems.Item(2))
            Else
                rsTreiPen.Fields(12) = 0
            End If
            rsTreiPen.Fields(13) = Val(txtINTD(0))
            rsTreiPen.Fields(14) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        End If
    Next
    If Y > 0 Then rsTreiPen.Update
    rsTreiPen.Close
    Set rsTreiPen = Nothing
End Sub

Private Function ValidaCampos()
    ValidaCampo = False
    
    If Date > DTPicker2 Then
        Label7.Visible = True
    End If
    
    If txtINTD(1).Text = "" Then
        MsgBox "Favor informar o campo " & Me.txtINTD(1).Tag, vbInformation, "Atenção"
        Me.txtINTD(1).SetFocus
        Exit Function
    End If
    
    If Label16.Caption = "" Then
        MsgBox "Clique no filtro antes de salvar " & Me.txtINTD(1).Tag, vbInformation, "Atenção"
        Exit Function
    End If
    
    
    If optINTD(0).Value = True Then
        If txtINTD(3).Text = "" Then
            MsgBox "Favor informar o campo " & Me.txtINTD(3).Tag, vbInformation, "Atenção"
            Me.txtINTD(3).SetFocus
            Exit Function
        End If
    Else
        If txtINTD(5).Text = "" Then
            MsgBox "Favor informar o campo " & Me.txtINTD(5).Tag, vbInformation, "Atenção"
            Me.txtINTD(5).SetFocus
            Exit Function
        End If
    End If
    
    'If ListView1.ListItems.Count = 0 Then
    '    MsgBox "Favor informar " & ListView1.Tag, vbInformation, "Atenção"
    '    SSTab1.Tab = 1
    '    Me.txtINTD(8).SetFocus
    '    Exit Function
    'End If
    'If ListView2.ListItems.Count = 0 Then
    '    MsgBox "Favor informar " & ListView2.Tag, vbInformation, "Atenção"
    '    SSTab1.Tab = 2
    '    Exit Function
    'End If
    ValidaCampos = True
End Function

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera As String
    AbrirINTD
    SqlGera = "Select top 1 * from tbINTD where codcoligada = '" & vCodcoligada & "' order by codINTD Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsINTD.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtINTD(0) = Format(GeraCodigo, "000000")
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharINTD
End Function

Private Sub AbrirINTD()
    SqlINTD = "Select * from tbINTD where codcoligada = '" & vCodcoligada & "' Order by codINTD"
    rsINTD.Open SqlINTD, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharINTD()
    rsINTD.Close
    Set rsINTD = Nothing
End Sub

Private Sub AtualizaListview()
'On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If lblStatusINTD.Caption = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtINTD(0), "000000"))
        ItemLst.SubItems(1) = DTPicker1
        ItemLst.SubItems(2) = DTPicker2
        ItemLst.SubItems(3) = txtINTD(3)
        ItemLst.SubItems(4) = txtINTD(4)
        ItemLst.SubItems(5) = Label14
        If Check1.Value = 0 Then
            ItemLst.SubItems(6) = ""
            ItemLst.ListSubItems.Item(6).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(6) = ""
            ItemLst.ListSubItems.Item(6).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = DTPicker1
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = DTPicker2
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = txtINTD(3)
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtINTD(4)
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = Label14
        If Check1.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(6).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    MsgBox "Não foi possível realizar as alterações", vbInformation, "Atenção"
    Exit Sub
End Sub

Private Function checkListINTD()
    checkListINTD = True
    
    If Mid$(txtINTD(10), 1, 6) = txtINTD(5).Text Then
        SqlINTD = "Update tbINTD set status = 'Fechado' Where codcoligada = '" & vCodcoligada & "' and codINTD = '" & Val(txtINTD(0)) & "'"
        rsINTD.Open SqlINTD, cnBanco
        Label14 = "Fechado"
        MsgBox "O processo de conclusão da INTD nº: " & txtINTD(0) & " foi realizado com sucesso!"
        AtualizaListview
        'rsINTD.Close
        'Set rsINTD = Nothing
        Exit Function
    End If
    
    '1º passo - Verificar se o campo "Objetivos" esta preenchido
    If txtINTD(7).Text = "" Then
        MsgBox "Não foram apresentados objetivos para a INTD"
        SSTab1.Tab = 0
        txtINTD(7).SetFocus
        checkListINTD = False
        Exit Function
    End If
    '2º passo - Verificar no LV de "Cursos/treinamentos" se todos os treinamentos estão concluidos
    Dim X As Integer, Y As Integer
    If ListView1.ListItems.Count > 0 Then
        Y = ListView1.ListItems.Count
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.SelectedItem.ListSubItems.Item(3) <> "Concluido" And ListView1.SelectedItem.ListSubItems.Item(3) <> "Cancelado" Then
                MsgBox "Existem Cursos/treinamentos não concluidos"
                checkListINTD = False
                Exit Function
            End If
        Next
    End If
    '3º passo - Verificar no LV de "Habilidades" se todas as habilidades foram avaliadas
    If optINTD(1).Value = True Then
        If ListView2.ListItems.Count > 0 Then
            Y = ListView2.ListItems.Count
            For X = 1 To Y
                ListView2.ListItems.Item(X).Selected = True
                If ListView2.SelectedItem.ListSubItems.Item(3) = 0 Then
                    MsgBox "Existem Habilidades não avaliadas"
                    checkListINTD = False
                    Exit Function
                End If
            Next
        End If
    End If
    '4º passo - Verificar se o combo de resultado da INTD na guia de "Resultado" está preenchido
    If cboINTD(1).Text = "" Then
        MsgBox "Não foi avaliado o Resultado da INTD"
        SSTab1.Tab = 3
        cboINTD(1).SetFocus
        checkListINTD = False
        Exit Function
    End If
    '5º passo - Verificar se é uma troca de cargo, caso seja, inicialiar gravação de troca de cargo
    
    'Apto e Alteração Funcional
    If optINTD(1).Value = True And cboINTD(1).Text = "Apto" Then
        'VERIFICAR SE HÁ INTEGRAÇÃO TOTVS
        If vIntegra = "S" Then
            If txtCons(8) = "" Then
                checkListINTD = False
                Exit Function
            Else
                
                Dim rsDadosTotvs As New ADODB.Recordset
                Dim SqlDBTotvs As String
                
                SqlDBTotvs = "Select a.nomecolaborador,a.datanascimento,a.ctpsnumero,a.foto,b.sexo,b.grauinst,b.tipoadm,b.motadm,b.forreceb,b.situacao,b.tipofunc,b.hortrab,b.funcao,b.secao,b.contsind,b.rais,b.memsind " & _
                "from tbColaboradores as a LEFT join tbColaboradoresIntTotvs as b on a.id = b.id where a.id = '" & Val(Label17) & "'"
                rsDadosTotvs.Open SqlDBTotvs, cnBanco, adOpenKeyset, adLockReadOnly
                
                vDadosTotvs(0) = txtINTD(3) 'Chapa
                vDadosTotvs(1) = rsDadosTotvs.Fields(0) 'Nome do colaborador
                vDadosTotvs(2) = rsDadosTotvs.Fields(1) 'Data de nascimento
                vDadosTotvs(3) = rsDadosTotvs.Fields(2) 'Carteira de trabalho
                vDadosTotvs(4) = rsDadosTotvs.Fields(3) 'caminho foto
                vDadosTotvs(5) = rsDadosTotvs.Fields(4) 'sexo
                vDadosTotvs(6) = rsDadosTotvs.Fields(5) 'grau de instrução
                vDadosTotvs(7) = rsDadosTotvs.Fields(6) 'tipo de admissão
                vDadosTotvs(8) = rsDadosTotvs.Fields(7) 'motivo da admissão
                vDadosTotvs(9) = rsDadosTotvs.Fields(8) 'forma de recebimento
                vDadosTotvs(10) = rsDadosTotvs.Fields(9) 'Situação
                vDadosTotvs(11) = rsDadosTotvs.Fields(10) 'Tipo de funcionário
                vDadosTotvs(12) = rsDadosTotvs.Fields(11) 'horário de trabalho
                vDadosTotvs(13) = rsDadosTotvs.Fields(12) 'função
                vDadosTotvs(14) = rsDadosTotvs.Fields(13) 'seção
                vDadosTotvs(15) = rsDadosTotvs.Fields(14) 'contribuição sindical
                vDadosTotvs(16) = rsDadosTotvs.Fields(15) 'situação/rais
                vDadosTotvs(17) = rsDadosTotvs.Fields(16) 'membro sindicato
                For X = 0 To 17
                    If vDadosTotvs(X) = "" Then
                        MsgBox "Dados incompletos para exportar para o Corpore RM Labore. Favor conferir no cadastro", vbCritical, "SGCH"
                        'GoTo TrataErro
                    End If
                Next
                GravaDadosDBTotvs txtINTD(3)
            End If
        End If
        'VERIFICAR PONTUAÇÃO DE ESCOLARIDADE E PERMISSÃO DO USUARIO
        If pontosEscolar = False Then
            checkListINTD = False
            Exit Function
        End If
        'COMPUTANDO TEMPO DE EXPERIÊNCIA
        periodoEmMeses = DateDiff("m", DTPicker1.Value, Now)
        registraExperiencia
        'ALTERANDO CARGO
        alteraCargo
        'Pontuar Habilidades
        pontuaHabilidade Val(txtINTD(5))
        
        'As rotinas abaixo apenas ocorrem quando há alteração funcional
        'Abaixo: Rotinas de gravacao de treinamentos para o novo cargo
        'excluiProgramacao txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        
        '(O BLOCO DE GRAVAÇÃO DE TREINAMENTOS INTRODUTORIOS/OBRIGATORIOS FORM MOVIDOS PARA A CRIAÇÃO DA INTD)
        
        'GravaTreiPen txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        'If GeraIntr = "S" Then GravaTreiIntrodutorio txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
        'If GeraObri = "S" Then GravaTreiObrigatorio txtINTD(12), Val(Mid$(txtINTD(10), 1, 6))
    
    
    
    'Não Apto e Alteração Funcional -->
    ElseIf optINTD(1).Value = True And cboINTD(1).Text = "Não apto" Then
    
    'Apto e Capacitação Funcional -->
    ElseIf optINTD(1).Value = False And cboINTD(1).Text = "Apto" Then
    
    'Não Apto e Capacitação Funcional -->
    ElseIf optINTD(1).Value = False And cboINTD(1).Text = "Não apto" Then
    End If
    
    SqlINTD = "Update tbINTD set status = 'Fechado' Where codcoligada = '" & vCodcoligada & "' and codINTD = '" & Val(txtINTD(0)) & "'"
    rsINTD.Open SqlINTD, cnBanco
    Label14 = "Fechado"
    
    MsgBox "O processo de conclusão da INTD nº: " & txtINTD(0) & " foi realizado com sucesso!"
    AtualizaListview
End Function

Private Function pontosEscolar()
    pontosEscolar = False
    Dim vNumPDO As Integer
    Dim vStatusPDO As String
    Dim vDecisao As String
    Dim rsPDOColab As New ADODB.Recordset
    Dim SqlPDOColab As String
    
    SqlPDOColab = "Select a.cpf,a.codcolaborador,a.nomecolaborador,b.id,b.status,b.tipo,b.decisao from tbcolaboradores as a left join tbautorizacao as b on a.autorizacao = b.id where a.codcoligada = '" & vCodcoligada & "' and a.ativo  = 'S' and a.cpf = '" & txtINTD(12) & "'"
    rsPDOColab.Open SqlPDOColab, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not IsNull(rsPDOColab.Fields(3)) Then
        If rsPDOColab.RecordCount > 0 Then
            vNumPDO = rsPDOColab.Fields(3)
            If rsPDOColab.Fields(4) = "N" Or IsNull(rsPDOColab.Fields(4)) Then
                MsgBox "O PDO nº: " & Format(vNumPDO, "000000") & " esta em aberto para este " & rsPDOColab.Fields(5) & ". Aguarde tomada de decisão", vbCritical, "Atenção"
                rsPDOColab.Close
                Set rsPDOColab = Nothing
                Exit Function
            Else
                If Not IsNull(rsPDOColab.Fields(4)) Then
                    vStatusPDO = rsPDOColab.Fields(4)
                    vDecisao = rsPDOColab.Fields(6)
                End If
            End If
        End If
    End If
    rsPDOColab.Close
    Set rsPDOColab = Nothing
    
    If vStatusPDO <> "S" Then
        If Val(RemoveMask(Label16)) < MediaGlobal And Val(RemoveMask(Label16)) >= vAprovadoRest Then
            If vAdiRes = "N" Then
                If MsgBox("Escolaridade abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                    gravaSolicitacao txtINTD(12), "colaborador", RemoveMask(Label16), "INTD: " & txtINTD(0) & " - escolaridade abaixo do exigido pelo cargo: " & txtINTD(6), NomUsu
                    MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                End If
                Exit Function
            End If
        End If
        If Val(RemoveMask(Label16)) < vAprovadoRest Then
            If vAdiRep = "N" Then
                If MsgBox("Escolaridade abaixo da média. Usúario não privilégios para admitir o colaborador selecionado. Deseja gerar um PDO?", vbQuestion + vbYesNo, "SGCH") = vbYes Then
                    gravaSolicitacao txtINTD(12), "colaborador", RemoveMask(Label16), "INTD: " & txtINTD(0) & " - escolaridade abaixo do exigido pelo cargo: " & txtINTD(6), NomUsu
                    MsgBox "Foi gerado o PDO nº: " & Format(vPDO, "000000") & ". Aguarde tomada de decisão", vbInformation, "SGCH"
                End If
                Exit Function
            End If
        End If
    Else
        If Trim(vDecisao) <> "Aprovado" Then
            MsgBox "O PDO nº: " & Format(vNumPDO, "000000") & " NÃO FOI APROVADO ", vbCritical, "Atenção"
            
            'Fechar INTD
            SqlPDOColab = "Update tbINTD set status = 'Fechado' Where codcoligada = '" & vCodcoligada & "' and codINTD = '" & Val(txtINTD(0)) & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
            Label14 = "Fechado"
            
            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtINTD(12) & "' and codcolaborador = '" & txtINTD(3) & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
            
            MsgBox "O processo de conclusão da INTD nº: " & txtINTD(0) & " foi realizado com sucesso!"
            AtualizaListview
            Exit Function
        Else
            'Remover Numero de PDO da tabela de colaboradores
            SqlPDOColab = "Update tbColaboradores set autorizacao = Null Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtINTD(12) & "' and codcolaborador = '" & txtINTD(3) & "'"
            rsPDOColab.Open SqlPDOColab, cnBanco
        End If
        
    End If
    pontosEscolar = True
End Function

Private Sub avaliaEscolar()
    Dim rsAvEscolar As New ADODB.Recordset
    Dim SqlAvEscolar As String
    Dim PontosColabFor As Double
    Dim VerificaNull As Integer
    If optINTD(0).Value = True Then
        SqlAvEscolar = "select c.codmatriz,c.codescolaridade,c.pontuacao,b.cpf,b.tipo,b.codescolaridade,a.peso from tbescolaridade as a left join tbcolaboradoresesc as b on a.codescolaridade = b.codescolaridade and b.cpf = '" & txtINTD(12) & "' and b.tipo = 'colaborador' left join tbmatrizEsc as c on a.codescolaridade = c.codescolaridade and c.codmatriz = '" & Val(Mid$(txtINTD(10), 1, 6)) & "' where a.codcoligada = '" & vCodcoligada & "'"
    Else
        SqlAvEscolar = "select c.codmatriz,c.codescolaridade,c.pontuacao,b.cpf,b.tipo,b.codescolaridade,a.peso from tbescolaridade as a left join tbcolaboradoresesc as b on a.codescolaridade = b.codescolaridade and b.cpf = '" & txtINTD(12) & "' and b.tipo = 'colaborador' left join tbmatrizEsc as c on a.codescolaridade = c.codescolaridade and c.codmatriz = '" & Val(txtINTD(5)) & "' where a.codcoligada = '" & vCodcoligada & "'"
    End If
    rsAvEscolar.Open SqlAvEscolar, cnBanco, adOpenKeyset, adLockReadOnly
    PontosColabFor = 0
    VerificaNull = 0
    Do While Not rsAvEscolar.EOF
        If Not IsNull(rsAvEscolar.Fields(5)) Then VerificaNull = VerificaNull + 1
        If Not IsNull(rsAvEscolar.Fields(2)) Then PontosColabFor = rsAvEscolar.Fields(2)
        If Not IsNull(rsAvEscolar.Fields(3)) Then
            Exit Do
        End If
        rsAvEscolar.MoveNext
    Loop
    If VerificaNull = 0 Then PontosColabFor = 0
    If PontosColabFor < MediaGlobal Then
        Label16.ForeColor = &HC0&
    Else
        Label16.ForeColor = &H8000&
    End If
    Label16 = Format(PontosColabFor, "#,##0.00;(#,##0.00)") & " %"
    rsAvEscolar.Close
    Set rsAvEscolar = Nothing
End Sub

Private Sub registraExperiencia()
On Error GoTo Err
    'Se nao houver experiencia para essa na empresa para esse cargo o sistema insere
    Dim rsExperiencia As New ADODB.Recordset
    Dim SqlExperiencia As String
    SqlExperiencia = "Insert into tbColaboradoresExp(cpf,tipo,nomeempresa,codcargo,tempoexp,codcoligada) Values('" & txtINTD(12) & "','colaborador','" & NomeEmpresa & "'' (INTD: ''" & txtINTD(0) & "'')','" & txtINTD(11) & "','" & Format(periodoEmMeses, "000") & "''Meses','" & vCodcoligada & "')"
    rsExperiencia.Open SqlExperiencia, cnBanco
Err:
    'Se houver experiencia o sistema atualiza
    SqlExperiencia = "Update tbColaboradoresExp set tempoexp = '" & Format(periodoEmMeses, "00") & "'' Meses' where cpf ='" & txtINTD(12) & "' and nomeempresa = '" & NomeEmpresa & "'' (INTD: ''" & txtINTD(0) & "'')''" & "' and codcargo ='" & txtINTD(11) & "' and codcoligada = '" & vCodcoligada & "'"
    rsExperiencia.Open SqlExperiencia, cnBanco
End Sub

Private Sub alteraCargo()
'On Error GoTo Err
    Dim rsSalvarNovoCol As New ADODB.Recordset
    Dim SqlSalvarNovoCol As String
    Dim rsSalvarNovoCol1 As New ADODB.Recordset
    Dim SqlSalvarNovoCol1 As String
    Dim vSequencia As Integer
    
    SqlSalvarNovoCol = "select sequencia from tbColaboradoreshist where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtINTD(12) & "'"
    rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco, adOpenKeyset, adLockReadOnly
    If Not rsSalvarNovoCol.EOF Then
        rsSalvarNovoCol.MoveLast
        vSequencia = rsSalvarNovoCol.Fields(0) + 1
    Else
        vSequencia = 1
    End If
    rsSalvarNovoCol.Close
    Set rsSalvarNovoCol = Nothing

    SqlSalvarNovoCol = "Update tbColaboradoreshist set ativo = 'N', datasai = '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "' Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtINTD(12) & "' and tipo = 'colaborador' and ativo = 'S'"
    rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
    SqlSalvarNovoCol1 = "Insert into tbColaboradoreshist(cpf,codmatriz,data,ativo,sequencia,tipo,codcoligada) Values(REPLICATE('0', 11 - Len(" & txtINTD(12) & ")) + RTrim(" & txtINTD(12) & "), " & Val(txtINTD(5)) & ", '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "', 'S','" & vSequencia & "', 'colaborador','" & vCodcoligada & "')"
    '                                                                                             Values(REPLICATE('0', 11 - Len(" & txtINTD(12) & ")) + RTrim(" & txtINTD(12) & ")


'    SqlSalvarNovoCol = "Update tbColaboradoreshist set ativo = 'N', datasai = CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))) Where codcoligada = '" & vCodcoligada & "' and cpf = '" & txtINTD(12) & "' and tipo = 'colaborador' and ativo = 'S'"
'    rsSalvarNovoCol.Open SqlSalvarNovoCol, cnBanco
        
'    SqlSalvarNovoCol1 = "Insert into tbColaboradoreshist(cpf,codmatriz,data,ativo,sequencia,tipo,codcoligada) Values(REPLICATE('0', 11 - Len(" & txtINTD(12) & ")) + RTrim(" & txtINTD(12) & "), " & Val(txtINTD(5)) & ", CONVERT(DATETIME, FLOOR(CONVERT(FLOAT(24), GETDATE()))), 'S','" & vSequencia & "', 'colaborador','" & vCodcoligada & "')"
'    '                                                                                             Values(REPLICATE('0', 11 - Len(" & txtINTD(12) & ")) + RTrim(" & txtINTD(12) & ")
    
    rsSalvarNovoCol1.Open SqlSalvarNovoCol1, cnBanco
    
    If vIntegra = "S" Then
        ateraCargoTotvs txtINTD(3), txtCons(8)
    End If
    Exit Sub
Err:
    MsgBox "Ocorreu um erro ao tentar inserir o colaborador no novo cargo. Informe ao administrador do sistema", vbCritical
End Sub

Private Sub pontuaHabilidade(vMatriz As Integer)
On Error Resume Next
    Dim rsPHabilidades As New ADODB.Recordset
    Dim SqlPHabilidades As String
    SqlPHabilidades = "Select * from tbColaboradoresHab where codcoligada = '" & vCodcoligada & "'"
    rsPHabilidades.Open SqlPHabilidades, cnBanco, adOpenKeyset, adLockOptimistic
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems.Item(X).Selected = True
        rsPHabilidades.AddNew
        rsPHabilidades.Fields(0) = txtINTD(12) 'CPF
        rsPHabilidades.Fields(1) = "colaborador" 'Tipo
        rsPHabilidades.Fields(2) = Val(ListView2.ListItems.Item(X)) 'Código da Habilidade
        rsPHabilidades.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(3) 'Pontuação
        rsPHabilidades.Fields(4) = Val(txtINTD(5)) 'Código da matriz
        rsPHabilidades.Fields(5) = vCodcoligada 'Código da coligada
    Next
    rsPHabilidades.Update
    rsPHabilidades.Close
    Set rsPHabilidades = Nothing
End Sub

'AS 4 ROTINAS ABAIXO SAO RESPONSAVEIS POR GRAVAR TODOS OS TREINAMENTOS
'DO NOVO COLABORADOR NA TABELA TBPENDENTESCUR
'TAIS ROTINAS DEVERAO SER GLOBALIZADAS
'Deixar GLOBAL as seguintes rotinas listadas abaixo:
'excluirProgramacao
'GravarTreiPen
'ApagaExcesso
'GeraIntr
'GeraObri
Private Sub excluiProgramacao(vCPF As String, vMatriz As Integer)
    ' Rotina deleta toda a programação "Agendada ou Pendente" se o
    ' colaborador sofrer alteração de cargo
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    sqlDeletar = "Delete from tbPendentesCur where tbPendentesCur.codcoligada = '" & vCodcoligada & "' and tbPendentesCur.cpf = '" & vCPF & "' and status = 'Pendente' and codmatriz <> '" & vMatriz & "' or tbPendentesCur.cpf = '" & vCPF & "' and status = 'Agendado' and codmatriz <> '" & vMatriz & "'"
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

Private Sub GravaTreiPen(vCPF As String, vMatriz As Integer)
    'On Error Resume Next
    Dim rsGravaTreiPen As New ADODB.Recordset
    Dim SqlGravaTreiPen As String
    Dim rsPendentesCur As New ADODB.Recordset
    Dim SqlPendentesCur As String
    Dim contaID As Integer

    SqlGravaTreiPen = "Select a.codmatriz,a.codtreinamento,b.codtreinamento,b.cpf,a.codnivel from tbmatrizcur as a left join tbcolaboradorescur as b on a.codtreinamento = b.codtreinamento and b.codnivel >= a.codnivel and b.tipo = 'colaborador' and b.cpf = '" & vCPF & "' where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & vMatriz & "' order by a.codtreinamento"
    rsGravaTreiPen.Open SqlGravaTreiPen, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlPendentesCur = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id"
    rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not rsPendentesCur.EOF Then
        rsPendentesCur.MoveLast
        contaID = rsPendentesCur.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsPendentesCur.Close
    Set rsPendentesCur = Nothing
    
    While Not rsGravaTreiPen.EOF
        If IsNull(rsGravaTreiPen.Fields(2)) Then
            SqlPendentesCur = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and codtreinamento= '" & rsGravaTreiPen.Fields(1) & "' order by id"
            rsPendentesCur.Open SqlPendentesCur, cnBanco, adOpenKeyset, adLockOptimistic
            If rsPendentesCur.RecordCount = 0 Then
                rsPendentesCur.AddNew
                rsPendentesCur.Fields(0) = vCPF
                rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                rsPendentesCur.Fields(4) = "S"
                rsPendentesCur.Fields(5) = contaID
                rsPendentesCur.Fields(6) = "Pendente"
                rsPendentesCur.Fields(7) = 0
                If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(12) = 0 Else rsPendentesCur.Fields(12) = rsGravaTreiPen.Fields(4)
                rsPendentesCur.Fields(14) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            Else
                If rsPendentesCur.Fields(4) = "N" Then
                    rsPendentesCur.AddNew
                    rsPendentesCur.Fields(0) = vCPF
                    rsPendentesCur.Fields(1) = rsGravaTreiPen.Fields(0)
                    rsPendentesCur.Fields(2) = rsGravaTreiPen.Fields(1)
                    rsPendentesCur.Fields(4) = "S"
                    rsPendentesCur.Fields(5) = contaID
                    rsPendentesCur.Fields(6) = "Pendente"
                    rsPendentesCur.Fields(7) = 0
                    If IsNull(rsGravaTreiPen.Fields(4)) Then rsPendentesCur.Fields(12) = 0 Else rsPendentesCur.Fields(12) = rsGravaTreiPen.Fields(4)
                    rsPendentesCur.Fields(14) = vCodcoligada 'Codigo da coligada
                    contaID = contaID + 1
                Else
                End If
            End If
            rsPendentesCur.Update
            rsPendentesCur.Close
        End If
        rsGravaTreiPen.MoveNext
    Wend
    Set rsPendentesCur = Nothing
    
    rsGravaTreiPen.Close
    Set rsGravaTreiPen = Nothing
End Sub

Private Sub GravaTreiIntrodutorio(vCPF As String, vMatriz As Integer)
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiInt As New ADODB.Recordset
    Dim SqlSelecionaTreiInt As String
    Dim rsGravaTreiInt As New ADODB.Recordset
    Dim SqlGravaTreiInt As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & vMatriz & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
    
'    If ListView5.ListItems.Count > 1 Then
'        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codsetor = '" & rsAchaSetor.Fields(0) & "'"
'    Else
        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
 '   End If
    
    rsSelecionaTreiInt.Open SqlSelecionaTreiInt, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao from tbPendentesCur where codcoligada = '" & vCodcoligada & "'"
    rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiInt.EOF Then
        rsGravaTreiInt.MoveLast
        contaID = rsGravaTreiInt.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiInt.Close
    Set rsGravaTreiInt = Nothing
    
    While Not rsSelecionaTreiInt.EOF
        SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and cpf = '" & vCPF & "' and codtreinamento ='" & rsSelecionaTreiInt.Fields(0) & "'"
        rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiInt.RecordCount = 0 Then
            rsGravaTreiInt.AddNew
            rsGravaTreiInt.Fields(0) = vCPF
            rsGravaTreiInt.Fields(1) = vMatriz
            rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
            rsGravaTreiInt.Fields(4) = "S"
            rsGravaTreiInt.Fields(5) = contaID
            rsGravaTreiInt.Fields(6) = "Pendente"
            rsGravaTreiInt.Fields(7) = 0
            rsGravaTreiInt.Fields(8) = 0
            rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiInt.Fields(4) = "N" Then
                rsGravaTreiInt.AddNew
                rsGravaTreiInt.Fields(0) = vCPF
                rsGravaTreiInt.Fields(1) = vMatriz
                rsGravaTreiInt.Fields(2) = rsSelecionaTreiInt.Fields(0)
                rsGravaTreiInt.Fields(4) = "S"
                rsGravaTreiInt.Fields(5) = contaID
                rsGravaTreiInt.Fields(6) = "Pendente"
                rsGravaTreiInt.Fields(7) = 0
                rsGravaTreiInt.Fields(8) = 0
                rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiInt.Update
        rsGravaTreiInt.Close
        rsSelecionaTreiInt.MoveNext
    Wend
    Set rsGravaTreiInt = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiInt.Close
    Set rsSelecionaTreiInt = Nothing
End Sub

Private Sub GravaTreiObrigatorio(vCPF As String, vMatriz As Integer)
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiObr As New ADODB.Recordset
    Dim SqlSelecionaTreiObr As String
    Dim rsGravaTreiObr As New ADODB.Recordset
    Dim SqlGravaTreiObr As String
    Dim contaID As Integer
    
    'LOCALIZAR SETOR DO COLABORADOR
    SqlAchaSetor = "select a.codsetor from tbsetores as a inner join tbmatriz as b on a.codcoligada = '" & vCodcoligada & "' and a.codsetor = b.codsetor where b.codmatriz = '" & vMatriz & "'"
    rsAchaSetor.Open SqlAchaSetor, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    
    SqlGravaTreiObr = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada from tbPendentesCur where codcoligada = '" & vCodcoligada & "'"
    rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiObr.EOF Then
        rsGravaTreiObr.MoveLast
        contaID = rsGravaTreiObr.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiObr.Close
    Set rsGravaTreiObr = Nothing
    
    While Not rsSelecionaTreiObr.EOF
        SqlGravaTreiObr = "Select a.cpf,a.codmatriz,a.codtreinamento,a.codprogramacao,a.ativo,a.id,a.status,a.tipoprogramacao,a.codnivel,a.codcoligada from tbPendentesCur as a  left join tbTreinamentosNiv as b on a.codnivel = b.codnivel where a.cpf = '" & vCPF & "' and a.codtreinamento ='" & rsSelecionaTreiObr.Fields(0) & "'"
        rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
        If rsGravaTreiObr.RecordCount = 0 Then
            rsGravaTreiObr.AddNew
            rsGravaTreiObr.Fields(0) = vCPF
            rsGravaTreiObr.Fields(1) = vMatriz
            rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
            rsGravaTreiObr.Fields(4) = "S"
            rsGravaTreiObr.Fields(5) = contaID
            rsGravaTreiObr.Fields(6) = "Pendente"
            rsGravaTreiObr.Fields(7) = 0
            rsGravaTreiObr.Fields(8) = 0
            rsGravaTreiObr.Fields(9) = vCodcoligada ' Codigo da coligada
            contaID = contaID + 1
        Else
            If rsGravaTreiObr.Fields(4) = "N" Then
                rsGravaTreiObr.AddNew
                rsGravaTreiObr.Fields(0) = vCPF
                rsGravaTreiObr.Fields(1) = vMatriz
                rsGravaTreiObr.Fields(2) = rsSelecionaTreiObr.Fields(0)
                rsGravaTreiObr.Fields(4) = "S"
                rsGravaTreiObr.Fields(5) = contaID
                rsGravaTreiObr.Fields(6) = "Pendente"
                rsGravaTreiObr.Fields(7) = 0
                rsGravaTreiObr.Fields(8) = 0
                rsGravaTreiObr.Fields(9) = vCodcoligada ' Codigo da coligada
                contaID = contaID + 1
            End If
        End If
        rsGravaTreiObr.Update
        rsGravaTreiObr.Close
        rsSelecionaTreiObr.MoveNext
    Wend
    Set rsGravaTreiObr = Nothing
    
    rsAchaSetor.Close
    Set rsAchaSetor = Nothing
    
    rsSelecionaTreiObr.Close
    Set rsSelecionaTreiObr = Nothing
End Sub

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
            'If ListView2.ListItems.Count >= 16 Then
            '    If txtLeft < 11000 Then .Left = txtLeft + 305 Else .Left = txtLeft - 140
            'Else
                .Left = 9580
            'End If
            .Top = txtTop + 3000
            .Width = 615
            .Height = ListView2.SelectedItem.Height - 9
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
    cboINTD(0).SetFocus
    'ListView2.SetFocus
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

Private Sub comporCombosTotvs()
    Dim X As Integer
    CompoeComboTotvs Combo(9), "PFUNCAO", "codigo", "nome"
End Sub

Private Sub comporControlesTotvs()
    On Error Resume Next
    Dim rsContrTotvs As New ADODB.Recordset
    Dim SqlContrTotvs As String
        
    SqlContrTotvs = "select * from tbColaboradoresIntTotvs where codcoligada = '" & vCodcoligada & "' and id = '" & Val(Label17) & "'"
    rsContrTotvs.Open SqlContrTotvs, cnBanco, adOpenKeyset, adLockReadOnly
    txtCons(8) = rsContrTotvs.Fields(10)
    txtCons_KeyDown 8, 13, 8
    rsContrTotvs.Close
    Set rsContrTotvs = Nothing
End Sub

Private Sub txtCons_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 8
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            If txtCons(8) <> "" Then CarregaComboTotvs "PFUNCAO", "CODIGO", txtCons(8).Text, Combo(9).Text, Index, "nome"
        End If
    End Select
End Sub

Private Sub Combo_Click(Index As Integer)
    Select Case Index
    Case 9
        AchaComboTotvs Combo(Index), "PFUNCAO", "CODIGO", Index, "nome"
    End Select
End Sub

