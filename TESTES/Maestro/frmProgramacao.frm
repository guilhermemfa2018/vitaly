VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgramacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programação de Treinamento"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmProgramacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   12
      Left            =   1320
      TabIndex        =   25
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
      MICON           =   "frmProgramacao.frx":0CCA
      PICN            =   "frmProgramacao.frx":0CE6
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
      Index           =   13
      Left            =   720
      TabIndex        =   66
      Tag             =   "Imprimir"
      ToolTipText     =   "Imprimir"
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
      MICON           =   "frmProgramacao.frx":19C0
      PICN            =   "frmProgramacao.frx":19DC
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
      Index           =   11
      Left            =   120
      TabIndex        =   24
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
      MICON           =   "frmProgramacao.frx":26B6
      PICN            =   "frmProgramacao.frx":26D2
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
      Caption         =   "Código do modelo da avaliação de eficácia "
      Height          =   615
      Left            =   4560
      TabIndex        =   61
      Top             =   8760
      Width           =   3735
      Begin ACTIVESKINLibCtl.SkinLabel Label25 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":33AC
         TabIndex        =   84
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CheckBox Check1 
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
      TabIndex        =   12
      Tag             =   "Informa se será considerado a avaliação de eficácia para este treinamento"
      ToolTipText     =   "Informa se será considerado a avaliação de eficácia para este treinamento"
      Top             =   3120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Frame Frame8 
      Caption         =   "Avaliação do treinamento "
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
      Left            =   8760
      TabIndex        =   33
      Top             =   3120
      Width           =   2535
      Begin VB.CommandButton cmdCAD 
         Caption         =   "Avaliar todos"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   2055
      End
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
         TabIndex        =   34
         Top             =   240
         Width           =   1935
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   31
      Tag             =   "Código da programação"
      ToolTipText     =   "Código da programação"
      Top             =   4200
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7858
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Colaboradores/candidatos"
      TabPicture(0)   =   "frmProgramacao.frx":3402
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(2)=   "ListView1"
      Tab(0).Control(3)=   "txtProgTrei(10)"
      Tab(0).Control(4)=   "txtProgTrei(9)"
      Tab(0).Control(5)=   "cmdCadastro(3)"
      Tab(0).Control(6)=   "cmdCadastro(2)"
      Tab(0).Control(7)=   "cmdCadastro(10)"
      Tab(0).Control(8)=   "cmdCAD(3)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Instrutores"
      TabPicture(1)   =   "frmProgramacao.frx":341E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ListView2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtProgTrei(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtProgTrei(6)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cboProgTrei(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboProgTrei(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdCAD(4)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdCadastro(6)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdCadastro(7)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdCadastro(8)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdCadastro(9)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Avaliador"
      TabPicture(2)   =   "frmProgramacao.frx":343A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).Control(1)=   "Frame12"
      Tab(2).ControlCount=   2
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   9
         Left            =   1920
         TabIndex        =   22
         Tag             =   "Excluir Instrutor"
         ToolTipText     =   "Excluir Instrutor"
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
         MICON           =   "frmProgramacao.frx":3456
         PICN            =   "frmProgramacao.frx":3472
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
         Left            =   1320
         TabIndex        =   21
         Tag             =   "Editar Instrutor"
         ToolTipText     =   "Editar Instrutor"
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
         MICON           =   "frmProgramacao.frx":414C
         PICN            =   "frmProgramacao.frx":4168
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
         TabIndex        =   20
         Tag             =   "Novo Instrutor"
         ToolTipText     =   "Novo Instrutor"
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
         MICON           =   "frmProgramacao.frx":4E42
         PICN            =   "frmProgramacao.frx":4E5E
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
         Left            =   120
         TabIndex        =   19
         Tag             =   "Incluir Instrutor"
         ToolTipText     =   "Incluir Instrutor"
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
         MICON           =   "frmProgramacao.frx":5B38
         PICN            =   "frmProgramacao.frx":5B54
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton cmdCAD 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   7920
         TabIndex        =   71
         Tag             =   "Pesquisar"
         ToolTipText     =   "Pesquisar"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdCAD 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   -67200
         TabIndex        =   69
         Tag             =   "Pesquisar"
         ToolTipText     =   "Pesquisar"
         Top             =   600
         Width           =   375
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   495
         Index           =   10
         Left            =   -73200
         TabIndex        =   14
         Tag             =   "Avaliar colaborador"
         ToolTipText     =   "Avaliar colaborador"
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         BTYPE           =   8
         TX              =   "chameleonButton3"
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
         MICON           =   "frmProgramacao.frx":682E
         PICN            =   "frmProgramacao.frx":684A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   2
         Left            =   -74280
         TabIndex        =   58
         Tag             =   "Excluir colaborador"
         ToolTipText     =   "Excluir colaborador"
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
         MICON           =   "frmProgramacao.frx":7524
         PICN            =   "frmProgramacao.frx":7540
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
         Left            =   -74880
         TabIndex        =   60
         Tag             =   "Incluir colaborador"
         ToolTipText     =   "Incluir colaborador"
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
         MICON           =   "frmProgramacao.frx":821A
         PICN            =   "frmProgramacao.frx":8236
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame12 
         Caption         =   "Observação "
         Height          =   2895
         Left            =   -74880
         TabIndex        =   56
         Top             =   1440
         Width           =   10935
         Begin VB.TextBox txtProgTrei 
            Height          =   2535
            Index           =   8
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   57
            Tag             =   "Observação referente ao treinamento"
            ToolTipText     =   "Observação referente ao treinamento"
            Top             =   240
            Width           =   10695
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Dados do Avaliador"
         Height          =   975
         Left            =   -74880
         TabIndex        =   48
         Top             =   360
         Width           =   10935
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   285
            Left            =   8160
            TabIndex        =   55
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   56623105
            CurrentDate     =   40604
         End
         Begin VB.TextBox txtProgTrei 
            Height          =   285
            Index           =   11
            Left            =   1680
            TabIndex        =   51
            Top             =   480
            Width           =   6255
         End
         Begin VB.ComboBox cboProgTrei 
            Height          =   315
            Index           =   2
            ItemData        =   "frmProgramacao.frx":8F10
            Left            =   120
            List            =   "frmProgramacao.frx":8F1D
            TabIndex        =   50
            Text            =   "Instrutor"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Data da avaliação de eficácia:"
            Height          =   255
            Left            =   8160
            TabIndex        =   54
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Responsável:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   "Nome:"
            Height          =   255
            Left            =   1680
            TabIndex        =   49
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtProgTrei 
         Height          =   285
         Index           =   9
         Left            =   -74880
         TabIndex        =   45
         Tag             =   "CPF do colaborador"
         ToolTipText     =   "CPF do colaborador"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtProgTrei 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -73440
         TabIndex        =   44
         Tag             =   "Nome do colaborador"
         ToolTipText     =   "Nome do colaborador"
         Top             =   600
         Width           =   6135
      End
      Begin VB.ComboBox cboProgTrei 
         Height          =   315
         Index           =   1
         ItemData        =   "frmProgramacao.frx":8F41
         Left            =   8640
         List            =   "frmProgramacao.frx":8F4E
         TabIndex        =   17
         Tag             =   "Tipo de aula ministrada pelo instrutor"
         Text            =   "Teórico-Práticas"
         ToolTipText     =   "Tipo de aula ministrada pelo instrutor"
         Top             =   630
         Width           =   2055
      End
      Begin VB.ComboBox cboProgTrei 
         Height          =   315
         Index           =   0
         ItemData        =   "frmProgramacao.frx":8F78
         Left            =   120
         List            =   "frmProgramacao.frx":8F82
         TabIndex        =   39
         Tag             =   "Origem do treinamento"
         Text            =   "Interno"
         ToolTipText     =   "Origem do treinamento"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtProgTrei 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   15
         Tag             =   "Registro do instrutor responsável pelo curso/treinamento"
         ToolTipText     =   "Registro do instrutor responsável pelo curso/treinamento"
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtProgTrei 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   3600
         TabIndex        =   16
         Tag             =   "Nome do instrutor responsável pelo curso/treinamento"
         ToolTipText     =   "Nome do instrutor responsável pelo curso/treinamento"
         Top             =   660
         Width           =   4215
      End
      Begin VB.Frame Frame9 
         Caption         =   "Identificador"
         Height          =   615
         Left            =   3120
         TabIndex        =   35
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
         Begin VB.Label Label17 
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
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2655
         Left            =   120
         TabIndex        =   18
         Tag             =   "Instrutores responsáveis por ministrar o treinamento"
         ToolTipText     =   "Instrutores responsáveis por ministrar o treinamento"
         Top             =   1680
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4683
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   32
         Tag             =   "Colaboradores participantes do treinamento"
         ToolTipText     =   "Colaboradores participantes do treinamento"
         Top             =   1680
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4683
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         Caption         =   "CPF:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   -73440
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "Tipo aula:"
         Height          =   255
         Left            =   8640
         TabIndex        =   41
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Origem:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Registro:"
         Height          =   255
         Left            =   2160
         TabIndex        =   38
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Nome:"
         Height          =   255
         Left            =   3600
         TabIndex        =   37
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "    Determinar avaliação de eficácia"
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
      TabIndex        =   29
      Top             =   3120
      Width           =   8535
      Begin ACTIVESKINLibCtl.SkinLabel Label10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":8F98
         TabIndex        =   83
         Top             =   360
         Width           =   3135
      End
      Begin MAESTRO.chameleonButton cmdCadastro 
         Height          =   615
         Index           =   14
         Left            =   7800
         TabIndex        =   68
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   8
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
         MICON           =   "frmProgramacao.frx":9038
         PICN            =   "frmProgramacao.frx":9054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkProgTrei 
         Caption         =   "Outro:"
         Height          =   255
         Index           =   4
         Left            =   5040
         TabIndex        =   67
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox chkProgTrei 
         Caption         =   "Prova técnica"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   65
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkProgTrei 
         Caption         =   "Supervisão"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   64
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkProgTrei 
         Caption         =   "Teste"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkProgTrei 
         Caption         =   "Auditoria"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtProgTrei 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   13
         Tag             =   "Nome de outro método de avaliação de eficácia"
         ToolTipText     =   "Nome de outro método de avaliação de eficácia"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label24 
         Height          =   255
         Left            =   7320
         TabIndex        =   59
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Treinamento "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   11175
      Begin ACTIVESKINLibCtl.SkinLabel Text9 
         Height          =   255
         Left            =   9960
         OleObjectBlob   =   "frmProgramacao.frx":9D2E
         TabIndex        =   89
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text8 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "frmProgramacao.frx":9D96
         TabIndex        =   88
         Top             =   480
         Width           =   2535
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text7 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "frmProgramacao.frx":9DFC
         TabIndex        =   87
         Top             =   480
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text6 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmProgramacao.frx":9E60
         TabIndex        =   86
         Top             =   480
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel Text5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":9ED8
         TabIndex        =   85
         Top             =   480
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   255
         Left            =   9960
         OleObjectBlob   =   "frmProgramacao.frx":9F42
         TabIndex        =   82
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "frmProgramacao.frx":9FBC
         TabIndex        =   81
         Top             =   240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   6600
         OleObjectBlob   =   "frmProgramacao.frx":A024
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   1200
         OleObjectBlob   =   "frmProgramacao.frx":A092
         TabIndex        =   79
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":A0FA
         TabIndex        =   78
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCAD 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   52
         Tag             =   "Pesquisar"
         ToolTipText     =   "Pesquisar"
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Agendado para"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8280
      TabIndex        =   23
      Top             =   120
      Width           =   3015
      Begin VB.Frame Frame5 
         Caption         =   "Datas: Início/Término "
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2775
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   285
            Left            =   1425
            TabIndex        =   9
            Tag             =   "Data de término do curso/treinamento"
            ToolTipText     =   "Data de término do curso/treinamento"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   56623105
            CurrentDate     =   40591
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Tag             =   "Data de início do curso/treinamento"
            ToolTipText     =   "Data de início do curso/treinamento"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   56623105
            CurrentDate     =   40591
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Horário: Início/Término"
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   2775
         Begin MSMask.MaskEdBox mskProgTrei 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Tag             =   "Hora de início do curso/treinamento"
            ToolTipText     =   "Hora de início do curso/treinamento"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskProgTrei 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   11
            Tag             =   "Hora de término do curso/treinamento"
            ToolTipText     =   "Hora de término do curso/treinamento"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "##:##:##"
            PromptChar      =   "_"
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados da Programação "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtProgTrei 
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   4
         Tag             =   "Local do treinamento"
         ToolTipText     =   "Local do treinamento"
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtProgTrei 
         Height          =   285
         Index           =   1
         Left            =   2640
         TabIndex        =   3
         Tag             =   "Entidade que irá realizar o treinamento"
         ToolTipText     =   "Entidade que irá realizar o treinamento"
         Top             =   480
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Tag             =   "Data da programação"
         ToolTipText     =   "Data da programação"
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56623105
         CurrentDate     =   40591
      End
      Begin VB.TextBox txtProgTrei 
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
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   5280
         OleObjectBlob   =   "frmProgramacao.frx":A166
         TabIndex        =   75
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2640
         OleObjectBlob   =   "frmProgramacao.frx":A1D0
         TabIndex        =   74
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmProgramacao.frx":A240
         TabIndex        =   73
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmProgramacao.frx":A2A8
         TabIndex        =   72
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Responsável "
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
         TabIndex        =   30
         Top             =   840
         Width           =   7815
         Begin VB.TextBox txtProgTrei 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   1200
            TabIndex        =   6
            Tag             =   "Nome do responsável pela programação do curso/treinamento"
            ToolTipText     =   "Nome do responsável pela programação do curso/treinamento"
            Top             =   480
            Width           =   6015
         End
         Begin VB.TextBox txtProgTrei 
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Tag             =   "Identificação do responsável pela programação do curso/treinamento"
            ToolTipText     =   "Identificação do responsável pela programação do curso/treinamento"
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   1200
            OleObjectBlob   =   "frmProgramacao.frx":A314
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmProgramacao.frx":A37C
            TabIndex        =   76
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdCAD 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   7320
            TabIndex        =   7
            Tag             =   "Pesquisar"
            ToolTipText     =   "Pesquisar"
            Top             =   480
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Status "
      Enabled         =   0   'False
      Height          =   735
      Left            =   10200
      TabIndex        =   42
      Top             =   8640
      Width           =   1095
      Begin VB.CheckBox Check2 
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
         TabIndex        =   43
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProgramacao As New ADODB.Recordset
Private sqlProgramacao As String
Private rsProgInstrutor As New ADODB.Recordset
Private sqlProgInstrutor As String
Private rsCursoProg As New ADODB.Recordset
Private sqlCursoProg As String
Private Status As String
Private rsResponsavel As New ADODB.Recordset
Private SqlResponsavel As String
Private rsInstrutores As New ADODB.Recordset
Private SqlInstrutores As String
Private rsLocal As New ADODB.Recordset
Private validaCurso As Integer
Private vMsgProg As String
Private vAnotaFase As Integer
Private vCPFFase As String

Private Sub cboProgTrei_Click(Index As Integer)
    Select Case Index
    Case 0
        MontaMascara
    End Select
End Sub

Private Sub cmdCad_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridColaborador 0
        CarregaColaborador 0
    Case 1
        ChamaGridTreinamento
    Case 2
        avaliarTodos
    Case 3
        ChamaGridColaborador 3
        CarregaColaborador 3
    Case 4
        ChamaGridColaborador 2
        CarregaColaborador 2
    End Select
End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 2
        mobjMsg.Abrir "Deseja EXCLUIR esse colaborador da Programação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView1
            LimpaControlesColaboradorProg
        End If
    Case 3
        IncluirColaboradorProg
        LimpaControlesColaboradorProg
    Case 6
        IncluirInstrutorProg
        LimpaControlesInstrutorProg
    Case 7
        LimpaControlesInstrutorProg
    Case 8
        AlteraInstrutorProg
    Case 9
        mobjMsg.Abrir "Deseja EXCLUIR esse instrutor da Programação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView2
            LimpaControlesInstrutorProg
        End If
    Case 10
        If ListView1.ListItems.Count = 0 Then
            mobjMsg.Abrir "Não existem colaboradores a serem avaliados", Ok, critico, "Atenção"
            Exit Sub
        End If
        If Check1.Value = 1 Then Legenda = "Marcado" Else Legenda = ""
        MarcaPosicaoLV
        varGlobal2 = ListView1.ListItems.Item(Posicao) & txtProgTrei(0) & Text5
        Pesquisa = 1
        If txtProgTrei(0) <> "-" Then
            frmFichaAvaliacao.Show 1
        Else
            mobjMsg.Abrir "O colaborador/candidato apenas poderá ser avaliado após o treinamento ser agendado", Ok, critico, "Atenção"
        End If
        If Pesquisa = 1 Then AtualizaLV
    Case 11
        mobjMsg.Abrir "Deseja salvar os dados da Programação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            gravaLog "Código prog: " & txtProgTrei(0), "Instrutor: " & txtProgTrei(11), "Data ini: " & DTPicker2 & "Data fim: " & DTPicker3 & "Hara ini: " & mskProgTrei(0) & "Hara fim: " & mskProgTrei(1)
            Pesquisa = "0"
            'Unload Me
        End If
    Case 12
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Programação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = "0"
            Unload Me
            Set frmProgramacao = Nothing
        End If
    Case 13
        frmPrinterProg.Show 1
    Case 14
        frmAvaliacaoProg.Show 1
        'Label25 = vCodModeloAval
    End Select
End Sub

Private Sub avaliarTodos()
On Error Resume Next
    Dim rsAvAT As New ADODB.Recordset
    Dim SqlAT As String
    Dim rsAvaliarTodos As New ADODB.Recordset
    Dim SqlAvaliarTodos As String
    Dim X As Integer, Y As Integer
    
    cnBanco.BeginTrans
    
    SqlAT = "select * from tbavaliacao where tipo = 'AT'"
    rsAvAT.Open SqlAT, cnBanco, adOpenKeyset, adLockReadOnly
    For Y = 1 To ListView1.ListItems.Count
        ListView1.ListItems(Y).Selected = True
        rsAvAT.MoveFirst
        For X = 1 To rsAvAT.RecordCount
            SqlAvaliarTodos = "Insert into tbAvaliacaoTrei(codprogramacao,CPF,codavaliacao,pontuacao,codcoligada) Values('" & Val(txtProgTrei(0)) & "','" & ListView1.ListItems.Item(Y) & "','" & rsAvAT.Fields(0) & "',100,'" & vCodcoligada & "')"
            rsAvaliarTodos.Open SqlAvaliarTodos, cnBanco
            rsAvAT.MoveNext
        Next
        ListView1.SelectedItem.ListSubItems.Item(2) = "Aprovado"
    Next
    cnBanco.CommitTrans
    rsAvAT.Close
    Set rsAvAT = Nothing
    MediaTreinamento
    Exit Sub
Err:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
End Sub

Private Function VerificaConcluido()
    VerificaConcluido = False
    'On Error GoTo Err
    Dim Y As Integer
    Dim qtdconc As Integer
    Y = ListView1.ListItems.Count
    vqtdconc = 0
    For X = 1 To Y
        ListView1.ListItems(X).Selected = True
        If ListView1.SelectedItem.ListSubItems.Item(2) <> "-" Then vqtdconc = vqtdconc + 1
    Next
    If vqtdconc = Y Then
        VerificaConcluido = True
    End If
    Exit Function
End Function

Private Sub AtualizaLV()
    'On Error GoTo Err
    Dim Y As Integer
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X) = Mid(varGlobal2, 1, 11) Then
            ListView1.ListItems(X).Selected = True
            ListView1.SelectedItem.ListSubItems.Item(2) = vsituacao
            ListView1.SelectedItem.ListSubItems.Item(3) = vNota
            Exit For
        End If
    Next
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub Form_Activate()
    Status = Pesquisa
    If varGlobal <> "-" And Status = "novo" Or varGlobal <> "" And Status = "novo" Then LibControlesTreinamento
    If validaCurso = 1 Then
        mobjMsg.Abrir vMsgProg, Ok, critico, "Atenção"
        Unload Me
    End If
    'If MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = "Cancelado" Then Unload Me
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    AchaMarca
    If MeuLV.ListView1.ListItems.Count > 0 Then varGlobal = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
    If varGlobal = "-" Or varGlobal = "" Then Status = "novo" Else Status = Pesquisa
    listview_cabecalho
    SSTab1.Tab = 0
    validaCurso = 0
    If Status = "novo" Then
        If MeuLV.ListView1.ListItems.Count > 0 Then
            If MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) = "Cancelado" Then BloqueiaTudo
        End If
        LimpaControles
        Label17.Caption = "000001"
        CompoeControles
        ListaColabProg
        If validaCurso <> 1 Then ConfereFases
    ElseIf Status = "editar" Then
        txtProgTrei(0) = varGlobal
        cmdCadastro(10).UseGreyscale = False
        ResultPesq
        MediaTreinamento
        LibControlesTreinamento
        If StatusTrei = "Concluido" Then BloqueiaTudo
    End If
    configControles
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub ConfereFases()
On Error Resume Next
    varGlobal = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer
    Dim vCurso As String
    Y = MeuLV.ListView1.ListItems.Count
    vCurso = ""
    For X = 1 To Y
        MeuLV.ListView1.ListItems.Item(X).Selected = True
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            vAnotaFase = Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(11))
            vCPFFase = MeuLV.ListView1.ListItems.Item(X)
            If LocalFase = True Then
                validaCurso = 1
                vMsgProg = "Foram selecionados colaboradores que possuem fases anteriores de treinamentos ainda NÃO CONCLUIDAS. Favor conclui-las"
                Exit Sub
            Else
                validaCurso = 0
            End If
        End If
    Next
End Sub

Private Function LocalFase()
    Dim rsLocalFase As New ADODB.Recordset
    Dim sqlLocalFase As String
    sqlLocalFase = "select a.cpf,a.status,e.idGrFase from tbPendentesCur as a inner join tbtreinamentos as e on e.codtreinamento = a.codtreinamento where a.ativo = 'S' and a.status <> 'Concluido' and a.cpf = '" & vCPFFase & "'"
    rsLocalFase.Open sqlLocalFase, cnBanco, adOpenKeyset, adLockReadOnly
        
    Dim Z As Integer, K As Integer
    LocalFase = False
    If Not rsLocalFase.EOF Then Z = rsLocalFase.RecordCount Else Z = 0
    If Z <> 0 Then
        For K = 1 To Z
            If Not IsNull(rsLocalFase.Fields(2)) Then
                If Val(rsLocalFase.Fields(2)) = vAnotaFase - 1 And rsLocalFase.Fields(1) <> "Concluido" Then
                    LocalFase = True
                End If
            End If
            rsLocalFase.MoveNext
        Next
    End If
    rsLocalFase.Close
End Function

Private Sub BloqueiaTudo()
On Error Resume Next
    Dim X As Integer
    ListView1.Enabled = False
    ListView2.Enabled = False
    For X = 0 To 11
        'cmdCadastro(X).Enabled = False
        cmdCadastro(X).DragMode = 1
        cmdCadastro(X).UseGreyscale = True
    Next
    cmdCadastro(14).DragMode = 1
    cmdCadastro(14).UseGreyscale = True
    cmdCadastro(15).DragMode = 1
    cmdCadastro(15).UseGreyscale = True
    For X = 0 To 11
        txtProgTrei(X).Enabled = False
    Next
    For X = 0 To 4
        cmdCad(X).Enabled = False
    Next
    For X = 0 To 4
        chkProgTrei(X).Enabled = False
    Next
    DTPicker1.Enabled = False
    DTPicker2.Enabled = False
    DTPicker3.Enabled = False
    DTPicker4.Enabled = False
    mskProgTrei(0).Enabled = False
    mskProgTrei(1).Enabled = False
    Check1.Enabled = False
    Check2.Enabled = False
    cboProgTrei(2).Enabled = False
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "CPF", ListView1.Width / 10
    ListView1.ColumnHeaders.Add , , "Nome do colaborador", ListView1.Width / 3
    ListView1.ColumnHeaders.Add , , "Situação", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "Pontuação", ListView1.Width / 11
    
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 11
    ListView2.ColumnHeaders.Add , , "Origem", ListView2.Width / 11
    ListView2.ColumnHeaders.Add , , "Registro", ListView2.Width / 11
    ListView2.ColumnHeaders.Add , , "Nome", ListView2.Width / 2
    ListView2.ColumnHeaders.Add , , "Tipo de aula", ListView2.Width / 5
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub MarcaPosicaoLV()
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            ListView1.ListItems.Item(X).Selected = True
            Exit For
        End If
    Next
    Posicao = X
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    DTPicker3 = Date
    DTPicker4 = Date
    For X = 0 To txtProgTrei.Count - 1
        txtProgTrei(X) = ""
    Next
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    'txtProgTrei(0) = Format(GeraCodigo, "000000")
    txtProgTrei(0) = "-"
    Status = "novo"
End Sub

Private Sub LimpaControlesColaboradorProg()
    Dim X As Integer
    For X = 9 To 10
        txtProgTrei(X) = ""
    Next
End Sub

Private Sub LimpaControlesInstrutorProg()
    Dim X As Integer
    For X = 6 To 7
        txtProgTrei(X) = ""
    Next
    cboProgTrei(0) = "Interno"
    cboProgTrei(1) = "Teórico-Práticas"
End Sub

Private Sub LibControlesTreinamento()
    If Text5.Caption = "Código" Then
        Text5.Caption = ""
        Text6.Caption = ""
        Text7.Caption = ""
        Text8.Caption = ""
        Text9.Caption = ""
    End If
    Text5.Enabled = True
'    Text5.BorderStyle = 1
'    Text5.BackColor = &H80000005
'    Text6.BorderStyle = 1
'    Text6.BackColor = &H80000005
    cmdCad(1).Enabled = True
End Sub

Private Sub ResultPesq()
    CompoeControles
    ListaColabProg
    ListaInstProg
    Label17.Caption = Format(GeraCodigo1, "000000")
End Sub

Private Sub CompoeControles()
    Dim X As Integer, Y As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            Exit For
        End If
    Next
    If X > Y Then X = X - 1
    If Status = "editar" Then
        sqlProgramacao = "Select a.codprogramacao,a.dataprogramacao,a.entidade,a.local,a.codcolaborador,b.nomecolaborador,a.datainicio,a.datafim,a.horainicio,a.horafim,a.dae,a.metodo,a.metodooutro,a.nota,a.observacao,a.avaltipo,a.avalnome,a.avaldata,a.codmodelo,a.metodoA,a.metodoT,a.metodoS,a.metodoPT from tbProgramacao as a left join tbcolaboradores as b on b.codcolaborador=a.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and a.codprogramacao = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)) & "'"
        rsProgramacao.Open sqlProgramacao, cnBanco, adOpenKeyset, adLockReadOnly
    
        txtProgTrei(0).Text = Format(rsProgramacao.Fields(0), "000000")
        txtProgTrei(1).Text = rsProgramacao.Fields(2)
        txtProgTrei(2).Text = rsProgramacao.Fields(3)
        txtProgTrei(3).Text = rsProgramacao.Fields(4)
        If Not IsNull(rsProgramacao.Fields(5)) Then
            txtProgTrei(4).Text = rsProgramacao.Fields(5)
        Else
            txtProgTrei(3) = ""
        End If
        DTPicker1 = rsProgramacao.Fields(1)
        DTPicker2 = rsProgramacao.Fields(6)
        DTPicker3 = rsProgramacao.Fields(7)
        mskProgTrei(0) = rsProgramacao.Fields(8)
        mskProgTrei(1) = rsProgramacao.Fields(9)
        If rsProgramacao.Fields(10) = True Then
            Check1.Value = 1
            
            If rsProgramacao.Fields(19) = 0 Then chkProgTrei(0).Value = 0 Else chkProgTrei(0).Value = 1
            If rsProgramacao.Fields(20) = 0 Then chkProgTrei(1).Value = 0 Else chkProgTrei(1).Value = 1
            If rsProgramacao.Fields(21) = 0 Then chkProgTrei(2).Value = 0 Else chkProgTrei(2).Value = 1
            If rsProgramacao.Fields(22) = 0 Then chkProgTrei(3).Value = 0 Else chkProgTrei(3).Value = 1
            If rsProgramacao.Fields(11) = 0 Then
                chkProgTrei(4).Value = 0
            Else
                chkProgTrei(4).Value = 1
                txtProgTrei(5) = rsProgramacao.Fields(12)
            End If
        Else
            Check1.Value = 0
        End If
        txtProgTrei(8).Text = rsProgramacao.Fields(14)
        Label13.Caption = Format(rsProgramacao.Fields(13), "#,##00.00;(#,##0.00)") & "%"
        If Not IsNull(rsProgramacao.Fields(15)) Then cboProgTrei(2) = rsProgramacao.Fields(15)
        If Not IsNull(rsProgramacao.Fields(16)) Then txtProgTrei(11) = rsProgramacao.Fields(16)
        If Not IsNull(rsProgramacao.Fields(17)) Then DTPicker4 = rsProgramacao.Fields(17)
        If Not IsNull(rsProgramacao.Fields(18)) Then Label25 = rsProgramacao.Fields(18)
        vCodModeloAval = Val(Label25)
        rsProgramacao.Close
        Set rsProgramacao = Nothing
    End If
'    sqlCursoProg = "Select a.codtreinamento,a.nometreinamento,b.revisao,a.tipo,a.cargahoraria from tbtreinamentos as a left join  tbtreinamentosrev as b on b.codtreinamento = a.codtreinamento Where a.codtreinamento = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)) & "' order by b.revisao desc"
    
    'VERIFICA SE POSSUI CODIGO DE PROGRAMAÇÃO ***
    If MeuLV.ListView1.ListItems.Count > 0 Then
        'If MeuLV.ListView1.SelectedItem.ListSubItems.Item(6) <> "-" Then
        ''If MeuLV.ListView1.ListItems.Item(X) <> "-" Then
        '    sqlCursoProg = "Select a.codtreinamento,c.nometreinamento + ' ('+isnull(b.nomenivel,'-')+')',d.revisao,c.tipo,c.cargahoraria from tbPendentesCur as a left join tbTreinamentosNiv as b on a.codtreinamento = b.codtreinamento and a.codnivel = b.codnivel inner join tbtreinamentos as c on a.codtreinamento = c.codtreinamento left join tbtreinamentosrev as d on c.codtreinamento = d.codtreinamento Where a.cpf = '" & MeuLV.ListView1.ListItems.Item(X) & "' and a.codtreinamento = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)) & "' and a.status = '" & MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) & "' order by d.revisao desc"
        'Else
            sqlCursoProg = "Select a.codtreinamento,c.nometreinamento + ' ('+isnull(b.nomenivel,'-')+')',d.revisao,c.tipo,c.cargahoraria from tbPendentesCur as a left join tbTreinamentosNiv as b on a.codtreinamento = b.codtreinamento and a.codnivel = b.codnivel inner join tbtreinamentos as c on a.codtreinamento = c.codtreinamento left join tbtreinamentosrev as d on c.codtreinamento = d.codtreinamento Where a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)) & "' order by d.revisao desc"
        'End If
        '********************************************
        rsCursoProg.Open sqlCursoProg, cnBanco, adOpenKeyset, adLockReadOnly
        Text5.Caption = Format(rsCursoProg.Fields(0), "000000")
        Text6.Caption = rsCursoProg.Fields(1)
        If Not IsNull(rsCursoProg.Fields(2)) Then Text7.Caption = rsCursoProg.Fields(2) Else Text7.Caption = "-"
        Text8.Caption = rsCursoProg.Fields(3)
        Text9.Caption = Format(rsCursoProg.Fields(4), "000:00")
        rsCursoProg.Close
        Set rsCursoProg = Nothing
    End If
End Sub

Private Sub ListaColabProg()
On Error Resume Next
    varGlobal = MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)
    Dim rsListarProg As New ADODB.Recordset
    Dim SqlListarProg As String
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer, Z As Integer
    Dim vCurso As String
    Y = MeuLV.ListView1.ListItems.Count
    vCurso = ""
    ListView1.ListItems.Clear
    For X = 1 To Y
        MeuLV.ListView1.ListItems.Item(X).Selected = True
        If Status = "editar" Then
            If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
                StatusTrei = MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) 'Status do treinamento
                SqlListarProg = "select a.codprogramacao,a.cpf,b.nomecolaborador,a.situacao,a.nota from tbPendentesCur as a inner join tbcolaboradores as b on b.cpf = a.cpf and b.ativo = 'S' and a.ativo = 'S' where a.codprogramacao = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)) & "'"
                rsListarProg.Open SqlListarProg, cnBanco, adOpenKeyset, adLockReadOnly
                For Z = 1 To rsListarProg.RecordCount
                    Set ItemLst = ListView1.ListItems.Add(, , rsListarProg.Fields(1))
                    ItemLst.SubItems(1) = rsListarProg.Fields(2)
                    
                    If Check1.Value = 1 Then
                        If Not IsNull(rsListarProg.Fields(4)) And rsListarProg.Fields(4) <> 0 Then ItemLst.SubItems(3) = Format(rsListarProg.Fields(4), "#,##00.00;(#,##0.00)") & "%" Else ItemLst.SubItems(3) = "-"
                        If rsListarProg.Fields(4) >= MediaGlobal Then
                            If Not IsNull(rsListarProg.Fields(3)) Then ItemLst.SubItems(2) = "Aprovado" Else ItemLst.SubItems(2) = "-"
                        
                        ElseIf rsListarProg.Fields(4) >= vAprovadoRest And rsListarProg.Fields(4) < MediaGlobal Then
                            If Not IsNull(rsListarProg.Fields(3)) And rsListarProg.Fields(3) <> "-" And rsListarProg.Fields(4) <> 0 Then ItemLst.SubItems(2) = "Aprovado com restrição" Else ItemLst.SubItems(2) = "-"
                        Else
                            If Not IsNull(rsListarProg.Fields(3)) And rsListarProg.Fields(3) <> "-" And rsListarProg.Fields(4) <> 0 Then ItemLst.SubItems(2) = "Reprovado" Else ItemLst.SubItems(2) = "-"
                        End If
                    Else
                        If Not IsNull(rsListarProg.Fields(3)) And rsListarProg.Fields(3) <> "-" Then ItemLst.SubItems(2) = "Aprovado" Else ItemLst.SubItems(2) = "-"
                        ItemLst.SubItems(3) = "-"
                    End If
                    
                    rsListarProg.MoveNext
                Next
                rsListarProg.Close
                Set rsListarProg = Nothing
                Exit For
            End If
        ElseIf Status = "novo" Then
            If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
                StatusTrei = MeuLV.ListView1.SelectedItem.ListSubItems.Item(8) 'Status do treinamento
                If vCurso = "" Then
                    vCurso = MeuLV.ListView1.SelectedItem.ListSubItems.Item(4)
                Else
                    If vCurso <> MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) Then validaCurso = 1
                    vMsgProg = "A programação não pode ser realizada para cursos/treinamentos diferentes"
                End If
                Set ItemLst = ListView1.ListItems.Add(, , MeuLV.ListView1.ListItems.Item(X))
                ItemLst.SubItems(1) = MeuLV.ListView1.SelectedItem.ListSubItems.Item(1)
                ItemLst.SubItems(2) = "-"
                ItemLst.SubItems(3) = "-"
            End If
        End If
    Next
End Sub

Private Sub ListaInstProg()
    Dim ItemLst As ListItem
    Dim X As Integer
    SqlInstrutores = "select a.sequencia,a.origem,a.codcolaborador,a.nomeinstrutor,b.nomecolaborador,a.tipoaula from tbProgramacaoInstrutores as a left join tbcolaboradores as b on a.codcolaborador=b.codcolaborador where a.codcoligada = '" & vCodcoligada & "' and a.codprogramacao = '" & Val(MeuLV.ListView1.SelectedItem.ListSubItems.Item(6)) & "'Order by a.sequencia"
    rsInstrutores.Open SqlInstrutores, cnBanco, adOpenKeyset, adLockReadOnly
    X = 0
    While Not rsInstrutores.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsInstrutores.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsInstrutores.Fields(1)
        ItemLst.SubItems(2) = "" & rsInstrutores.Fields(2)
        If Val(rsInstrutores.Fields(2)) = 0 Then
            ItemLst.SubItems(3) = "" & rsInstrutores.Fields(3)
        Else
            ItemLst.SubItems(3) = "" & rsInstrutores.Fields(4)
        End If
        ItemLst.SubItems(4) = rsInstrutores.Fields(5)
        rsInstrutores.MoveNext
        X = X + 1
    Wend
    rsInstrutores.Close
    Set rsInstrutores = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub AchaMarca()
    Dim ItemLst As ListItem
    Dim Y As Integer, X As Integer
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Checked = True Then
            MeuLV.ListView1.ListItems.Item(X).Selected = True
            Exit For
        End If
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirProgramacao
    SqlGera = "Select top 1 * from tbProgramacao where codcoligada = '" & vCodcoligada & "' order by codprogramacao Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsProgramacao.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtProgTrei(0) = Format(GeraCodigo, "000000")
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharProgramacao
End Function

Private Sub AbrirProgramacao()
    sqlProgramacao = "Select * from tbProgramacao where codcoligada = '" & vCodcoligada & "' Order by codprogramacao"
    rsProgramacao.Open sqlProgramacao, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharProgramacao()
    rsProgramacao.Close
    Set rsProgramacao = Nothing
End Sub

Private Sub AbrirProgramacaoInstrutores()
    sqlProgInstrutor = "Select * from tbProgramacaoInstrutores where codcoligada = '" & vCodcoligada & "' Order by codprogramacao"
    rsProgInstrutor.Open sqlProgInstrutor, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharProgramacaoInstrutores()
    rsProgInstrutor.Close
    Set rsProgInstrutor = Nothing
End Sub

Private Function GeraCodigo1()
On Error GoTo Err
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    SqlGera = "Select top 1 * from tbProgramacaoInstrutores where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & Val(txtProgTrei(0)) & "' order by codprogramacao Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    AbrirProgramacaoInstrutores
    If rsProgInstrutor.RecordCount > 0 Then
        GeraCodigo1 = rsGeraCodigo.Fields(5) + 1
    Else
        GeraCodigo1 = 1
    End If
    Label17.Caption = Format(GeraCodigo1, "000000")
    FecharProgramacaoInstrutores
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    Exit Function
Err:
    GeraCodigo1 = 1
    Label17.Caption = Format(GeraCodigo1, "000000")
    Exit Function
End Function

Private Function GeraCodigoPen()
    Dim rsGeraCodigoPen As New ADODB.Recordset
    Dim SqlGeraCOdigoPen
    SqlGeraCOdigoPen = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' order by id"
    rsGeraCodigoPen.Open SqlGeraCOdigoPen, cnBanco, adOpenKeyset, adLockReadOnly
    
    If Not rsGeraCodigoPen.EOF Then
        rsGeraCodigoPen.MoveLast
        GeraCodigoPen = rsGeraCodigoPen.Fields(5) + 1
    Else
        GeraCodigoPen = 1
    End If
    rsGeraCodigoPen.Close
    Set rsGeraCodigoPen = Nothing
End Function

Private Sub ListView_DblClick(Index As Integer)
    AlteraInstrutorProg
End Sub

Private Sub ListView2_DblClick()
    If vEdi <> "N" Then
        AlteraInstrutorProg
    End If
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
        CarregaColaborador 4
    End If
End Sub

Private Sub txtProgTrei_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 0
        End If
    Case 6
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 2
        End If
    Case 9
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaColaborador 3
        End If
    End Select
Error:
    Exit Sub
End Sub

Private Sub CarregaColaborador(indice As Integer)
    Dim X As Integer
    If indice = 0 Then
        SqlInstrutores = "select  a.codcolaborador,a.nomecolaborador,d.nomedepartamento,e.nomesetor from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor where a.codcolaborador = '" & txtProgTrei(3) & "'"
        rsInstrutores.Open SqlInstrutores, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    If indice = 2 Then
        SqlInstrutores = "select a.codcolaborador,a.nomecolaborador,d.nomedepartamento,e.nomesetor from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor where a.codcolaborador = '" & txtProgTrei(6) & "'"
        rsInstrutores.Open SqlInstrutores, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    If indice = 3 Then
        SqlInstrutores = "select a.cpf,a.nomecolaborador,d.nomedepartamento,e.nomesetor from tbcolaboradores as a inner join tbcolaboradoreshist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf and b.ativo = 'S' inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbdepartamentos as d on c.coddepartamento=d.coddepartamento inner join tbsetores as e on c.codsetor = e.codsetor where a.cpf = '" & txtProgTrei(9) & "'"
        rsInstrutores.Open SqlInstrutores, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    If indice = 4 Then
        SqlInstrutores = "select a.codtreinamento,a.nometreinamento,b.revisao,a.tipo,a.cargahoraria from tbtreinamentos as a left join tbtreinamentosrev as b on b.codtreinamento = a.codtreinamento where a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = '" & Val(Text5) & "' order by b.revisao desc"
        rsInstrutores.Open SqlInstrutores, cnBanco, adOpenKeyset, adLockReadOnly
    End If
    
    
    If indice = 0 Then
        If rsInstrutores.RecordCount <= 0 Then
            If txtProgTrei(3).Text <> "000000" And txtProgTrei(3).Text <> "" Then mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtProgTrei(4) = ""
        Else
            txtProgTrei(3).Text = rsInstrutores.Fields(0)
            txtProgTrei(4).Text = rsInstrutores.Fields(1)
        End If
    End If
    If indice = 2 Then
        If rsInstrutores.RecordCount <= 0 Then
            If txtProgTrei(6).Text <> "000000" And txtProgTrei(6).Text <> "" Then mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtProgTrei(7) = ""
        Else
            txtProgTrei(6).Text = rsInstrutores.Fields(0)
            txtProgTrei(7).Text = rsInstrutores.Fields(1)
        End If
    End If
    If indice = 3 Then
        If rsInstrutores.RecordCount <= 0 Then
            If txtProgTrei(9).Text <> "000000" And txtProgTrei(9).Text <> "" Then mobjMsg.Abrir "Colaborador não cadastrado", Ok, critico, "Atenção"
            txtProgTrei(10) = ""
        Else
            txtProgTrei(9).Text = rsInstrutores.Fields(0)
            txtProgTrei(10).Text = rsInstrutores.Fields(1)
        End If
    End If
   
    
    If indice = 4 Then
        If rsInstrutores.RecordCount <= 0 Then
            If Text5.Text <> "000000" And Text5.Text <> "" Then mobjMsg.Abrir "Treinamento não cadastrado", Ok, critico, "Atenção"
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
        Else
            Text5.Text = Format(rsInstrutores.Fields(0), "000000")
            Text6.Text = rsInstrutores.Fields(1)
            If Not IsNull(rsInstrutores.Fields(2)) Then Text7.Text = rsInstrutores.Fields(2) Else Text7.Text = "-"
            Text8.Text = rsInstrutores.Fields(3)
            Text9.Text = Format(rsInstrutores.Fields(4), "000:00")
        End If
    End If
    rsInstrutores.Close
    Set rsInstrutores = Nothing
End Sub

Private Sub ChamaGridColaborador(indice As Integer)
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    If indice = 3 Then
        Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by nomecolaborador"
    Else
        Sqlp = "Select * from tbcolaboradores where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and tipo = 'colaborador' order by nomecolaborador"
    End If
    procnom = "nomecolaborador"
    If indice <> 3 Then
        campo = 3
        Campo1 = 1
    Else
        campo = 3
        Campo1 = 0
    End If
    Load F
    F.Caption = "Pesquisa de Colaborador"
    Pesquisa = frmProgramacao.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomecolaborador=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            If indice = 0 Then
                txtProgTrei(3).Text = rsLocal.Fields(1)
            End If
            If indice = 2 Then
                txtProgTrei(6).Text = rsLocal.Fields(1)
            End If
            If indice = 3 Then
                txtProgTrei(9).Text = rsLocal.Fields(0)
            End If
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub ChamaGridTreinamento()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select a.codtreinamento,a.nometreinamento,b.revisao,a.tipo,a.cargahoraria from tbtreinamentos as a left join tbtreinamentosrev as b on b.codtreinamento = a.codtreinamento where a.codcoligada = '" & vCodcoligada & "' and a.ativo = 'S' order by a.nometreinamento"
    procnom = "nometreinamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Treinamento"
    Pesquisa = frmProgramacao.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nometreinamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            Text5.Text = Format(rsLocal.Fields(0), "000000")
            Text6.Text = rsLocal.Fields(1)
            If Not IsNull(rsLocal.Fields(2)) Then Text7.Text = rsLocal.Fields(2) Else Text7.Text = "-"
            Text8.Text = rsLocal.Fields(3)
            Text9.Text = Format(rsLocal.Fields(4), "000:00")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub MontaMascara()
    If cboProgTrei(0) <> "Interno" Then
        txtProgTrei(6) = Format(0, "000000")
        txtProgTrei(6).Enabled = False
        txtProgTrei(7).Enabled = True
        txtProgTrei(7).BackColor = &H80000018
        If txtProgTrei(7) = "" Then txtProgTrei(7).Text = "Digite o nome do instrutor"
        cmdCad(4).Enabled = False
    ElseIf cboProgTrei(0) = "Interno" Then
        txtProgTrei(6).Enabled = True
        txtProgTrei(7).Enabled = False
        txtProgTrei(7).BackColor = &H80000005
        If txtProgTrei(6) <> "000000" And txtProgTrei(6) = "" Then
            txtProgTrei(6).Text = ""
            txtProgTrei(7).Text = ""
        End If
        cmdCad(4).Enabled = True
        CarregaColaborador 2
    End If
End Sub

Private Sub IncluirColaboradorProg()
    If ValidaCampoCol = False Then Exit Sub
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView1.ListItems.Item(X).Selected = True
            If ListView1.ListItems.Item(X) = Me.txtProgTrei(9) Then
                txtProgTrei(9).Text = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtProgTrei(10).Text
                ListView1.SelectedItem.ListSubItems.Item(2) = "-"
                ListView1.SelectedItem.ListSubItems.Item(3) = "-"
                Y = ListView1.ListItems.Count
                Me.ListView1.Sorted = True
                Me.ListView1.SortKey = 0
                Me.ListView1.SortOrder = lvwAscending
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtProgTrei(9))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtProgTrei(9))
        Y = ListView1.ListItems.Count
        Me.ListView1.Sorted = True
        Me.ListView1.SortKey = 0
        Me.ListView1.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = txtProgTrei(10).Text
    ItemLst.SubItems(2) = "-"
    ItemLst.SubItems(3) = "-"
    Me.ListView1.SortOrder = lvwAscending
    txtProgTrei(9).SetFocus
End Sub

Private Sub IncluirInstrutorProg()
    If ValidaCampoInst = False Then Exit Sub
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView2.ListItems.Item(X).Selected = True
            If ListView2.ListItems.Item(X) = Me.Label17.Caption Then
                Label17.Caption = ListView2.ListItems.Item(X)
                ListView2.SelectedItem.ListSubItems.Item(1) = cboProgTrei(0).Text
                ListView2.SelectedItem.ListSubItems.Item(2) = txtProgTrei(6).Text
                ListView2.SelectedItem.ListSubItems.Item(3) = txtProgTrei(7).Text
                ListView2.SelectedItem.ListSubItems.Item(4) = cboProgTrei(1).Text
                Y = ListView2.ListItems.Count
                Me.ListView2.Sorted = True
                Me.ListView2.SortKey = 0
                Me.ListView2.SortOrder = lvwAscending
                Exit Sub
            End If
        Next
        Set ItemLst = ListView2.ListItems.Add(, , Label17)
        Y = ListView2.ListItems.Count
    Else
        Set ItemLst = ListView2.ListItems.Add(, , Label17)
        Y = ListView2.ListItems.Count
        Me.ListView2.Sorted = True
        Me.ListView2.SortKey = 0
        Me.ListView2.SortOrder = lvwDescending
    End If
    ItemLst.SubItems(1) = cboProgTrei(0).Text
    ItemLst.SubItems(2) = txtProgTrei(6).Text
    ItemLst.SubItems(3) = txtProgTrei(7).Text
    ItemLst.SubItems(4) = cboProgTrei(1).Text
    Me.ListView2.SortOrder = lvwAscending
    cboProgTrei(0).SetFocus
End Sub

Private Sub AlteraInstrutorProg()
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        If ListView2.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.Label17.Caption = ListView2.ListItems.Item(X)
    Me.cboProgTrei(0).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    Me.txtProgTrei(6).Text = ListView2.SelectedItem.ListSubItems.Item(2)
    Me.txtProgTrei(7).Text = ListView2.SelectedItem.ListSubItems.Item(3)
    Me.cboProgTrei(1).Text = ListView2.SelectedItem.ListSubItems.Item(1)
    MontaMascara
End Sub

Private Function ValidaCampoCol()
    Dim Y As Integer, X As Integer
    ValidaCampoCol = False
    For X = 9 To 10
        If txtProgTrei(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtProgTrei(X).Tag, Ok, critico, "Atenção"
            Me.txtProgTrei(X).SetFocus
            Exit Function
        End If
    Next
    ValidaCampoCol = True
End Function

Private Function ValidaCampoInst()
    Dim Y As Integer, X As Integer
    ValidaCampoInst = False
    For X = 6 To 7
        If txtProgTrei(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtProgTrei(X).Tag, Ok, critico, "Atenção"
            Me.txtProgTrei(X).SetFocus
            Exit Function
        End If
    Next
    ValidaCampoInst = True
End Function

Private Function ValidaCampo()
    Dim Y As Integer, X As Integer
    ValidaCampo = False
    
    If ListView2.ListItems.Count = 0 Then
        mobjMsg.Abrir "Não foi informado nenhum instrutor para esse treinamento", Ok, critico, "Atenção"
        Exit Function
    End If
    
    If mskProgTrei(0) = "__:__:__" Then
        mobjMsg.Abrir "Não foi informado o horário de início do treinamento", Ok, critico, "Atenção"
        Me.mskProgTrei(0).SetFocus
        Exit Function
    End If
    If mskProgTrei(1) = "__:__:__" Then
        mobjMsg.Abrir "Não foi informado o horário de término do treinamento", Ok, critico, "Atenção"
        Me.mskProgTrei(1).SetFocus
        Exit Function
    End If
    
    For X = 1 To 3
        If txtProgTrei(X).Text = "" Then
            mobjMsg.Abrir "Favor informar o campo " & Me.txtProgTrei(X).Tag, Ok, critico, "Atenção"
            Me.txtProgTrei(X).SetFocus
            Exit Function
        End If
    Next
    If Check1.Value = 1 Then
        If Label25 = "" Then
            mobjMsg.Abrir "Não foram definidos tópicos da Avaliação de Eficácia", Ok, critico, "Atenção"
            cmdCadastro(14).SetFocus
            Exit Function
        End If
    End If
    ValidaCampo = True
End Function

Private Sub GravarDados()
'On Error GoTo TrataErro
    
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarProgramacao As New ADODB.Recordset
    Dim SqlSalvarProgramacao As String
    Dim rsAchaMatriz As New ADODB.Recordset
    Dim SqlAchaMatriz As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim Y As Integer, X As Integer
    cnBanco.BeginTrans
   
    'SALVAR PROGRAMAÇÃO
    Dim statusProg As String
    
    If chamaForm.txtProgTrei(0) = "-" Then Status = "novo" Else Status = "editar"
    
    If chamaForm.txtProgTrei(0) <> "-" Then
        SqlSalvarProgramacao = "select * from tbProgramacao where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & Val(txtProgTrei(0)) & "'"
        rsSalvarProgramacao.Open SqlSalvarProgramacao, cnBanco, adOpenKeyset, adLockOptimistic
        If rsSalvarProgramacao.Fields(5) <> DTPicker2 Then 'data de inicio do treinamento
            statusProg = "Reagendado"
        Else
            statusProg = "Agendado"
        End If
    Else
        SqlSalvarProgramacao = "select * from tbProgramacao where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & GeraCodigo & "'"
        rsSalvarProgramacao.Open SqlSalvarProgramacao, cnBanco, adOpenKeyset, adLockOptimistic
        rsSalvarProgramacao.AddNew
        statusProg = "Agendado"
    End If
    
    rsSalvarProgramacao.Fields(0) = Val(txtProgTrei(0)) 'codigo da programação
    rsSalvarProgramacao.Fields(1) = DTPicker1 'Data da programação
    rsSalvarProgramacao.Fields(2) = txtProgTrei(1) 'Entidade responsavel pelo treinamento
    rsSalvarProgramacao.Fields(3) = txtProgTrei(2) 'Local do treinamento
    rsSalvarProgramacao.Fields(4) = txtProgTrei(3) 'código do colaborador responsavel pelo treinamento
    rsSalvarProgramacao.Fields(5) = DTPicker2 'data de inicio do treinamento
    rsSalvarProgramacao.Fields(6) = DTPicker3 'data de fim do treinamento
    rsSalvarProgramacao.Fields(7) = mskProgTrei(0) 'hora de inicio do treinamento
    rsSalvarProgramacao.Fields(8) = mskProgTrei(1) 'Hora de fim do treinamento
    rsSalvarProgramacao.Fields(24) = vCodcoligada 'Codigo da coligada
    If Check1.Value = 1 Then
        rsSalvarProgramacao.Fields(9) = 1 'DAE - Determinar avaliação de eficiencia
        'For X = 0 To 4
        '    If optProgTrei(X).Value = True Then rsSalvarProgramacao.Fields(10) = X 'metodo
        'Next
        'If X = 4 Then rsSalvarProgramacao.Fields(11) = txtProgTrei(5) 'metodo outro
        If chkProgTrei(0).Value = 0 Then rsSalvarProgramacao.Fields(20) = 0 Else rsSalvarProgramacao.Fields(20) = 1
        If chkProgTrei(1).Value = 0 Then rsSalvarProgramacao.Fields(21) = 0 Else rsSalvarProgramacao.Fields(21) = 1
        If chkProgTrei(2).Value = 0 Then rsSalvarProgramacao.Fields(22) = 0 Else rsSalvarProgramacao.Fields(22) = 1
        If chkProgTrei(3).Value = 0 Then rsSalvarProgramacao.Fields(23) = 0 Else rsSalvarProgramacao.Fields(23) = 1
        If chkProgTrei(4).Value = 0 Then
            rsSalvarProgramacao.Fields(10) = 0
        Else
            rsSalvarProgramacao.Fields(10) = 1
            rsSalvarProgramacao.Fields(11) = txtProgTrei(5) 'metodo outro
        End If
        rsSalvarProgramacao.Fields(16) = cboProgTrei(2).Text 'Tipo avaliador responsavel
        rsSalvarProgramacao.Fields(17) = txtProgTrei(11) 'Nome avaliador responsavel
        rsSalvarProgramacao.Fields(18) = DTPicker4 'Data da avaliação do avaliador
        rsSalvarProgramacao.Fields(19) = Label25 'Código do modelo de avaliação de eficácia
    Else
        rsSalvarProgramacao.Fields(9) = 0 'DAE - Determinar avaliação de eficiencia
        rsSalvarProgramacao.Fields(10) = Null 'metodo
        rsSalvarProgramacao.Fields(16) = "" 'Tipo avaliador responsavel
        rsSalvarProgramacao.Fields(17) = "" 'Nome avaliador responsavel
        rsSalvarProgramacao.Fields(18) = Null 'Data da avaliação do avaliador
    End If
    'CALCULA NOTA DE AVALIACAO DO TREINAMENTO >>>>>
    MediaTreinamento
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    If Label13 <> "" Then rsSalvarProgramacao.Fields(12) = RemoveMask(Label13) 'nota
    rsSalvarProgramacao.Fields(13) = txtProgTrei(8) 'observação
    
    If Status = "novo" Then
        rsSalvarProgramacao.Fields(14) = statusProg 'Status
    Else
        If IsNull(rsSalvarProgramacao.Fields(14)) And rsSalvarProgramacao.Fields(14) = "Agendado" Then rsSalvarProgramacao.Fields(14) = statusProg 'Status
    End If
    If Check2.Value = 1 Then rsSalvarProgramacao.Fields(15) = "S" Else rsSalvarProgramacao.Fields(9) = "N" 'ativo
    rsSalvarProgramacao.Update
    '-----------------------
    'SALVAR COLABORADORES EM TREINAMENTO - Listview1
    'SE O STATUS FOR DE EDIÇÃO PASSARÁ POR ESSE BLOCO PRIMEIRO >>>>>>>>>>>>>>>>>>>>>
    If Status <> "novo" Then
        SqlSalvar = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & Val(txtProgTrei(0)) & "'"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
        For X = 1 To rsSalvar.RecordCount
            rsSalvar.Fields(3) = Null
            rsSalvar.Fields(4) = "S"
            rsSalvar.Fields(6) = "Desmarcado"
            rsSalvar.MoveNext
        Next
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
    End If
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    'SE O STATUS FOR DE NOVO PASSARÁ IRÁ IGNORAR O BLOCO ACIMA E FARÁ DAKI P BAIXO>>
    Dim newCodigo As Integer
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        SqlAchaMatriz = "Select codmatriz from tbcolaboradoreshist where codcoligada = '" & vCodcoligada & "' and cpf = '" & ListView1.ListItems.Item(X) & "' and ativo = 'S'"
        rsAchaMatriz.Open SqlAchaMatriz, cnBanco, adOpenKeyset, adLockOptimistic
        
        SqlSalvar = "Select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(Text5) & "' and cpf = '" & ListView1.ListItems.Item(X) & "' and codprogramacao is null"
        rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
        
        If Not rsSalvar.EOF Then
            rsSalvar.Fields(3) = Val(txtProgTrei(0).Text)
            rsSalvar.Fields(6) = statusProg
            rsSalvar.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(2) 'Situação
            If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then rsSalvar.Fields(9) = RemoveMask(ListView1.SelectedItem.ListSubItems.Item(3)) 'Nota
        Else
            rsSalvar.AddNew
            'newCodigo = GeraCodigoPen
            rsSalvar.Fields(0) = ListView1.ListItems.Item(X) 'CPF do colaborador
            rsSalvar.Fields(1) = rsAchaMatriz.Fields(0) ' Código da matriz
            rsSalvar.Fields(2) = Val(Text5) 'Código do treinamento
            rsSalvar.Fields(3) = txtProgTrei(0) ' Código da programação
            rsSalvar.Fields(4) = "S" ' Ativo
            rsSalvar.Fields(5) = GeraCodigoPen 'Código de Identificação
            If IsNull(rsSalvar.Fields(6)) And rsSalvar.Fields(6) <> "Reagendado" Then rsSalvar.Fields(6) = statusProg Else rsSalvar.Fields(6) = "Agendado" 'Status
            'rsSalvar.Fields(6) = statusProg 'Status
            rsSalvar.Fields(7) = 1 'Tipo de programação
            rsSalvar.Fields(8) = ListView1.SelectedItem.ListSubItems.Item(2) 'Situação
            If ListView1.SelectedItem.ListSubItems.Item(3) <> "-" Then rsSalvar.Fields(9) = RemoveMask(ListView1.SelectedItem.ListSubItems.Item(3)) Else rsSalvar.Fields(9) = 0 'Nota
            rsSalvar.Fields(14) = vCodcoligada 'Codigo da coligada
        End If
        If Not rsSalvar.EOF Then rsSalvar.Update
        rsSalvar.Close
        Set rsSalvar = Nothing
        rsAchaMatriz.Close
        Set rsAchaMatriz = Nothing
    Next
    
    'SALVAR INSTRUTORES DA PROGRAMAÇÃO - Listview2
    sqlDeletar = "Delete from tbProgramacaoInstrutores where tbProgramacaoInstrutores.codcoligada = '" & vCodcoligada & "' and tbProgramacaoInstrutores.codprogramacao = '" & Val(txtProgTrei(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbProgramacaoInstrutores where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtProgTrei(0).Text)
        rsSalvar.Fields(1) = ListView2.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(2) = ListView2.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = ListView2.SelectedItem.ListSubItems.Item(3)
        rsSalvar.Fields(4) = ListView2.SelectedItem.ListSubItems.Item(4)
        rsSalvar.Fields(5) = ListView2.ListItems.Item(X)
        rsSalvar.Fields(6) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    cnBanco.CommitTrans
    rsSalvarProgramacao.Close
    Set rsSalvarProgramacao = Nothing
    rsSalvar.Close
    Set rsSalvar = Nothing
    If VerificaConcluido = True Then 'And RemoveMask(Label13) <> 0 Then
        mobjMsg.Abrir "Todos os participantes da programação foram avaliados. Deseja concluir a Programação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ConcluiTreinamento
            AchaColabor
        End If
    End If
    
    'GRAVA TOPICOS AVALIADOS NA AVALIAÇÃO DE EFICÁCIA DA PROGRAMACAO
    'SqlSalvar = "Select * from tbAvaliacaoProg where codprogramacao = '" & 0 & "'"
    'rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    'If rsSalvar.RecordCount > 0 Then
    '    For X = 1 To rsSalvar.RecordCount
    '        rsSalvar.Fields(0) = Val(txtProgTrei(0))
    '        rsSalvar.MoveNext
    '    Next
    'End If
    'If Not rsSalvar.EOF Then rsSalvar.Update
    'rsSalvar.Close
    'Set rsSalvar = Nothing

    mobjMsg.Abrir "Os dados da Programação foram salvos com sucesso", Ok, informacao, "SGC"
    'AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub MediaTreinamento()
    Dim rsMediaTreinamento As New ADODB.Recordset
    Dim SqlMediaTreinamento As String
    Dim mediaTrei As Double
    SqlMediaTreinamento = "select count(a.codprogramacao) as codprogramacao,a.cpf,count(a.codavaliacao) as codavaliacao,sum(a.pontuacao) as pontuacao,sum(b.peso) as peso,sum(a.pontuacao)/sum(b.peso)*100 as percentual from tbAvaliacaoTrei as a inner join tbAvaliacao as b on a.codcoligada = '" & vCodcoligada & "' and a.codavaliacao = b.codavaliacao and b.tipo = 'AT' and a.codprogramacao = '" & Val(txtProgTrei(0)) & "' group by a.cpf"
    rsMediaTreinamento.Open SqlMediaTreinamento, cnBanco, adOpenKeyset, adLockReadOnly
    mediaTrei = 0
    For X = 1 To rsMediaTreinamento.RecordCount
        mediaTrei = mediaTrei + rsMediaTreinamento.Fields(5)
        rsMediaTreinamento.MoveNext
    Next
    If mediaTrei > 0 Then Label13 = mediaTrei / rsMediaTreinamento.RecordCount Else Label13 = 0
    Label13 = Format(Label13, "#,##00.00;(#,##0.00)") & "%"
    rsMediaTreinamento.Close
    Set rsMediaTreinamento = Nothing
End Sub

Private Sub ConcluiTreinamento()
    Dim rsConcluir As New ADODB.Recordset
    Dim SqlConcluir As String
    
    Dim rsColabCandi As New ADODB.Recordset
    Dim SqlColabCandi As String
    Dim vNivel As Integer
    Dim X As Integer
    
    SqlConcluir = "Update tbProgramacao set status = 'Concluido' Where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & Val(txtProgTrei(0)) & "'"
    rsConcluir.Open SqlConcluir, cnBanco

    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        SqlConcluir = "select * from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & Val(txtProgTrei(0)) & "' and cpf = '" & ListView1.ListItems.Item(X) & "'"
        rsConcluir.Open SqlConcluir, cnBanco, adOpenKeyset, adLockOptimistic
        rsConcluir.Fields(6) = "Concluido"
        If Not IsNull(rsConcluir.Fields(12)) Then
            vNivel = rsConcluir.Fields(12)
        Else
            vNivel = 0
        End If
        rsConcluir.Update
        rsConcluir.Close
        Set rsConcluir = Nothing
    Next
    Dim ColabCandi As String
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        SqlColabCandi = "select tipo from tbColaboradores where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and cpf = '" & ListView1.ListItems.Item(X) & "'"
        rsColabCandi.Open SqlColabCandi, cnBanco, adOpenKeyset, adLockOptimistic
        ColabCandi = rsColabCandi.Fields(0)
        rsColabCandi.Update
        rsColabCandi.Close
        Set rsColabCandi = Nothing
        
        SqlConcluir = "select * from tbColaboradoresCur where codcoligada = '" & vCodcoligada & "'"
        rsConcluir.Open SqlConcluir, cnBanco, adOpenKeyset, adLockOptimistic
        rsConcluir.AddNew
        rsConcluir.Fields(0) = ListView1.ListItems.Item(X) 'CPF
        rsConcluir.Fields(1) = ColabCandi 'tipo
        rsConcluir.Fields(2) = Text5.Caption 'codtreinamento
        If ListView1.SelectedItem.ListSubItems.Item(2) = "Aprovado" Then
            rsConcluir.Fields(3) = "SA" ' SA - Sistema Aprovado
        Else
            rsConcluir.Fields(3) = "SR" ' SR - Sistema Reprovado
        End If
        rsConcluir.Fields(5) = vNivel
        rsConcluir.Fields(6) = vCodcoligada 'Codigo da coligada
        rsConcluir.Update
        rsConcluir.Close
        Set rsConcluir = Nothing
    Next
End Sub

Private Sub AchaColabor()
    Dim rsAchaColabor As New ADODB.Recordset
    Dim SqlAchaColabor As String
    
    Dim rsSelecionaTreiObr As New ADODB.Recordset
    Dim SqlSelecionaTreiObr As String
    
    Dim cCPF As String, cMatriz As Integer
    Dim X As Integer, vcodprog As Integer, vProcProg As Integer
    
    SqlSelecionaTreiObr = "select aplicavel from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(Text5) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    If rsSelecionaTreiObr.Fields(0) = "N" Then
        rsSelecionaTreiObr.Close
        Set rsSelecionaTreiObr = Nothing
        Exit Sub
    Else
        rsSelecionaTreiObr.Close
        Set rsSelecionaTreiObr = Nothing
    End If
    
    vProcProg = Val(txtProgTrei(0))
    GravaProgObr vcodprog
    
    For X = 1 To ListView1.ListItems.Count
        SqlAchaColabor = "select a.cpf,a.codmatriz,a.codprogramacao from tbpendentescur as a where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & ListView1.ListItems.Item(X) & "' and a.codprogramacao = '" & vProcProg & "'"
        rsAchaColabor.Open SqlAchaColabor, cnBanco, adOpenKeyset, adLockReadOnly
        cCPF = ListView1.ListItems.Item(X)
        cMatriz = rsAchaColabor.Fields(1)
        GravaTreiObrigatorio cCPF, cMatriz, vcodprog, vProcProg
        rsAchaColabor.Close
    Next
    Set rsAchaColabor = Nothing
End Sub

Private Sub GravaProgObr(vcodprog As Integer)
    Dim rsSalvarProgramacao As New ADODB.Recordset
    Dim SqlSalvarProgramacao As String
    
    SqlSalvarProgramacao = "select * from tbProgramacao where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & GeraCodigo & "'"
    rsSalvarProgramacao.Open SqlSalvarProgramacao, cnBanco, adOpenKeyset, adLockOptimistic
    rsSalvarProgramacao.AddNew
    statusProg = "Agendado"
    vcodprog = Val(txtProgTrei(0))
    
    rsSalvarProgramacao.Fields(0) = Val(txtProgTrei(0)) 'codigo da programação
    rsSalvarProgramacao.Fields(1) = DTPicker1 'Data da programação
    rsSalvarProgramacao.Fields(2) = txtProgTrei(1) 'Entidade responsavel pelo treinamento
    rsSalvarProgramacao.Fields(3) = txtProgTrei(2) 'Local do treinamento
    rsSalvarProgramacao.Fields(4) = txtProgTrei(3) 'código do colaborador responsavel pelo treinamento
    rsSalvarProgramacao.Fields(5) = DTPicker2 'data de inicio do treinamento
    rsSalvarProgramacao.Fields(6) = DTPicker3 'data de fim do treinamento
    rsSalvarProgramacao.Fields(7) = mskProgTrei(0) 'hora de inicio do treinamento
    rsSalvarProgramacao.Fields(8) = mskProgTrei(1) 'Hora de fim do treinamento
    rsSalvarProgramacao.Fields(24) = vCodcoligada 'Codigo da coligada
    If Check1.Value = 1 Then
        rsSalvarProgramacao.Fields(9) = 1 'DAE - Determinar avaliação de eficiencia
        'For X = 0 To 4
        '    If optProgTrei(X).Value = True Then rsSalvarProgramacao.Fields(10) = X 'metodo
        'Next
        'If X = 4 Then rsSalvarProgramacao.Fields(11) = txtProgTrei(5) 'metodo outro
        If chkProgTrei(0).Value = 0 Then rsSalvarProgramacao.Fields(20) = 0 Else rsSalvarProgramacao.Fields(20) = 1
        If chkProgTrei(1).Value = 0 Then rsSalvarProgramacao.Fields(21) = 0 Else rsSalvarProgramacao.Fields(21) = 1
        If chkProgTrei(2).Value = 0 Then rsSalvarProgramacao.Fields(22) = 0 Else rsSalvarProgramacao.Fields(22) = 1
        If chkProgTrei(3).Value = 0 Then rsSalvarProgramacao.Fields(23) = 0 Else rsSalvarProgramacao.Fields(23) = 1
        If chkProgTrei(4).Value = 0 Then
            rsSalvarProgramacao.Fields(10) = 0
        Else
            rsSalvarProgramacao.Fields(10) = 1
            rsSalvarProgramacao.Fields(11) = txtProgTrei(5) 'metodo outro
        End If
        rsSalvarProgramacao.Fields(16) = cboProgTrei(2).Text 'Tipo avaliador responsavel
        rsSalvarProgramacao.Fields(17) = txtProgTrei(11) 'Nome avaliador responsavel
        rsSalvarProgramacao.Fields(18) = DTPicker4 'Data da avaliação do avaliador
        rsSalvarProgramacao.Fields(19) = Label25 'Código do modelo de avaliação de eficácia
    Else
        rsSalvarProgramacao.Fields(9) = 0 'DAE - Determinar avaliação de eficiencia
        rsSalvarProgramacao.Fields(10) = Null 'metodo
        rsSalvarProgramacao.Fields(16) = "" 'Tipo avaliador responsavel
        rsSalvarProgramacao.Fields(17) = "" 'Nome avaliador responsavel
        rsSalvarProgramacao.Fields(18) = Null 'Data da avaliação do avaliador
    End If
    rsSalvarProgramacao.Fields(13) = txtProgTrei(8) 'observação
    rsSalvarProgramacao.Fields(14) = "Pré-agendado" 'Status
    If Check2.Value = 1 Then rsSalvarProgramacao.Fields(15) = "S" Else rsSalvarProgramacao.Fields(9) = "N" 'ativo
    rsSalvarProgramacao.Update
    'rsSalvarProgramacao.Close
    Set rsSalvarProgramacao = Nothing
End Sub

Private Sub GravaTreiObrigatorio(vCPF As String, vMatriz As Integer, vcodprog As Integer, vProcProg As Integer)
    'On Error Resume Next
    Dim rsAchaSetor As New ADODB.Recordset
    Dim SqlAchaSetor As String
    
    Dim rsSelecionaTreiObr As New ADODB.Recordset
    Dim SqlSelecionaTreiObr As String
    Dim rsGravaTreiObr As New ADODB.Recordset
    Dim SqlGravaTreiObr As String
    Dim contaID As Integer
    
    
    SqlGravaTreiObr = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel from tbPendentesCur where codcoligada = '" & vCodcoligada & "'"
    rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiObr.EOF Then
        rsGravaTreiObr.MoveLast
        contaID = rsGravaTreiObr.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiObr.Close
    Set rsGravaTreiObr = Nothing
    
    SqlGravaTreiObr = "Select a.cpf,a.codmatriz,a.codtreinamento,a.codprogramacao,a.ativo,a.id,a.status,a.tipoprogramacao,a.codnivel,a.codcoligada from tbPendentesCur as a  left join tbTreinamentosNiv as b on a.codnivel = b.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codprogramacao = '" & vProcProg & "'"
    rsGravaTreiObr.Open SqlGravaTreiObr, cnBanco, adOpenKeyset, adLockOptimistic
    
    rsGravaTreiObr.AddNew
    rsGravaTreiObr.Fields(0) = vCPF
    rsGravaTreiObr.Fields(1) = vMatriz
    rsGravaTreiObr.Fields(2) = Val(Text5)
    rsGravaTreiObr.Fields(3) = vcodprog
    rsGravaTreiObr.Fields(4) = "S"
    rsGravaTreiObr.Fields(5) = contaID
    rsGravaTreiObr.Fields(6) = "Pendente"
    rsGravaTreiObr.Fields(7) = 0
    rsGravaTreiObr.Fields(8) = 0
    rsGravaTreiObr.Fields(9) = vCodcoligada 'Codigo da coligada
    contaID = contaID + 1
    
    rsGravaTreiObr.Update
    rsGravaTreiObr.Close
    'rsSelecionaTreiObr.MoveNext
    
    Set rsGravaTreiObr = Nothing
    
    SqlSelecionaTreiObr = "select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(Text5) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    Dim vdiasPeriodico As Integer
    If rsSelecionaTreiObr.Fields(9) = "Meses" Then
        vdiasPeriodico = rsSelecionaTreiObr.Fields(8) * 30
    ElseIf rsSelecionaTreiObr.Fields(9) <> "Meses" Then
        vdiasPeriodico = rsSelecionaTreiObr.Fields(8) * 365
    End If
    rsSelecionaTreiObr.Close
    Set rsSelecionaTreiObr = Nothing
    
    Dim rsSalvarProgramacao As New ADODB.Recordset
    Dim SqlSalvarProgramacao As String
    SqlSalvarProgramacao = "Update tbProgramacao set dataprogramacao = '" & Format(DTPicker1.Value + vdiasPeriodico, "YYYY-MM-DD") & "' Where codcoligada = '" & vCodcoligada & "' and codprogramacao = '" & vcodprog & "'"
    rsSalvarProgramacao.Open SqlSalvarProgramacao, cnBanco
    
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Frame2.Enabled = True
        For X = 0 To chkProgTrei.Count - 1
            chkProgTrei(X).Enabled = True
        Next
        txtProgTrei(5).Enabled = True
        Label10.Enabled = True
        cboProgTrei(2).Enabled = True
        txtProgTrei(11).Enabled = True
        DTPicker4.Enabled = True
        Frame11.Enabled = True
        Label22.Enabled = True
        Label21.Enabled = True
    Else
        Frame2.Enabled = False
        For X = 0 To chkProgTrei.Count - 1
            chkProgTrei(X).Enabled = False
            chkProgTrei(X).Value = False
        Next
        txtProgTrei(5).Enabled = False
        Label10.Enabled = False
        cboProgTrei(2).Enabled = False
        txtProgTrei(11).Enabled = False
        DTPicker4.Enabled = False
        Frame11.Enabled = False
        Label22.Enabled = False
        Label21.Enabled = False
    End If
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(3).UseGreyscale = True
        cmdCadastro(3).DragMode = 1
        cmdCadastro(3).SpecialEffect = cbEngraved
    
        cmdCadastro(6).UseGreyscale = True
        cmdCadastro(6).DragMode = 1
        cmdCadastro(6).SpecialEffect = cbEngraved
    
        cmdCadastro(7).UseGreyscale = True
        cmdCadastro(7).DragMode = 1
        cmdCadastro(7).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(8).UseGreyscale = True
        cmdCadastro(8).DragMode = 1
        cmdCadastro(8).SpecialEffect = cbEngraved
    End If
    If vSal = "N" Then
        cmdCadastro(11).UseGreyscale = True
        cmdCadastro(11).DragMode = 1
        cmdCadastro(11).SpecialEffect = cbEngraved
    End If
    If vExc = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    
        cmdCadastro(9).UseGreyscale = True
        cmdCadastro(9).DragMode = 1
        cmdCadastro(9).SpecialEffect = cbEngraved
    End If
    If vAva = "N" Then
        cmdCadastro(10).UseGreyscale = True
        cmdCadastro(10).DragMode = 1
        cmdCadastro(10).SpecialEffect = cbEngraved
    
        cmdCadastro(14).UseGreyscale = True
        cmdCadastro(14).DragMode = 1
        cmdCadastro(14).SpecialEffect = cbEngraved
    End If
    If vImp = "N" Then
        cmdCadastro(13).UseGreyscale = True
        cmdCadastro(13).DragMode = 1
        cmdCadastro(13).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

