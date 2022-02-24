VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMatrizCapacitacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matriz de Capacitação"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmMatrizCapacitacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Status"
      Height          =   615
      Left            =   8760
      TabIndex        =   33
      Top             =   8520
      Width           =   1095
      Begin VB.CheckBox Check2 
         Caption         =   "Ativo"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Determinações de competência "
      Height          =   4575
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Width           =   9735
      Begin TabDlg.SSTab SSTab1 
         Height          =   4215
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7435
         _Version        =   393216
         Tabs            =   5
         Tab             =   2
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Experiências"
         TabPicture(0)   =   "frmMatrizCapacitacao.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtCadMatriz(8)"
         Tab(0).Control(1)=   "txtCadMatriz(9)"
         Tab(0).Control(2)=   "cboCadMatriz(3)"
         Tab(0).Control(3)=   "cboCadMatriz(2)"
         Tab(0).Control(4)=   "SkinLabel10"
         Tab(0).Control(5)=   "SkinLabel9"
         Tab(0).Control(6)=   "SkinLabel8"
         Tab(0).Control(7)=   "cmdCad(3)"
         Tab(0).Control(8)=   "ListView1"
         Tab(0).Control(9)=   "cmdCadastro(3)"
         Tab(0).Control(10)=   "cmdCadastro(2)"
         Tab(0).Control(11)=   "cmdCadastro(1)"
         Tab(0).Control(12)=   "cmdCadastro(0)"
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Habilidades"
         TabPicture(1)   =   "frmMatrizCapacitacao.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Check1"
         Tab(1).Control(1)=   "ListView2"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Cursos/treinamentos"
         TabPicture(2)   =   "frmMatrizCapacitacao.frx":0D02
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "cmdCadastro(8)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdCadastro(7)"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdCadastro(5)"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "ListView3"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "cmdCad(4)"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "SkinLabel12"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "SkinLabel13"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "SkinLabel11"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "txtCadMatriz(10)"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).Control(9)=   "txtCadMatriz(11)"
         Tab(2).Control(9).Enabled=   0   'False
         Tab(2).Control(10)=   "cboCadMatriz(5)"
         Tab(2).Control(10).Enabled=   0   'False
         Tab(2).ControlCount=   11
         TabCaption(3)   =   "Formação Escolar"
         TabPicture(3)   =   "frmMatrizCapacitacao.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtCadMatriz(14)"
         Tab(3).Control(1)=   "txtCadMatriz(13)"
         Tab(3).Control(2)=   "txtCadMatriz(12)"
         Tab(3).Control(3)=   "SkinLabel16"
         Tab(3).Control(4)=   "SkinLabel15"
         Tab(3).Control(5)=   "SkinLabel14"
         Tab(3).Control(6)=   "SkinLabel7"
         Tab(3).Control(7)=   "cmdCad(5)"
         Tab(3).Control(8)=   "cmdCadastro(6)"
         Tab(3).Control(9)=   "cmdCadastro(12)"
         Tab(3).Control(10)=   "cmdCadastro(13)"
         Tab(3).Control(11)=   "cmdCadastro(16)"
         Tab(3).Control(12)=   "ListView4"
         Tab(3).ControlCount=   13
         TabCaption(4)   =   "Revisões"
         TabPicture(4)   =   "frmMatrizCapacitacao.frx":0D3A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtCadMatriz(15)"
         Tab(4).Control(1)=   "DTPicker1"
         Tab(4).Control(2)=   "txtCadMatriz(16)"
         Tab(4).Control(3)=   "SkinLabel19"
         Tab(4).Control(4)=   "SkinLabel18"
         Tab(4).Control(5)=   "SkinLabel17"
         Tab(4).Control(6)=   "cmdCadastro(19)"
         Tab(4).Control(7)=   "cmdCadastro(20)"
         Tab(4).Control(8)=   "cmdCadastro(21)"
         Tab(4).Control(9)=   "cmdCadastro(22)"
         Tab(4).Control(10)=   "Frame7"
         Tab(4).Control(11)=   "ListView5"
         Tab(4).ControlCount=   12
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   15
            Left            =   -74880
            TabIndex        =   47
            Tag             =   "número de revisão do curso/treinamento"
            ToolTipText     =   "número de revisão do curso/treinamento"
            Top             =   720
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   -74040
            TabIndex        =   48
            Tag             =   "Data da revisão do curso/treinamento"
            ToolTipText     =   "Data da revisão do curso/treinamento"
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   16252929
            CurrentDate     =   40518
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   16
            Left            =   -72600
            TabIndex        =   49
            Tag             =   "Descritivo da revisão do curso/treinamento"
            ToolTipText     =   "Descritivo da revisão do curso/treinamento"
            Top             =   720
            Width           =   6975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
            Height          =   255
            Left            =   -72600
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0D56
            TabIndex        =   69
            Top             =   480
            Width           =   855
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
            Height          =   255
            Left            =   -74040
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0DC6
            TabIndex        =   94
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0E2E
            TabIndex        =   93
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   14
            Left            =   -66960
            TabIndex        =   38
            Tag             =   "Pontuação da formação escolar"
            ToolTipText     =   "Pontuação da formação escolar"
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   -73680
            TabIndex        =   37
            Top             =   720
            Width           =   5895
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   12
            Left            =   -74880
            TabIndex        =   36
            Tag             =   "Código da formação escolar"
            ToolTipText     =   "Código da formação escolar"
            Top             =   720
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
            Height          =   255
            Left            =   -66960
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0E9C
            TabIndex        =   92
            Top             =   480
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
            Height          =   255
            Left            =   -73680
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0F0E
            TabIndex        =   91
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0F7E
            TabIndex        =   90
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   5
            Left            =   6480
            TabIndex        =   39
            Top             =   690
            Width           =   2895
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   1320
            TabIndex        =   16
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":0FEA
            TabIndex        =   87
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
            Height          =   255
            Left            =   6480
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":1056
            TabIndex        =   89
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
            Height          =   255
            Left            =   1320
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":10C0
            TabIndex        =   88
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox txtCadMatriz 
            Height          =   285
            Index           =   8
            Left            =   -74880
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCadMatriz 
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   -73680
            TabIndex        =   9
            Top             =   720
            Width           =   5295
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   3
            ItemData        =   "frmMatrizCapacitacao.frx":1152
            Left            =   -66840
            List            =   "frmMatrizCapacitacao.frx":115C
            TabIndex        =   11
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "Meses"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   2
            ItemData        =   "frmMatrizCapacitacao.frx":116D
            Left            =   -67560
            List            =   "frmMatrizCapacitacao.frx":1195
            TabIndex        =   10
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "01"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   720
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   -67560
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":11C9
            TabIndex        =   86
            Top             =   480
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   -73680
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":1251
            TabIndex        =   85
            Top             =   480
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":12CB
            TabIndex        =   84
            Top             =   480
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   -72240
            OleObjectBlob   =   "frmMatrizCapacitacao.frx":1343
            TabIndex        =   83
            Top             =   1320
            Width           =   4335
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   5
            Left            =   -67680
            TabIndex        =   76
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   75
            Top             =   720
            Width           =   375
         End
         Begin VB.CommandButton cmdCad 
            Caption         =   "..."
            Height          =   255
            Index           =   3
            Left            =   -68280
            TabIndex        =   74
            Top             =   720
            Width           =   375
         End
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   19
            Left            =   -73080
            TabIndex        =   60
            Tag             =   "Excluir revisão"
            ToolTipText     =   "Excluir revisão"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":1413
            PICN            =   "frmMatrizCapacitacao.frx":142F
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
            TabIndex        =   59
            Tag             =   "Editar revisão"
            ToolTipText     =   "Editar revisão"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":2109
            PICN            =   "frmMatrizCapacitacao.frx":2125
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
            TabIndex        =   58
            Tag             =   "Novo revisão"
            ToolTipText     =   "Novo revisão"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":2DFF
            PICN            =   "frmMatrizCapacitacao.frx":2E1B
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
            TabIndex        =   57
            Tag             =   "Incluir revisão"
            ToolTipText     =   "Incluir revisão"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":3AF5
            PICN            =   "frmMatrizCapacitacao.frx":3B11
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
            TabIndex        =   40
            Tag             =   "Excluir Formação Escolar"
            ToolTipText     =   "Excluir Formação Escolar"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":47EB
            PICN            =   "frmMatrizCapacitacao.frx":4807
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
            Left            =   -73680
            TabIndex        =   41
            Tag             =   "Editar Formação Escolar"
            ToolTipText     =   "Editar Formação Escolar"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":54E1
            PICN            =   "frmMatrizCapacitacao.frx":54FD
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
            Left            =   -74280
            TabIndex        =   43
            Tag             =   "Nova Formação Escolar"
            ToolTipText     =   "Nova Formação Escolar"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":61D7
            PICN            =   "frmMatrizCapacitacao.frx":61F3
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
            TabIndex        =   45
            Tag             =   "Incluir Formação Escolar"
            ToolTipText     =   "Incluir Formação Escolar"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":6ECD
            PICN            =   "frmMatrizCapacitacao.frx":6EE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame7 
            Caption         =   "Parâmetros do Módulo Avaliador"
            Height          =   1695
            Left            =   -73200
            TabIndex        =   51
            Top             =   2160
            Visible         =   0   'False
            Width           =   7455
            Begin VB.TextBox txtCadMatriz 
               Height          =   285
               Index           =   4
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.CheckBox chkAvaliador 
               Caption         =   "Experiência:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   62
               Top             =   600
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox chkAvaliador 
               Caption         =   "Habilidades:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   61
               Top             =   840
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.CheckBox chkAvaliador 
               Caption         =   "Cursos/treinamentos:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   56
               Top             =   1080
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CheckBox chkAvaliador 
               Caption         =   "Formação escolar:"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   55
               Top             =   1320
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Frame Frame10 
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
               Left            =   2880
               TabIndex        =   53
               Top             =   600
               Visible         =   0   'False
               Width           =   1335
               Begin VB.Label Label41 
                  Caption         =   "Label41"
                  Height          =   255
                  Left            =   360
                  TabIndex        =   54
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   615
               End
            End
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   5160
               TabIndex        =   52
               Top             =   240
               Width           =   2175
            End
            Begin MSMask.MaskEdBox mskCadMatriz 
               Height          =   285
               Left            =   2520
               TabIndex        =   63
               Top             =   240
               Visible         =   0   'False
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   503
               _Version        =   393216
               PromptChar      =   "_"
            End
            Begin VB.Label Label37 
               Caption         =   "Label37"
               Height          =   255
               Left            =   2040
               TabIndex        =   68
               Top             =   600
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label Label38 
               Caption         =   "Label38"
               Height          =   255
               Left            =   2040
               TabIndex        =   67
               Top             =   840
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label39 
               Caption         =   "Label39"
               Height          =   255
               Left            =   2040
               TabIndex        =   66
               Top             =   1080
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label Label40 
               Caption         =   "Label40"
               Height          =   255
               Left            =   2040
               TabIndex        =   65
               Top             =   1320
               Visible         =   0   'False
               Width           =   615
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Marcar/Desmarcar"
            Height          =   255
            Left            =   -74880
            TabIndex        =   13
            Top             =   720
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   3015
            Left            =   -74880
            TabIndex        =   14
            Top             =   1080
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5318
            LabelEdit       =   1
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
         Begin MSComctlLib.ListView ListView1 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   12
            Top             =   1920
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3836
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
            Height          =   2175
            Left            =   120
            TabIndex        =   21
            Top             =   1920
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3836
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
            Height          =   2175
            Left            =   -74880
            TabIndex        =   35
            Top             =   1920
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3836
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
         Begin MSComctlLib.ListView ListView5 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   50
            Tag             =   "Grade de revisões"
            ToolTipText     =   "Grade de revisões"
            Top             =   1920
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   3625
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
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   3
            Left            =   -73080
            TabIndex        =   20
            Tag             =   "Excluir Experiência"
            ToolTipText     =   "Excluir Experiência"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":7BC3
            PICN            =   "frmMatrizCapacitacao.frx":7BDF
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
            TabIndex        =   19
            Tag             =   "Editar Experiência"
            ToolTipText     =   "Editar Experiência"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":88B9
            PICN            =   "frmMatrizCapacitacao.frx":88D5
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
            TabIndex        =   18
            Tag             =   "Novo Experiência"
            ToolTipText     =   "Novo Experiência"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":95AF
            PICN            =   "frmMatrizCapacitacao.frx":95CB
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
            TabIndex        =   17
            Tag             =   "Incluir Experiência"
            ToolTipText     =   "Incluir Experiência"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":A2A5
            PICN            =   "frmMatrizCapacitacao.frx":A2C1
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
            TabIndex        =   26
            Tag             =   "Excluir curso/treinamento"
            ToolTipText     =   "Excluir curso/treinamento"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":AF9B
            PICN            =   "frmMatrizCapacitacao.frx":AFB7
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
            TabIndex        =   25
            Tag             =   "Nova curso/treinamento"
            ToolTipText     =   "Nova curso/treinamento"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":BC91
            PICN            =   "frmMatrizCapacitacao.frx":BCAD
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
            TabIndex        =   24
            Tag             =   "Incluir curso/treinamento"
            ToolTipText     =   "Incluir curso/treinamento"
            Top             =   1200
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
            MICON           =   "frmMatrizCapacitacao.frx":C987
            PICN            =   "frmMatrizCapacitacao.frx":C9A3
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
   Begin VB.Frame Frame1 
      Caption         =   "Dados da matriz de capacitação"
      Height          =   3615
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txtCadMatriz 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   5
         Tag             =   "Nome do cargo"
         ToolTipText     =   "Nome do cargo"
         Top             =   1680
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D67D
         TabIndex        =   82
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Tag             =   "Código do cargo"
         ToolTipText     =   "Código do cargo"
         Top             =   1680
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D6F1
         TabIndex        =   81
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtCadMatriz 
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   1320
         TabIndex        =   3
         Tag             =   "Nome do setor"
         ToolTipText     =   "Nome do setor"
         Top             =   1080
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D769
         TabIndex        =   80
         Top             =   840
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D7DD
         TabIndex        =   79
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCadMatriz 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do departamento"
         ToolTipText     =   "Nome do departamento"
         Top             =   480
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D855
         TabIndex        =   78
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do departamento"
         ToolTipText     =   "Código do departamento"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmMatrizCapacitacao.frx":D8D7
         TabIndex        =   77
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   73
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   72
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdCad 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   71
         Top             =   480
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tempo mínimo de permanência no cargo "
         Height          =   855
         Left            =   6000
         TabIndex        =   42
         Top             =   1320
         Width           =   3615
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   4
            ItemData        =   "frmMatrizCapacitacao.frx":D94D
            Left            =   1080
            List            =   "frmMatrizCapacitacao.frx":D957
            TabIndex        =   46
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "Meses"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   1
            ItemData        =   "frmMatrizCapacitacao.frx":D968
            Left            =   240
            List            =   "frmMatrizCapacitacao.frx":D990
            TabIndex        =   44
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "001"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Código da Matriz "
         Height          =   855
         Left            =   6000
         TabIndex        =   32
         Top             =   360
         Width           =   2415
         Begin VB.TextBox txtCadMatriz 
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
            TabIndex        =   70
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Nível "
         Height          =   855
         Left            =   8520
         TabIndex        =   31
         Top             =   360
         Width           =   1095
         Begin VB.ComboBox cboCadMatriz 
            Height          =   315
            Index           =   0
            ItemData        =   "frmMatrizCapacitacao.frx":D9D0
            Left            =   120
            List            =   "frmMatrizCapacitacao.frx":DA04
            TabIndex        =   6
            Tag             =   "Nível do cargo"
            Text            =   "A"
            ToolTipText     =   "Nível do cargo"
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   1215
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2280
         Width           =   9495
      End
      Begin VB.TextBox txtCadMatriz 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Tag             =   "Código do setor"
         ToolTipText     =   "Código do setor"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Atividades:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   855
      End
   End
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   15
      Left            =   720
      TabIndex        =   28
      Top             =   8520
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
      MICON           =   "frmMatrizCapacitacao.frx":DA38
      PICN            =   "frmMatrizCapacitacao.frx":DA54
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
      TabIndex        =   27
      Top             =   8520
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
      MICON           =   "frmMatrizCapacitacao.frx":E72E
      PICN            =   "frmMatrizCapacitacao.frx":E74A
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
Attribute VB_Name = "frmMatrizCapacitacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsMatriz As New ADODB.Recordset
Private SqlMatriz As String
Private rsDepartamento As New ADODB.Recordset
Private sqlDepartamento As String
Private rsSetores As New ADODB.Recordset
Private SqlSetores As String
Private rsCargos As New ADODB.Recordset
Private SqlCargos As String
Private rsCursos As New ADODB.Recordset
Private SqlCursos As String
Private rsEscolaridade As New ADODB.Recordset
Private sqlEscolaridade As String
Private rsCandidatos As New ADODB.Recordset
Private sqlCandidatos As String

Private Status As String
Private rsLocal As New ADODB.Recordset

Private Sub Check1_Click()
    MarcaDesmarca
End Sub

Private Sub cmdCad_Click(Index As Integer)
    Select Case Index
    Case 0
        ChamaGridDepartamento 'rotina n desenvolvida
        CarregaDepartamento 'rotina n desenvolvida
    Case 1
        ChamaGridSetor 'rotina n desenvolvida
        CarregaSetor 'rotina n desenvolvida
    Case 2
        ChamaGridCargo 5 'rotina n desenvolvida
        CarregaCargo 5 'rotina n desenvolvida
    Case 3
        ChamaGridCargo 8
        CarregaCargo 8
    Case 4
        ChamaGridCurso
        CarregaCurso
        CompoeComboNivel cboCadMatriz(5), txtCadMatriz(10)
    Case 5
        ChamaGridEscolaridade
        CarregaEscolaridade
    End Select
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
        mobjMsg.Abrir "Deseja EXCLUIR essa experiência do cargo?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView1
            LimpaControlesExp
        End If
    Case 5
        mobjMsg.Abrir "Deseja EXCLUIR essa curso/treinamento do cargo?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView3
            LimpaControlesTreinamento
        End If
    Case 6
        mobjMsg.Abrir "Deseja EXCLUIR essa formação escolar do cargo?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView4
            LimpaControlesEscolaridade
        End If
    Case 7
        LimpaControlesTreinamento
    Case 8
        IncluirTreinamento
        LimpaControlesTreinamento
    Case 12
        AlteraEscolaridade
    Case 13
        LimpaControlesEscolaridade
    Case 14
        mobjMsg.Abrir "Deseja salvar os dados da matriz de capacitação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            '--- ROTINA EM TESTE ---
            atualizaCandidatos 1
            atualizaCandidatos 2
            '-----------------------
            gravaLog "Departamento: " & txtCadMatriz(1) & "-" & txtCadMatriz(2), "Setor: " & txtCadMatriz(3) & "-" & txtCadMatriz(41), "Cargo: " & txtCadMatriz(5) & "-" & txtCadMatriz(6) & "- Matriz: " & txtCadMatriz(0) & ", Nível: " & cboCadMatriz(0)
            Pesquisa = "0"
        End If
    Case 15
        mobjMsg.Abrir "Deseja sair da tela de cadastro da matriz de capacitação?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = "0"
            Unload Me
        End If
    Case 16
        IncluirEscolaridade
        LimpaControlesEscolaridade
    Case 19
        mobjMsg.Abrir "Deseja EXCLUIR essa revisão do treinamento?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView5
            LimpaControlesRevisao
        End If
    Case 20
        AlteraRevisao
    Case 21
        LimpaControlesRevisao
    Case 22
        IncluirRevisao
        LimpaControlesRevisao
    End Select
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    Legenda = cmdCadastro(Index).ToolTipText
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'ESSA SUB EH PARA QDO TECLA ENTER, ELE FUNCIONAR COMO TAB
    'PARA ISSO, A PROPRIEDADE KEYPREVIEW DO FORM DEVE ESTAR TRUE
    If KeyAscii = 13 Then SendKeys "{TAB}": KeyAscii = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Form_Load()
    Status = Pesquisa
    listview_cabecalho
    SSTab1.Tab = 0
    'CompoeCombo cboCadMatriz(1), "tbescolaridade", "codescolaridade", "nomeescolaridade"
    If Status = "novo" Then
        LimpaControles
        Compoe_Listview
        MarcaDesmarca
    ElseIf Status = "editar" Then
        ResultPesq
        CompoeMatrizExp
        Compoe_Listview
        CompoeMatrizHab
        CompoeMatrizTrei
        CompoeMatrizRev
        CompoeFor
        'DesbloqueiaControles
    End If
    configControles
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
    ListView1.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView1.ColumnHeaders.Add , , "Nome do cargo", ListView1.Width / 2
    ListView1.ColumnHeaders.Add , , "Tempo de experiência", ListView1.Width / 5
    
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView2.ColumnHeaders.Add , , "Habilidade", ListView1.Width / 1.5
    ListView2.ColumnHeaders.Add , , "Peso", ListView1.Width / 10
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView3.ColumnHeaders.Add , , "Nome do curso/treinamento", ListView1.Width / 1.5
    ListView3.ColumnHeaders.Add , , "Nível", ListView3.Width / 4.5
    
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "Código", ListView1.Width / 12
    ListView4.ColumnHeaders.Add , , "Formação escolar", ListView1.Width / 1.5
    ListView4.ColumnHeaders.Add , , "Pontuação", ListView1.Width / 8.5
    
    ListView5.ColumnHeaders.Add , , "Revisão", ListView5.Width / 11
    ListView5.ColumnHeaders.Add , , "Data", ListView5.Width / 8
    ListView5.ColumnHeaders.Add , , "Detalhes", ListView5.Width / 1.5
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview
End Sub

Private Sub GravarDados()
On Error GoTo TrataErro
    If ValidaCampo = False Then Exit Sub
    Dim rsSalvarMatriz As New ADODB.Recordset
    Dim SqlSalvarMatriz As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarMatriz = "select * from tbMatriz where codcoligada = '" & vCodcoligada & "' and codMatriz = '" & txtCadMatriz(0) & "'"
    rsSalvarMatriz.Open SqlSalvarMatriz, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarMatriz.EOF Then rsSalvarMatriz.AddNew
    rsSalvarMatriz.Fields(0) = Val(txtCadMatriz(0)) 'codmatriz
    rsSalvarMatriz.Fields(1) = txtCadMatriz(1) 'coddepartamento
    rsSalvarMatriz.Fields(2) = txtCadMatriz(3) 'codsetor
    rsSalvarMatriz.Fields(3) = txtCadMatriz(5) 'codcargo
    rsSalvarMatriz.Fields(4) = cboCadMatriz(0) 'nivel
    rsSalvarMatriz.Fields(5) = txtCadMatriz(7) 'atividades
    If Check2.Value = 1 Then rsSalvarMatriz.Fields(6) = "S" Else rsSalvarMatriz.Fields(6) = "N" 'ativo
    rsSalvarMatriz.Fields(7) = Format(cboCadMatriz(1), "000") & " " & cboCadMatriz(4) 'Tempo minimo de permanencia no cargo
    rsSalvarMatriz.Fields(8) = vCodcoligada ' Codigo da coligada
    rsSalvarMatriz.Update
    
    '>>>>>> GRAVAR EXPERIENCIA <<<<<<<<<
    sqlDeletar = "Delete from tbMatrizExp where tbMatrizExp.codcoligada = '" & vCodcoligada & "' and tbMatrizExp.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbMatrizExp where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadMatriz(0).Text)
        rsSalvar.Fields(1) = Val(ListView1.ListItems.Item(X))
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>>>> GRAVAR HABILIDADE <<<<<<<<<
    sqlDeletar = "Delete from tbMatrizHab where tbMatrizHab.codcoligada = '" & vCodcoligada & "' and tbMatrizHab.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbMatrizHab where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(X).Selected = True
        If ListView2.ListItems.Item(X).Checked = True Then
            rsSalvar.AddNew
            rsSalvar.Fields(0) = Val(txtCadMatriz(0).Text)
            rsSalvar.Fields(1) = Val(ListView2.ListItems.Item(X))
            rsSalvar.Fields(2) = vCodcoligada ' Codigo da coligada
        End If
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>>>> GRAVAR CURSO/TREINAMENTO <<<<<<<<<
    sqlDeletar = "Delete from tbMatrizCur where tbMatrizCur.codcoligada = '" & vCodcoligada & "' and tbMatrizCur.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbMatrizCur where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView3.ListItems.Count
        ListView3.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadMatriz(0).Text)
        rsSalvar.Fields(1) = Val(ListView3.ListItems.Item(X))
        If ListView3.SelectedItem.ListSubItems.Item(2) <> "-" Then rsSalvar.Fields(2) = Val(Mid$(ListView3.SelectedItem.ListSubItems.Item(2), 1, 2)) Else rsSalvar.Fields(2) = 0
        rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>>>> GRAVAR ESCOLARIDADE <<<<<<<<<
    sqlDeletar = "Delete from tbMatrizEsc where tbMatrizEsc.codcoligada = '" & vCodcoligada & "' and tbMatrizEsc.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbMatrizEsc where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView4.ListItems.Count
        ListView4.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadMatriz(0).Text)
        rsSalvar.Fields(1) = Val(ListView4.ListItems.Item(X))
        rsSalvar.Fields(2) = ListView4.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(3) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>> GRAVA REVISAO DA MATRIZ
    sqlDeletar = "Delete from tbMatrizRev where tbMatrizRev.codcoligada = '" & vCodcoligada & "' and tbMatrizRev.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbMatrizRev where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView5.ListItems.Count
        ListView5.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadMatriz(0).Text)
        rsSalvar.Fields(1) = ListView5.ListItems.Item(X)
        rsSalvar.Fields(2) = ListView5.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(3) = ListView5.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(4) = vCodcoligada ' Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    
    
    cnBanco.CommitTrans
    rsSalvarMatriz.Close
    Set rsSalvarMatriz = Nothing
    rsSalvar.Close
    Set rsSalvar = Nothing
    mobjMsg.Abrir "Os dados da Matriz de capacitação foram salvos com sucesso", Ok, informacao, "SGC"
    AtualizaListview
    Unload Me
    Set frmMatrizCapacitacao = Nothing
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    'cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    'DesbloqueiaControles
    For X = 0 To 16
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(41) = ""
    txtCadMatriz(0) = Format(GeraCodigo, "000000")
End Sub

Private Sub LimpaControlesExp()
    Dim X As Integer
    txtCadMatriz(8).Enabled = True
    cmdCad(4).Enabled = True
    
    cboCadMatriz(2).Text = ""
    cboCadMatriz(3).Text = ""
    For X = 8 To 9
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(8).SetFocus
End Sub

Private Sub LimpaControlesTreinamento()
    Dim X As Integer
    txtCadMatriz(10).Enabled = True
    cmdCad(4).Enabled = True
    
    For X = 10 To 11
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(10).SetFocus
End Sub

Private Sub LimpaControlesEscolaridade()
    Dim X As Integer
    txtCadMatriz(12).Enabled = True
    cmdCad(5).Enabled = True
    
    For X = 12 To 14
        txtCadMatriz(X) = ""
    Next
    txtCadMatriz(12).SetFocus
End Sub

Private Sub CompoeControles()
    Dim X As Integer
    DTPicker1 = Date
    txtCadMatriz(0).Text = Format(rsMatriz.Fields(0), "000000") 'Código da Matriz
    txtCadMatriz(1).Text = Format(rsMatriz.Fields(1), "000000") 'Código do Departamento
    txtCadMatriz(2).Text = rsMatriz.Fields(7) 'Nome do Departamento
    txtCadMatriz(3).Text = Format(rsMatriz.Fields(2), "000000") 'Nome do Setor
    txtCadMatriz(41).Text = rsMatriz.Fields(8) 'Nome do Setor
    txtCadMatriz(5).Text = Format(rsMatriz.Fields(3), "000000") 'Nome do Cargo
    txtCadMatriz(6).Text = rsMatriz.Fields(9) 'Nome do Cargo
    txtCadMatriz(7).Text = rsMatriz.Fields(5) 'Atividades
    cboCadMatriz(0).Text = rsMatriz.Fields(4) 'Nível
    If rsMatriz.Fields(6) = "S" Then
        Check2.Value = 1
    Else
        Check2.Value = 0
    End If
    If Not IsNull(rsMatriz.Fields(10)) Then
        cboCadMatriz(1).Text = Format(Mid$(rsMatriz.Fields(10), 1, 3), "000")
        cboCadMatriz(4).Text = Mid$(rsMatriz.Fields(10), 4, 10)
    End If
End Sub

Private Sub Compoe_Listview()
    Dim rsHabilidade As New ADODB.Recordset
    Dim sqlHabilidades As String
    sqlHabilidades = "Select * from tbHabilidades where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by codhabilidade"
    rsHabilidade.Open sqlHabilidades, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    
    X = 0
    While Not rsHabilidade.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsHabilidade.Fields(0), "00"))
        ItemLst.SubItems(1) = "" & rsHabilidade.Fields(1)
        ItemLst.SubItems(2) = "" & rsHabilidade.Fields(2)
        rsHabilidade.MoveNext
        X = X + 1
    Wend
    rsHabilidade.Close
    Set rsHabilidade = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub CompoeMatrizExp()
    Dim rsExp As New ADODB.Recordset
    Dim sqlExp As String
    sqlExp = "Select tbMatrizExp.*, tbcargos.nomecargo from tbMatrizExp,tbcargos where tbMatrizExp.codcoligada = '" & vCodcoligada & "' and tbMatrizExp.codcargo=tbcargos.codcargo and tbMatrizExp.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsExp.Open sqlExp, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsExp.EOF
    
        Set ItemLst = ListView1.ListItems.Add(, , Format(rsExp.Fields(1), "000000"))
        ItemLst.SubItems(1) = "" & rsExp.Fields(4)
        ItemLst.SubItems(2) = "" & rsExp.Fields(2)
        rsExp.MoveNext
        X = X + 1
    Wend
    rsExp.Close
    Set rsExp = Nothing
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwAscending
End Sub

Private Sub CompoeMatrizHab()
    Dim rsHab As New ADODB.Recordset
    Dim sqlHab As String
    sqlHab = "Select * from tbMatrizHab where tbMatrizHab.codcoligada = '" & vCodcoligada & "' and tbMatrizHab.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsHab.Open sqlHab, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    Y = ListView2.ListItems.Count
    While Not rsHab.EOF
        For X = 1 To Y
            ListView2.ListItems(X).Selected = True
            If Val(ListView2.ListItems.Item(X)) = rsHab.Fields(1) Then
                ListView2.ListItems.Item(X).Checked = True
            End If
        Next
        rsHab.MoveNext
    Wend
    rsHab.Close
    Set rsHab = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub CompoeMatrizTrei()
    Dim rsTrei As New ADODB.Recordset
    Dim sqlTrei As String
    'sqlTrei = "Select tbMatrizCur.*, tbTreinamentos.nometreinamento from tbMatrizCur,tbTreinamentos where tbMatrizCur.codtreinamento=tbTreinamentos.codtreinamento and tbMatrizCur.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    sqlTrei = "Select a.*, b.nometreinamento, c.codnivel, c.nomenivel from tbMatrizCur as a left join tbTreinamentos as b on a.codtreinamento=b.codtreinamento left join tbTreinamentosNiv as c on b.codtreinamento = c.codtreinamento and a.codnivel = c.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsTrei.Open sqlTrei, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsTrei.EOF
        Set ItemLst = ListView3.ListItems.Add(, , Format(rsTrei.Fields(1), "000000"))
        ItemLst.SubItems(1) = "" & rsTrei.Fields(4)
        If Not IsNull(rsTrei.Fields(5)) Then ItemLst.SubItems(2) = Format(rsTrei.Fields(5), "00") & " - " & rsTrei.Fields(6) Else ItemLst.SubItems(2) = "-"
        rsTrei.MoveNext
        X = X + 1
    Wend
    rsTrei.Close
    Set rsTrei = Nothing
    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending
End Sub

Private Sub CompoeMatrizRev()
    Dim rsRevisao As New ADODB.Recordset
    Dim SqlRevisao As String
    SqlRevisao = "Select * from tbMatrizRev where codcoligada = '" & vCodcoligada & "' and codmatriz = '" & Val(txtCadMatriz(0)) & "'Order by codmatriz,revisao"
    rsRevisao.Open SqlRevisao, cnBanco, adOpenKeyset, adLockOptimistic
    
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsRevisao.EOF
        Set ItemLst = ListView5.ListItems.Add(, , rsRevisao.Fields(1))
        ItemLst.SubItems(1) = "" & rsRevisao.Fields(2)
        ItemLst.SubItems(2) = "" & rsRevisao.Fields(3)
        rsRevisao.MoveNext
        X = X + 1
    Wend
    rsRevisao.Close
    Set rsRevisao = Nothing
End Sub

Private Sub CompoeFor()
    Dim rsEsc As New ADODB.Recordset
    Dim sqlEsc As String
    sqlEsc = "Select tbMatrizEsc.*, tbEscolaridade.nomeescolaridade from tbMatrizEsc,tbEscolaridade where tbMatrizEsc.codcoligada = '" & vCodcoligada & "' and tbMatrizEsc.codescolaridade=tbEscolaridade.codescolaridade and tbMatrizEsc.codmatriz = '" & Val(txtCadMatriz(0).Text) & "'"
    rsEsc.Open sqlEsc, cnBanco, adOpenKeyset, adLockOptimistic
    Dim ItemLst As ListItem
    Dim X As Integer
    
    X = 0
    While Not rsEsc.EOF
        Set ItemLst = ListView4.ListItems.Add(, , Format(rsEsc.Fields(1), "000000"))
        ItemLst.SubItems(1) = "" & rsEsc.Fields(4)
        ItemLst.SubItems(2) = "" & rsEsc.Fields(2)
        rsEsc.MoveNext
        X = X + 1
    Wend
    rsEsc.Close
    Set rsEsc = Nothing
    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 0
    Me.ListView4.SortOrder = lvwAscending
End Sub

' (INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<
Private Sub IncluirExperiencia()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadMatriz(8) Then
                Me.txtCadMatriz(8) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(9)
                ListView1.SelectedItem.ListSubItems.Item(2) = cboCadMatriz(2) & " " & cboCadMatriz(3)
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
    ItemLst.SubItems(2) = cboCadMatriz(2) & " " & cboCadMatriz(3)
    txtCadMatriz(8).SetFocus
End Sub

Private Sub AlteraExperiencia()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(8).Text = ListView1.ListItems.Item(X)
    Me.txtCadMatriz(9).Text = ListView1.SelectedItem.ListSubItems.Item(1)
    Me.cboCadMatriz(2).Text = Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 1, 2)
    Me.cboCadMatriz(3).Text = Mid$(ListView1.SelectedItem.ListSubItems.Item(2), 4, 10)
    txtCadMatriz(8).Enabled = False
    txtCadMatriz(9).Enabled = False
    cmdCad(4).Enabled = False
End Sub
' (FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE EXPERIÊNCIA <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE CURSOS/TREINAMENTOS <<<<<<<<<<
Private Sub IncluirTreinamento()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView3.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView3.ListItems.Item(X) = Me.txtCadMatriz(10) Then
                Me.txtCadMatriz(10) = ListView3.ListItems.Item(X)
                ListView3.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(11)
                ListView3.SelectedItem.ListSubItems.Item(2) = cboCadMatriz(5)
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
    ItemLst.SubItems(2) = cboCadMatriz(5)
    txtCadMatriz(10).SetFocus
End Sub
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE CURSOS/TREINAMENTOS <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE FORMAÇÃO ESCOLAR <<<<<<<<<<
Private Sub IncluirEscolaridade()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    If ValidaCampoEscolar = False Then Exit Sub
    Y = ListView4.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView4.ListItems.Item(X) = Me.txtCadMatriz(12) Then
                Me.txtCadMatriz(12) = ListView4.ListItems.Item(X)
                ListView4.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(13)
                ListView4.SelectedItem.ListSubItems.Item(2) = txtCadMatriz(14)
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
    ItemLst.SubItems(2) = txtCadMatriz(14)
    txtCadMatriz(10).SetFocus
End Sub

Private Sub AlteraEscolaridade()
    Dim Y As Integer, X As Integer
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        If ListView4.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(12).Text = ListView4.ListItems.Item(X)
    Me.txtCadMatriz(13).Text = ListView4.SelectedItem.ListSubItems.Item(1)
    Me.txtCadMatriz(14).Text = ListView4.SelectedItem.ListSubItems.Item(2)
    txtCadMatriz(12).Enabled = False
    txtCadMatriz(13).Enabled = False
    cmdCad(5).Enabled = False
End Sub
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE FORMAÇÃO ESCOLAR <<<<<<<<<<

'(INICIO) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE REVISOES <<<<<<<<<<
Private Sub IncluirRevisao()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView5.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView5.ListItems.Item(X) = Me.txtCadMatriz(15) Then
                ListView5.ListItems.Item(X).Selected = True
                Me.txtCadMatriz(6) = ListView5.ListItems.Item(X)
                ListView5.SelectedItem.ListSubItems.Item(1) = DTPicker1
                ListView5.SelectedItem.ListSubItems.Item(2) = txtCadMatriz(16)
                Y = ListView5.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView5.ListItems.Add(, , txtCadMatriz(15))
        Y = ListView5.ListItems.Count
    Else
        Set ItemLst = ListView5.ListItems.Add(, , txtCadMatriz(15))
        Y = ListView5.ListItems.Count
    End If
    ItemLst.SubItems(1) = DTPicker1
    ItemLst.SubItems(2) = txtCadMatriz(16)
    txtCadMatriz(15).Text = ""
    DTPicker1 = Date
    txtCadMatriz(16).Text = ""
    txtCadMatriz(15).SetFocus
End Sub

Private Sub LimpaControlesRevisao()
    Dim X As Integer
    For X = 15 To 16
        txtCadMatriz(X) = ""
    Next
    DTPicker1 = Date
End Sub

Private Sub AlteraRevisao()
    Dim Y As Integer, X As Integer
    Y = ListView5.ListItems.Count
    For X = 1 To Y
        If ListView5.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadMatriz(15).Text = ListView5.ListItems.Item(X)
    Me.txtCadMatriz(16).Text = ListView5.SelectedItem.ListSubItems.Item(2)
    DTPicker1 = ListView5.SelectedItem.ListSubItems.Item(1)
End Sub
'(FIM) >>>>>>>> CONTROLES DOS BOTOES DA GUIA DE REVISOES <<<<<<<<<<

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadMatriz(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(0).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(0).SetFocus
        Exit Function
    End If
    If txtCadMatriz(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(1).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(1).SetFocus
        Exit Function
    End If
    If txtCadMatriz(2).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(2).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(2).SetFocus
        Exit Function
    End If
    If txtCadMatriz(41).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(41).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(41).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Function ValidaCampoEscolar()
    ValidaCampoEscolar = False
    If txtCadMatriz(12).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(12).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(12).SetFocus
        Exit Function
    End If
    If txtCadMatriz(14).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadMatriz(14).Tag, Ok, critico, "Atenção"
        Me.txtCadMatriz(14).SetFocus
        Exit Function
    End If
    ValidaCampoEscolar = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadMatriz.Count - 1
        txtCadMatriz(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadMatriz.Count - 1
        txtCadMatriz(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirMatriz
    SqlGera = "Select top 1 * from tbMatriz where codcoligada = '" & vCodcoligada & "' order by codMatriz Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsSetores.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadMatriz(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharMatriz
End Function

Private Sub AbrirMatriz()
    SqlSetores = "Select * from tbMatriz where codcoligada = '" & vCodcoligada & "' Order by codMatriz"
    rsSetores.Open SqlSetores, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharMatriz()
    rsSetores.Close
    Set rsSetores = Nothing
End Sub

Private Sub ResultPesq()
    SqlMatriz = "Select tbMatriz.codmatriz,tbMatriz.coddepartamento,tbMatriz.codsetor,tbMatriz.codcargo,tbMatriz.nivel,tbMatriz.atividades,tbMatriz.ativo,tbdepartamentos.nomedepartamento,tbsetores.nomesetor,tbcargos.nomecargo,tbmatriz.tempoMin from tbMatriz,tbdepartamentos,tbsetores,tbcargos Where tbmatriz.codcoligada = '" & vCodcoligada & "' and codmatriz = '" & Val(varGlobal) & "' and tbdepartamentos.coddepartamento = tbmatriz.coddepartamento and tbsetores.codsetor = tbmatriz.codsetor and tbMatriz.codcargo = tbCargos.codcargo order by codMatriz"
    rsMatriz.Open SqlMatriz, cnBanco, adOpenKeyset, adLockReadOnly
    If rsMatriz.RecordCount > 0 Then
        CompoeControles
    Else
        mobjMsg.Abrir "Matriz não encontrada", Ok, critico, "Atenção"
    End If
    rsMatriz.Close
    Set rsMatriz = Nothing
End Sub

Private Sub AtualizaListview()
    On Error GoTo Err
    Dim ItemLst As ListItem 'variavel q recebe as propriedades do Listview,
    Y = MeuLV.ListView1.ListItems.Count
    For X = 1 To Y
        If MeuLV.ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadMatriz(0), "000000")) 'ID Matriz
        ItemLst.SubItems(1) = txtCadMatriz(1).Text ' Código do cargo
        ItemLst.SubItems(2) = txtCadMatriz(6).Text 'Nome do cargo
        ItemLst.SubItems(3) = cboCadMatriz(0).Text 'Nível do cargo
        ItemLst.SubItems(4) = txtCadMatriz(7).Text 'Atividade dessenvolvida pelo cargo
        If Check2.Value = 0 Then
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            ItemLst.SubItems(5) = ""
            ItemLst.ListSubItems.Item(5).ReportIcon = "OK"
        End If
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadMatriz(1).Text ' Código do cargo
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = txtCadMatriz(6).Text 'Nome do cargo
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = cboCadMatriz(0).Text 'Nível do cargo
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = txtCadMatriz(7).Text 'Atividade dessenvolvida pelo cargo
        If Check2.Value = 0 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "EXC"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(5).ReportIcon = "OK"
        End If
    End If
    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    AlteraExperiencia
End Sub

Private Sub ListView5_DblClick()
    AlteraRevisao
End Sub

Private Sub txtCadMatriz_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    Case 1
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaDepartamento
        End If
    Case 3
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaSetor
        End If
    Case 5
        If KeyCode = 13 Or KeyCode = 9 Then ' Enter ou TAB
            Pesquisa = 1
            CarregaCargo 5
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

Private Sub CarregaDepartamento()
    Dim X As Integer
    sqlDepartamento = "Select * from tbdepartamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by coddepartamento"
    rsDepartamento.Open sqlDepartamento, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsDepartamento.EOF Then rsDepartamento.MoveFirst
    rsDepartamento.Find "coddepartamento=" & "'" & Val(Me.txtCadMatriz(1)) & "'"
    If rsDepartamento.EOF Then
        txtCadMatriz(1).Text = Format(txtCadMatriz(1), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Departamento não cadastrado", Ok, critico, "Atenção"
            txtCadMatriz(2) = ""
        End If
    Else
        txtCadMatriz(1).Text = Format(rsDepartamento.Fields(0), "000000") & ""
        txtCadMatriz(2).Text = rsDepartamento.Fields(1)
    End If
    rsDepartamento.Close
    Set rsDepartamento = Nothing
End Sub

Private Sub ChamaGridDepartamento()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbdepartamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by nomedepartamento"
    procnom = "nomedepartamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de departamento"
    Pesquisa = frmMatrizCapacitacao.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomedepartamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadMatriz(1).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub CarregaSetor()
    Dim X As Integer
    SqlSetores = "Select * from tbSetores where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and tbSetores.coddepartamento = '" & Val(txtCadMatriz(1)) & "' order by tbSetores.codsetor"
    rsSetores.Open SqlSetores, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsSetores.EOF Then rsSetores.MoveFirst
    rsSetores.Find "codSetor=" & "'" & Val(Me.txtCadMatriz(3)) & "'"
    If rsSetores.EOF Then
        txtCadMatriz(3).Text = Format(txtCadMatriz(3), "000000") & ""
        If Val(Pesquisa) <> 0 Then
            mobjMsg.Abrir "Setor não cadastrado para o departamento", Ok, critico, "Atenção"
            txtCadMatriz(41) = ""
        End If
    Else
        txtCadMatriz(3).Text = Format(rsSetores.Fields(0), "000000") & ""
        txtCadMatriz(41).Text = rsSetores.Fields(1)
    End If
    rsSetores.Close
    Set rsSetores = Nothing
End Sub

Private Sub ChamaGridSetor()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbSetores where codcoligada = '" & vCodcoligada & "' and ativo = 'S' and tbSetores.coddepartamento = '" & Val(txtCadMatriz(1)) & "' order by tbSetores.nomesetor"
    procnom = "nomeSetor"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Setor"
    Pesquisa = frmMatrizCapacitacao.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nomeSetor=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            txtCadMatriz(3).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
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
    Sqlp = "Select * from tbCargos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by nomecargo"
    procnom = "nomeCargo"
    campo = 2
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Cargo"
    Pesquisa = frmMatrizCapacitacao.Tag
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
    Sqlp = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo = 'S' order by tbTreinamentos.nometreinamento"
    procnom = "nometreinamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa de Treinamento"
    Pesquisa = frmMatrizCapacitacao.Tag
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
    Pesquisa = frmMatrizCapacitacao.Tag
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

Private Sub MarcaDesmarca()
    'Adiciona processo ao item selecionado no Listview
    Dim Y As Integer, X As Integer
    Y = ListView2.ListItems.Count
    For X = 1 To Y
        ListView2.ListItems(X).Selected = True
        If ListView2.ListItems.Item(X).Checked = True Then
            ListView2.ListItems.Item(X).Checked = False
        Else
            ListView2.ListItems.Item(X).Checked = True
        End If
    Next
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    
        cmdCadastro(8).UseGreyscale = True
        cmdCadastro(8).DragMode = 1
        cmdCadastro(8).SpecialEffect = cbEngraved
    
        cmdCadastro(7).UseGreyscale = True
        cmdCadastro(7).DragMode = 1
        cmdCadastro(7).SpecialEffect = cbEngraved
    
        cmdCadastro(16).UseGreyscale = True
        cmdCadastro(16).DragMode = 1
        cmdCadastro(16).SpecialEffect = cbEngraved
    
        cmdCadastro(13).UseGreyscale = True
        cmdCadastro(13).DragMode = 1
        cmdCadastro(13).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
    
        cmdCadastro(12).UseGreyscale = True
        cmdCadastro(12).DragMode = 1
        cmdCadastro(12).SpecialEffect = cbEngraved
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
    End If
End Sub

'ESTUDAR SE AS ROTINAS ABAIXO VAO ATENDER
Private Sub atualizaCandidatos(filtra As Integer)
    'FILTRA
    '1 = Colaborador
    '2 = Candidato
'    Dim X As Integer
    txtCadMatriz(4) = txtCadMatriz(0) ' Matriz
    Text1 = txtCadMatriz(0) & txtCadMatriz(6) ' Matrix+nome do cargo
    If filtra = 1 Then
        'If Check1.Value = 0 Then Exit Sub
        sqlCandidatos = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.id,a.compav from tbColaboradores as a inner join tbColaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join  tbcargos as d on c.codcargo = d.codcargo where a.tipo = 'colaborador' and b.ativo = 'S' and b.codmatriz = '" & Val(txtCadMatriz(4)) & "' order by a.id"
    Else
        'If Check2.Value = 0 Then Exit Sub
        sqlCandidatos = "Select a.cpf,a.nomecolaborador,b.codmatriz,d.nomecargo,a.id,a.compav from tbColaboradores as a inner join tbColaboradoresHist as b on a.codcoligada = '" & vCodcoligada & "' and a.cpf = b.cpf inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join  tbcargos as d on c.codcargo = d.codcargo where a.tipo = 'candidato' and b.ativo = 'S' order by a.id"
    End If
    rsCandidatos.Open sqlCandidatos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsCandidatos.RecordCount = 0 Then
        rsCandidatos.Close
        Set rsCandidatos = Nothing
        Exit Sub
    End If
    
    If Not rsCandidatos.EOF Then
        While Not rsCandidatos.EOF '.Move(Val(Combo1.Text))
            chkAvaliador(0).Value = 0
            chkAvaliador(1).Value = 0
            chkAvaliador(2).Value = 0
            chkAvaliador(3).Value = 0
            For X = 0 To Len(rsCandidatos.Fields(5))
                If Mid$(rsCandidatos.Fields(5), X + 1, 1) = "E" Then chkAvaliador(0).Value = 1
                If Mid$(rsCandidatos.Fields(5), X + 1, 1) = "H" Then chkAvaliador(1).Value = 1
                If Mid$(rsCandidatos.Fields(5), X + 1, 1) = "T" Then chkAvaliador(2).Value = 1
                If Mid$(rsCandidatos.Fields(5), X + 1, 1) = "F" Then chkAvaliador(3).Value = 1
            Next
            mskCadMatriz = rsCandidatos.Fields(0) ' CPF
            If filtra = 1 Then
                Avaliador "colaborador"
                GravaColaboradores
                excluiProgramacao mskCadMatriz, txtCadMatriz(4)
                GravaTreiPen mskCadMatriz, txtCadMatriz(4)
                If GeraIntr = "S" Then GravaTreiIntrodutorio mskCadMatriz, txtCadMatriz(4)
                If GeraObri = "S" Then GravaTreiObrigatorio mskCadMatriz, txtCadMatriz(4)
            Else
                Avaliador "candidato"
                GravaColaboradores
                excluiProgramacao mskCadMatriz, txtCadMatriz(4)
                'GravaTreiPen mskCadMatriz, txtCadMatriz(4)
                'If GeraIntr = "S" Then GravaTreiIntrodutorio mskCadMatriz, txtCadMatriz(4)
                'If GeraObri = "S" Then GravaTreiObrigatorio mskCadMatriz, txtCadMatriz(4)
            End If
            rsCandidatos.MoveNext
        Wend
    End If
    rsCandidatos.Close
    Set rsCandidatos = Nothing
End Sub

Private Sub GravaColaboradores()
    Dim rsGravaColaboradores As New ADODB.Recordset
    Dim sqlGravaColaboradores As String
    Dim vIdent As Integer
    vIdent = rsCandidatos.Fields(4)
    sqlGravaColaboradores = "Update tbColaboradores set mediageral = '" & Replace(RemoveMask(Label41), ",", ".") & "' Where codcoligada = '" & vCodcoligada & "' and id = '" & vIdent & "'"
    rsGravaColaboradores.Open sqlGravaColaboradores, cnBanco
End Sub

'AS 4 ROTINAS ABAIXO SAO RESPONSAVEIS POR GRAVAR TODOS OS TREINAMENTOS
'DO NOVO COLABORADOR NA TABELA TBPENDENTESCUR
'TAIS ROTINAS DEVERAO SER GLOBALIZADAS
'Deixar GLOBAL as seguintes rotinas listadas abaixo:
'excluirProgramacao
'GravarTreiPen
'GravaTreiIntrodutorio
'GravaTreiObrigatorio
Private Sub excluiProgramacao(vCPF As String, vMatriz As Integer)
    ' Rotina deleta toda a programação "Agendada ou Pendente" se o
    ' colaborador sofrer alteração de cargo
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    sqlDeletar = "Delete from tbPendentesCur where codcoligada = '" & vCodcoligada & "' and tbPendentesCur.cpf = '" & vCPF & "' and status = 'Pendente' and codmatriz = '" & vMatriz & "'"
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

Private Sub GravaTreiPen(vCPF As String, vMatriz As Integer)
    On Error Resume Next
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
                rsPendentesCur.Fields(14) = vCodcoligada ' Codigo da coligada
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
                    rsPendentesCur.Fields(14) = vCodcoligada ' Codigo da coligada
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
    
    If ListView5.ListItems.Count > 1 Then
        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    Else
        SqlSelecionaTreiInt = "select * from tbTreinamentosint where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codcoligada = '" & vCodcoligada & "' and codsetor = '" & rsAchaSetor.Fields(0) & "'"
    End If
    
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
            rsGravaTreiInt.Fields(9) = vCodcoligada ' Codigo da coligada
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
                rsGravaTreiInt.Fields(9) = vCodcoligada ' Codigo da coligada
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
    
    SqlSelecionaTreiObr = "select * from tbTreinamentosObr where codcoligada = '" & vCodcoligada & "' and codsetor = 0 or codsetor = '" & rsAchaSetor.Fields(0) & "'"
    rsSelecionaTreiObr.Open SqlSelecionaTreiObr, cnBanco, adOpenKeyset, adLockReadOnly
    
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
    
    While Not rsSelecionaTreiObr.EOF
        SqlGravaTreiObr = "Select a.cpf,a.codmatriz,a.codtreinamento,a.codprogramacao,a.ativo,a.id,a.status,a.tipoprogramacao,a.codnivel,a.codcoligada from tbPendentesCur as a left join tbTreinamentosNiv as b on a.codnivel = b.codnivel where a.codcoligada = '" & vCodcoligada & "' and a.cpf = '" & vCPF & "' and a.codtreinamento ='" & rsSelecionaTreiObr.Fields(0) & "'"
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

