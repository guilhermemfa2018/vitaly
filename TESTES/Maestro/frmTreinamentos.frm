VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreinamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cursos/Treinamentos"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   Icon            =   "frmTreinamentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin MAESTRO.chameleonButton cmdCadastro 
      Height          =   615
      Index           =   15
      Left            =   720
      TabIndex        =   39
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
      MICON           =   "frmTreinamentos.frx":0CCA
      PICN            =   "frmTreinamentos.frx":0CE6
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
      TabIndex        =   38
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
      MICON           =   "frmTreinamentos.frx":19C0
      PICN            =   "frmTreinamentos.frx":19DC
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
      Left            =   9120
      TabIndex        =   46
      Top             =   8760
      Width           =   1095
      Begin VB.CheckBox Check4 
         Caption         =   "Ativo"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Tag             =   "Status do curso/treinamento"
         ToolTipText     =   "Status do curso/treinamento"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dados "
      Height          =   8535
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtCadTreinamento 
         Height          =   975
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "Objetivo do curso/treinamento"
         ToolTipText     =   "Objetivo do curso/treinamento"
         Top             =   2400
         Width           =   8055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTreinamentos.frx":26B6
         TabIndex        =   70
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtCadTreinamento 
         Height          =   975
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Conteúdo do curso/treinamento"
         ToolTipText     =   "Conteúdo do curso/treinamento"
         Top             =   1080
         Width           =   8055
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTreinamentos.frx":2726
         TabIndex        =   69
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   8280
         TabIndex        =   65
         ToolTipText     =   "Cadastrar"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtCadTreinamento 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "Código do curso/treinamento"
         ToolTipText     =   "Código do curso/treinamento"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCadTreinamento 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Nome do curso/treinamento"
         ToolTipText     =   "Nome do curso/treinamento"
         Top             =   480
         Width           =   4815
      End
      Begin VB.ComboBox cboCadTreinamento 
         Height          =   315
         Index           =   0
         ItemData        =   "frmTreinamentos.frx":2796
         Left            =   6240
         List            =   "frmTreinamentos.frx":27AF
         TabIndex        =   2
         Tag             =   "Tipo do curso/treinamento"
         ToolTipText     =   "Tipo do curso/treinamento"
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cboCadTreinamento 
         Height          =   315
         Index           =   1
         ItemData        =   "frmTreinamentos.frx":280A
         Left            =   8880
         List            =   "frmTreinamentos.frx":2814
         TabIndex        =   3
         Tag             =   "Origem do curso/treinamento"
         Text            =   "Interno"
         ToolTipText     =   "Origem do curso/treinamento"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   8880
         OleObjectBlob   =   "frmTreinamentos.frx":282A
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   6240
         OleObjectBlob   =   "frmTreinamentos.frx":2896
         TabIndex        =   56
         Top             =   240
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   1320
         OleObjectBlob   =   "frmTreinamentos.frx":28FE
         TabIndex        =   55
         Top             =   240
         Width           =   1935
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "frmTreinamentos.frx":2984
         TabIndex        =   54
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Caption         =   "Valor (R$)"
         Height          =   1095
         Left            =   8280
         TabIndex        =   48
         Top             =   960
         Width           =   1695
         Begin VB.TextBox txtCadTreinamento 
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "habilitar níveis"
         Height          =   735
         Left            =   6840
         TabIndex        =   47
         Top             =   3480
         Width           =   3135
         Begin VB.CheckBox Check5 
            Caption         =   "Níveis"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Carga Horária (hs)"
         Height          =   1095
         Left            =   8280
         TabIndex        =   45
         Top             =   2280
         Width           =   1695
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Tag             =   "Carga horária do curso/treinamento"
            ToolTipText     =   "Carga horária do curso/treinamento"
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###:##"
            PromptChar      =   "_"
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   120
         TabIndex        =   44
         Top             =   4320
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7223
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Introdutórios"
         TabPicture(0)   =   "frmTreinamentos.frx":29F0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ListView2"
         Tab(0).Control(1)=   "cmdCadastro(4)"
         Tab(0).Control(2)=   "cmdCadastro(5)"
         Tab(0).Control(3)=   "SkinLabel5"
         Tab(0).Control(4)=   "cboCadTreinamento(5)"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Obrigatórios"
         TabPicture(1)   =   "frmTreinamentos.frx":2A0C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ListView3"
         Tab(1).Control(1)=   "cmdCadastro(7)"
         Tab(1).Control(2)=   "cmdCadastro(6)"
         Tab(1).Control(3)=   "SkinLabel6"
         Tab(1).Control(4)=   "cboCadTreinamento(4)"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Níveis"
         TabPicture(2)   =   "frmTreinamentos.frx":2A28
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListView4"
         Tab(2).Control(1)=   "cmdCadastro(11)"
         Tab(2).Control(2)=   "cmdCadastro(10)"
         Tab(2).Control(3)=   "cmdCadastro(9)"
         Tab(2).Control(4)=   "cmdCadastro(8)"
         Tab(2).Control(5)=   "SkinLabel7"
         Tab(2).Control(6)=   "txtCadTreinamento(5)"
         Tab(2).Control(7)=   "SkinLabel8"
         Tab(2).Control(8)=   "txtCadTreinamento(4)"
         Tab(2).ControlCount=   9
         TabCaption(3)   =   "Histórico de revisões"
         TabPicture(3)   =   "frmTreinamentos.frx":2A44
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtCadTreinamento(6)"
         Tab(3).Control(1)=   "DTPicker1"
         Tab(3).Control(2)=   "txtCadTreinamento(7)"
         Tab(3).Control(3)=   "SkinLabel11"
         Tab(3).Control(4)=   "SkinLabel10"
         Tab(3).Control(5)=   "SkinLabel9"
         Tab(3).Control(6)=   "cmdCadastro(3)"
         Tab(3).Control(7)=   "cmdCadastro(2)"
         Tab(3).Control(8)=   "cmdCadastro(1)"
         Tab(3).Control(9)=   "cmdCadastro(0)"
         Tab(3).Control(10)=   "ListView1"
         Tab(3).Control(11)=   "lblStatusRev"
         Tab(3).ControlCount=   12
         TabCaption(4)   =   "Observações"
         TabPicture(4)   =   "frmTreinamentos.frx":2A60
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "txtCadTreinamento(8)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "Frame8"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "cmdCadastro(13)"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "cmdCadastro(12)"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "Frame9"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).ControlCount=   5
         Begin VB.Frame Frame9 
            Caption         =   "      Definir FASE"
            Height          =   735
            Left            =   8040
            TabIndex        =   66
            Top             =   3120
            Width           =   1575
            Begin VB.CheckBox Check6 
               Height          =   255
               Left            =   120
               TabIndex        =   68
               Top             =   0
               Width           =   255
            End
            Begin VB.ComboBox cboCadTreinamento 
               Enabled         =   0   'False
               Height          =   315
               Index           =   6
               ItemData        =   "frmTreinamentos.frx":2A7C
               Left            =   480
               List            =   "frmTreinamentos.frx":2A9E
               TabIndex        =   67
               Text            =   "01"
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.TextBox txtCadTreinamento 
            Height          =   285
            Index           =   6
            Left            =   -74880
            TabIndex        =   29
            Tag             =   "número de revisão do curso/treinamento"
            ToolTipText     =   "número de revisão do curso/treinamento"
            Top             =   720
            Width           =   735
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   -74040
            TabIndex        =   30
            Tag             =   "Data da revisão do curso/treinamento"
            ToolTipText     =   "Data da revisão do curso/treinamento"
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            Format          =   55902209
            CurrentDate     =   40518
         End
         Begin VB.TextBox txtCadTreinamento 
            Height          =   885
            Index           =   7
            Left            =   -72360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Tag             =   "Descritivo da revisão do curso/treinamento"
            ToolTipText     =   "Descritivo da revisão do curso/treinamento"
            Top             =   720
            Width           =   7095
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
            Height          =   255
            Left            =   -72360
            OleObjectBlob   =   "frmTreinamentos.frx":2ACA
            TabIndex        =   64
            Top             =   480
            Width           =   1335
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   255
            Left            =   -74040
            OleObjectBlob   =   "frmTreinamentos.frx":2B3A
            TabIndex        =   63
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmTreinamentos.frx":2BA2
            TabIndex        =   62
            Top             =   480
            Width           =   735
         End
         Begin VB.ComboBox cboCadTreinamento 
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            ItemData        =   "frmTreinamentos.frx":2C10
            Left            =   -74880
            List            =   "frmTreinamentos.frx":2C12
            TabIndex        =   18
            Top             =   720
            Width           =   2775
         End
         Begin VB.ComboBox cboCadTreinamento 
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   -74880
            TabIndex        =   14
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtCadTreinamento 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   -74040
            TabIndex        =   23
            Tag             =   "Descritivo da revisão do curso/treinamento"
            ToolTipText     =   "Descritivo da revisão do curso/treinamento"
            Top             =   720
            Width           =   8775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
            Height          =   255
            Left            =   -74040
            OleObjectBlob   =   "frmTreinamentos.frx":2C14
            TabIndex        =   61
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtCadTreinamento 
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   -74880
            TabIndex        =   22
            Tag             =   "número de revisão do curso/treinamento"
            ToolTipText     =   "número de revisão do curso/treinamento"
            Top             =   720
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmTreinamentos.frx":2C86
            TabIndex        =   60
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmTreinamentos.frx":2CF0
            TabIndex        =   59
            Top             =   480
            Width           =   1575
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   -74880
            OleObjectBlob   =   "frmTreinamentos.frx":2D76
            TabIndex        =   58
            Top             =   480
            Width           =   1575
         End
         Begin MAESTRO.chameleonButton cmdCadastro 
            Height          =   615
            Index           =   8
            Left            =   -73080
            TabIndex        =   27
            Tag             =   "Excluir Nível"
            ToolTipText     =   "Excluir Nível"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":2DFC
            PICN            =   "frmTreinamentos.frx":2E18
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
            TabIndex        =   26
            Tag             =   "Editar Nível"
            ToolTipText     =   "Editar Nível"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":3AF2
            PICN            =   "frmTreinamentos.frx":3B0E
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
            Left            =   5760
            TabIndex        =   53
            Tag             =   "Excluir treinamento"
            ToolTipText     =   "Excluir treinamento"
            Top             =   3240
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
            MICON           =   "frmTreinamentos.frx":47E8
            PICN            =   "frmTreinamentos.frx":4804
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
            Left            =   5160
            TabIndex        =   52
            Tag             =   "Incluir treinamento"
            ToolTipText     =   "Incluir treinamento"
            Top             =   3240
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
            MICON           =   "frmTreinamentos.frx":54DE
            PICN            =   "frmTreinamentos.frx":54FA
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
            TabIndex        =   35
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
            MICON           =   "frmTreinamentos.frx":61D4
            PICN            =   "frmTreinamentos.frx":61F0
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
            TabIndex        =   34
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
            MICON           =   "frmTreinamentos.frx":6ECA
            PICN            =   "frmTreinamentos.frx":6EE6
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
            TabIndex        =   33
            Tag             =   "Nova revisão"
            ToolTipText     =   "Nova revisão"
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
            MICON           =   "frmTreinamentos.frx":7BC0
            PICN            =   "frmTreinamentos.frx":7BDC
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
            TabIndex        =   32
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
            MICON           =   "frmTreinamentos.frx":88B6
            PICN            =   "frmTreinamentos.frx":88D2
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
            TabIndex        =   25
            Tag             =   "Novo Nível"
            ToolTipText     =   "Novo Nível"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":95AC
            PICN            =   "frmTreinamentos.frx":95C8
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
            Left            =   -74880
            TabIndex        =   24
            Tag             =   "Incluir Nível"
            ToolTipText     =   "Incluir Nível"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":A2A2
            PICN            =   "frmTreinamentos.frx":A2BE
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
            Left            =   -74280
            TabIndex        =   20
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":AF98
            PICN            =   "frmTreinamentos.frx":AFB4
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
            Left            =   -74880
            TabIndex        =   19
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":BC8E
            PICN            =   "frmTreinamentos.frx":BCAA
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
            Left            =   -74280
            TabIndex        =   16
            Tag             =   "Excluir Setor"
            ToolTipText     =   "Excluir Setor"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":C984
            PICN            =   "frmTreinamentos.frx":C9A0
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
            Index           =   4
            Left            =   -74880
            TabIndex        =   15
            Tag             =   "Incluir Setor"
            ToolTipText     =   "Incluir Setor"
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   1085
            BTYPE           =   2
            TX              =   ""
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
            MICON           =   "frmTreinamentos.frx":D67A
            PICN            =   "frmTreinamentos.frx":D696
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame Frame8 
            Caption         =   "Agrupar treinamentos "
            Height          =   2775
            Left            =   5160
            TabIndex        =   50
            Top             =   360
            Width           =   4455
            Begin MSComctlLib.ListView ListView5 
               Height          =   2415
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   4260
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
         End
         Begin VB.TextBox txtCadTreinamento 
            Height          =   3375
            Index           =   8
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            Tag             =   "Observação do curso/treinamento"
            ToolTipText     =   "Observação do curso/treinamento"
            Top             =   480
            Width           =   4815
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   17
            Top             =   1920
            Width           =   9615
            _ExtentX        =   16960
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
            Enabled         =   0   'False
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   21
            Top             =   1920
            Width           =   9615
            _ExtentX        =   16960
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
            Enabled         =   0   'False
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   36
            Tag             =   "Grade de revisões"
            ToolTipText     =   "Grade de revisões"
            Top             =   1920
            Width           =   9615
            _ExtentX        =   16960
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
         Begin MSComctlLib.ListView ListView4 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   28
            Tag             =   "Grade de revisões"
            ToolTipText     =   "Grade de revisões"
            Top             =   1920
            Width           =   9615
            _ExtentX        =   16960
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
            Enabled         =   0   'False
            NumItems        =   0
         End
         Begin VB.Label lblStatusRev 
            BackColor       =   &H8000000C&
            Height          =   255
            Left            =   -67080
            TabIndex        =   49
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Periodicidade "
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   3480
         Width           =   3495
         Begin VB.ComboBox cboCadTreinamento 
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            ItemData        =   "frmTreinamentos.frx":E370
            Left            =   2520
            List            =   "frmTreinamentos.frx":E37A
            TabIndex        =   10
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "Meses"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboCadTreinamento 
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "frmTreinamentos.frx":E38B
            Left            =   1800
            List            =   "frmTreinamentos.frx":E3B3
            TabIndex        =   9
            Tag             =   "Periodicidade do curso/treinamento"
            Text            =   "01"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Aplicável a cada:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Tag             =   "Periodicidade do curso/treinamento"
            ToolTipText     =   "Periodicidade do curso/treinamento"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Requisitos "
         Height          =   735
         Left            =   3720
         TabIndex        =   42
         Top             =   3480
         Width           =   3015
         Begin VB.CheckBox Check2 
            Caption         =   "Obrigatório"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Introdutório"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Tag             =   "Requisitos do curso/treinamento"
            ToolTipText     =   "Requisitos do do curso/treinamento"
            Top             =   360
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmTreinamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsTreinamentos As New ADODB.Recordset
Private SqlTreinamentos As String
Private rsRevisao As New ADODB.Recordset
Private SqlRevisao As String
Private rsNivel As New ADODB.Recordset
Private SqlNivel As String
Private rsSet As New ADODB.Recordset
Private SqlSet As String
Private rsSetObr As New ADODB.Recordset
Private SqlSetObr As String
Private rsAgrup As New ADODB.Recordset
Private SqlAgrup As String

Private rsListaIntObr As New ADODB.Recordset
Private SqlListaIntObr As String
Private rsListaAgrup As New ADODB.Recordset
Private SqlListaAgrup As String

Private Status As String
Private rsColaborador As New ADODB.Recordset
Private SqlColaborador As String
Private rsLocal As New ADODB.Recordset

Private Sub cboCadTreinamento_LostFocus(Index As Integer)
    Select Case Index
    'Case 4
    '    MontaMascara
    End Select
End Sub

Private Sub chameleonButton1_Click()
    ChamaGridTrei
End Sub

Private Sub ChamaGridTrei()
    Dim F As New frmpesqger
    Dim Iposicao As Variant
    Sqlp = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and ativo <> 'N' order by nometreinamento"
    procnom = "nometreinamento"
    campo = 1
    Campo1 = 0
    Load F
    F.Caption = "Pesquisa treinamentos"
    Pesquisa = frmTreinamentos.Tag
    F.Show 1
    If Pesquisa <> "" Then
        rsLocal.Open Sqlp, cnBanco, adOpenKeyset, adLockOptimistic
        If rsLocal.RecordCount < 1 Then Exit Sub
        Iposicao = rsLocal.Bookmark
        rsLocal.MoveFirst
        rsLocal.Find "nometreinamento=" & "'" & Pesquisa & "'"
        If Not rsLocal.EOF Then
            IncluirTrei
            'txtCadSetor(2).Text = Format(rsLocal.Fields(0), "000000")
        End If
        rsLocal.Close
        Set rsLocal = Nothing
    End If
End Sub

Private Sub IncluirTrei()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView5.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView5.ListItems.Item(X) = Format(rsLocal.Fields(0), "000000") Then
                ListView5.ListItems.Item(X).Selected = True
                'Me.txtCadTreinamento(5) = ListView5.ListItems.Item(X)
                ListView5.SelectedItem.ListSubItems.Item(1) = rsLocal.Fields(1)
                If Check6.Value = 1 Then
                    ListView5.SelectedItem.ListSubItems.Item(2) = cboCadTreinamento(6)
                Else
                    ListView5.SelectedItem.ListSubItems.Item(2) = "-"
                End If
                Y = ListView5.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView5.ListItems.Add(, , Format(rsLocal.Fields(0), "000000"))
        Y = ListView5.ListItems.Count
    Else
        Set ItemLst = ListView5.ListItems.Add(, , Format(rsLocal.Fields(0), "000000"))
        Y = ListView5.ListItems.Count
    End If
    ItemLst.SubItems(1) = rsLocal.Fields(1)
    If Check6.Value = 1 Then
        ItemLst.SubItems(2) = cboCadTreinamento(6)
    Else
        ItemLst.SubItems(2) = "-"
    End If
    cmdCadastro(13).SetFocus
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        cboCadTreinamento(5).Enabled = True
        SSTab1.TabEnabled(0) = True
        cmdCadastro(4).Enabled = True
        cmdCadastro(5).Enabled = True
        ListView2.Enabled = True
    Else
        cboCadTreinamento(5).Enabled = False
        SSTab1.TabEnabled(0) = False
        cmdCadastro(4).Enabled = False
        cmdCadastro(5).Enabled = False
        ListView2.Enabled = False
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        cboCadTreinamento(4).Enabled = True
        SSTab1.TabEnabled(1) = True
        cmdCadastro(6).Enabled = True
        cmdCadastro(7).Enabled = True
        ListView3.Enabled = True
    Else
        cboCadTreinamento(4).Enabled = False
        SSTab1.TabEnabled(1) = False
        cmdCadastro(6).Enabled = False
        cmdCadastro(7).Enabled = False
        ListView3.Enabled = False
        Check3.Value = 0
    End If
End Sub

Private Sub Check5_Click()
    If Check5.Value = 1 Then
        SSTab1.TabEnabled(2) = True
        cmdCadastro(8).Enabled = True
        cmdCadastro(9).Enabled = True
        cmdCadastro(10).Enabled = True
        cmdCadastro(11).Enabled = True
        txtCadTreinamento(4).Enabled = True
        txtCadTreinamento(5).Enabled = True
        ListView4.Enabled = True
    Else
        SSTab1.TabEnabled(2) = False
        cmdCadastro(8).Enabled = False
        cmdCadastro(9).Enabled = False
        cmdCadastro(10).Enabled = False
        cmdCadastro(11).Enabled = False
        txtCadTreinamento(4).Enabled = False
        txtCadTreinamento(5).Enabled = False
        ListView4.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        cboCadTreinamento(2).Enabled = True
        cboCadTreinamento(3).Enabled = True
        Check2.Value = 1
    Else
        cboCadTreinamento(2).Enabled = False
        cboCadTreinamento(3).Enabled = False
    End If
End Sub

Private Sub Check6_Click()
    If Check6.Value = 1 Then
        cboCadTreinamento(6).Enabled = True
    Else
        cboCadTreinamento(6).Enabled = False
    End If

End Sub

Private Sub cmdCadastro_Click(Index As Integer)
    Select Case Index
    Case 0
        IncluirRevisao
        LimpaControlesRevisao
    Case 1
        LimpaControlesRevisao
    Case 2
        AlteraRevisao
    Case 3
        mobjMsg.Abrir "Deseja EXCLUIR essa revisão do treinamento?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView1
            LimpaControlesRevisao
        End If
    Case 4
        IncluirSetor 4
    Case 5
        mobjMsg.Abrir "Deseja EXCLUIR esse setor da guia introdutório?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView2
        End If
    Case 6
        mobjMsg.Abrir "Deseja EXCLUIR esse setor da guia obrigatorio?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView3
        End If
    Case 7
        IncluirSetor 7
    Case 8
        mobjMsg.Abrir "Deseja EXCLUIR esse nível?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            ExcluirItemLV ListView4
        End If
    Case 9
        AlteraNivel
    Case 10
        LimpaControlesNivel
    Case 11
        IncluirNivel
        LimpaControlesNivel
    Case 12
        ExcluirItemLV ListView5
        montaConteudo
    Case 13
        ChamaGridTrei
        montaConteudo
    Case 14
        mobjMsg.Abrir "Deseja salvar os dados do Treinamento?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            GravarDados
            gravaLog "Código trei: " & txtCadTreinamento(0), "Nome: " & txtCadTreinamento(1), "Carga Horária: " & MaskEdBox1
            'AtivaLD
            Pesquisa = "0"
            Unload Me
            Set frmTreinamentos = Nothing
        End If
    Case 15
        mobjMsg.Abrir "Deseja sair da tela de cadastro de Treinamentos?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            Pesquisa = "0"
            Unload Me
            Set frmTreinamentos = Nothing
        End If
    End Select
End Sub

Private Sub montaConteudo()
    Dim vConteudo As String
    Dim Y As Integer, X As Integer
    Y = ListView5.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            ListView5.ListItems.Item(X).Selected = True
            If vConteudo = "" Then
                vConteudo = ListView5.SelectedItem.ListSubItems.Item(1)
            Else
                vConteudo = vConteudo & ", " & ListView5.SelectedItem.ListSubItems.Item(1)
            End If
        Next
    End If
    Me.txtCadTreinamento(2).Text = vConteudo
End Sub

Private Sub IncluirRevisao()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView1.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView1.ListItems.Item(X) = Me.txtCadTreinamento(6) Then
                ListView1.ListItems.Item(X).Selected = True
                Me.txtCadTreinamento(6) = ListView1.ListItems.Item(X)
                ListView1.SelectedItem.ListSubItems.Item(1) = DTPicker1
                ListView1.SelectedItem.ListSubItems.Item(2) = txtCadTreinamento(7)
                Y = ListView1.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView1.ListItems.Add(, , txtCadTreinamento(6))
        Y = ListView1.ListItems.Count
    Else
        Set ItemLst = ListView1.ListItems.Add(, , txtCadTreinamento(6))
        Y = ListView1.ListItems.Count
    End If
    ItemLst.SubItems(1) = DTPicker1
    ItemLst.SubItems(2) = txtCadTreinamento(7)
    txtCadTreinamento(6).Text = ""
    DTPicker1 = Date
    txtCadTreinamento(7).Text = ""
    txtCadTreinamento(6).SetFocus
    lblStatusRev = "REVISADO"
End Sub

Private Sub IncluirNivel()
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer
    'If ValidaCampo = False Then Exit Sub
    Y = ListView4.ListItems.Count
    If Y > 0 Then
        For X = 1 To Y
            If ListView4.ListItems.Item(X) = Me.txtCadTreinamento(5) Then
                ListView4.ListItems.Item(X).Selected = True
                Me.txtCadTreinamento(5) = ListView4.ListItems.Item(X)
                ListView4.SelectedItem.ListSubItems.Item(1) = txtCadTreinamento(4)
                Y = ListView4.ListItems.Count
                Exit Sub
            End If
        Next
        Set ItemLst = ListView4.ListItems.Add(, , txtCadTreinamento(5))
        Y = ListView4.ListItems.Count
    Else
        Set ItemLst = ListView4.ListItems.Add(, , txtCadTreinamento(5))
        Y = ListView4.ListItems.Count
    End If
    ItemLst.SubItems(1) = txtCadTreinamento(4)
    txtCadTreinamento(5).Text = ""
    txtCadTreinamento(4).Text = ""
    txtCadTreinamento(5).SetFocus
End Sub

Private Sub IncluirSetor(indice As Integer)
    Dim rsSetor As New ADODB.Recordset
    Dim SqlSetor As String
    Dim ItemLst As ListItem
    Dim X As Integer, Y As Integer, vCodSetor As Integer
    
    If indice = 4 Then
        SqlSetor = "Select codsetor from tbsetores where nomesetor= '" & cboCadTreinamento(5) & "'"
    Else
        SqlSetor = "Select codsetor from tbsetores where nomesetor= '" & cboCadTreinamento(4) & "'"
    End If
    rsSetor.Open SqlSetor, cnBanco, adOpenKeyset, adLockOptimistic
    
    If indice = 4 Then
        If cboCadTreinamento(5) = "Todos" Then
            ListView2.ListItems.Clear
            vCodSetor = "000"
        Else
            vCodSetor = rsSetor.Fields(0)
        End If
    Else
        If cboCadTreinamento(4) = "Todos" Then
            ListView3.ListItems.Clear
            vCodSetor = "000"
        Else
            vCodSetor = rsSetor.Fields(0)
        End If
    End If
    
    If indice = 4 Then
        Y = ListView2.ListItems.Count
        If Y > 0 Then
            For X = 1 To Y
                If Val(ListView2.ListItems.Item(X)) = vCodSetor Then
                    ListView2.ListItems.Item(X).Selected = True
                    ListView2.SelectedItem.ListSubItems.Item(1) = cboCadTreinamento(5)
                    Y = ListView2.ListItems.Count
                    Me.ListView2.Sorted = True
                    Me.ListView2.SortKey = 0
                    Me.ListView2.SortOrder = lvwAscending
                    Exit Sub
                End If
            Next
            Set ItemLst = ListView2.ListItems.Add(, , Format(vCodSetor, "000"))
            Y = ListView2.ListItems.Count
        Else
            Set ItemLst = ListView2.ListItems.Add(, , Format(vCodSetor, "000"))
            Y = ListView2.ListItems.Count
            Me.ListView2.Sorted = True
            Me.ListView2.SortKey = 0
            Me.ListView2.SortOrder = lvwAscending
        End If
        ItemLst.SubItems(1) = cboCadTreinamento(5)
        cboCadTreinamento(5).SetFocus
    Else
        Y = ListView3.ListItems.Count
        If Y > 0 Then
            For X = 1 To Y
                If Val(ListView3.ListItems.Item(X)) = vCodSetor Then
                    ListView3.ListItems.Item(X).Selected = True
                    ListView3.SelectedItem.ListSubItems.Item(1) = cboCadTreinamento(4)
                    Y = ListView3.ListItems.Count
                    Me.ListView3.Sorted = True
                    Me.ListView3.SortKey = 0
                    Me.ListView3.SortOrder = lvwAscending
                    Exit Sub
                End If
            Next
            Set ItemLst = ListView3.ListItems.Add(, , Format(vCodSetor, "000"))
            Y = ListView3.ListItems.Count
        Else
            Set ItemLst = ListView3.ListItems.Add(, , Format(vCodSetor, "000"))
            Y = ListView3.ListItems.Count
            Me.ListView3.Sorted = True
            Me.ListView3.SortKey = 0
            Me.ListView3.SortOrder = lvwAscending
        End If
        ItemLst.SubItems(1) = cboCadTreinamento(4)
        cboCadTreinamento(4).SetFocus
    End If
    rsSetor.Close
    Set rsSetor = Nothing
End Sub

Private Sub LimpaControlesRevisao()
    Dim X As Integer
    For X = 6 To 7
        txtCadTreinamento(X) = ""
    Next
    DTPicker1 = Date
End Sub

Private Sub LimpaControlesNivel()
    Dim X As Integer
    For X = 4 To 5
        txtCadTreinamento(X) = ""
    Next
End Sub

Private Sub AlteraRevisao()
    Dim Y As Integer, X As Integer
    Y = ListView1.ListItems.Count
    For X = 1 To Y
        If ListView1.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadTreinamento(6).Text = ListView1.ListItems.Item(X)
    Me.txtCadTreinamento(7).Text = ListView1.SelectedItem.ListSubItems.Item(2)
    DTPicker1 = ListView1.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub AlteraNivel()
    Dim Y As Integer, X As Integer
    Y = ListView4.ListItems.Count
    For X = 1 To Y
        If ListView4.ListItems.Item(X).Selected = True Then
            Exit For
        End If
    Next
    Me.txtCadTreinamento(5).Text = ListView4.ListItems.Item(X)
    Me.txtCadTreinamento(4).Text = ListView4.SelectedItem.ListSubItems.Item(1)
End Sub

Private Sub cmdCadastro_MouseOver(Index As Integer)
    Legenda = cmdCadastro(Index).ToolTipText
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub cmdCadastro_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Legenda = ""
    'frmMenu2.StatusBar1.Panels(3).Text = Legenda
End Sub

Private Sub Command1_Click()
    frmTipoTrei.Show 1
    cboCadTreinamento(0).Clear
    CompoeCombo cboCadTreinamento(0), "tbTipoTrei", "codigo", "nome"
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
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    
    cboCadTreinamento(0).Clear
    CompoeCombo cboCadTreinamento(0), "tbTipoTrei", "codigo", "nome"
   
    If Status = "novo" Then
        LimpaControles
    ElseIf Status = "editar" Then
        ResultPesq
        'DesbloqueiaControles
    End If
    configControles
'    cboCadTreinamento(0).Clear
'    CompoeCombo cboCadTreinamento(0), "tbTipoTrei", "codigo", "nome"
    AplicarSkin Me, Principal.Skin1
    NewColorDBGrid Me
    On Error GoTo ErrHandler
    'OrganizaForm
    Exit Sub
ErrHandler:
    mobjMsg.Abrir "ERROR: " & Err.Number & Chr(13) & "Informe ao Suporte Técnico.", , critico
End Sub

Private Sub listview_cabecalho()
    'Exemplo bem simples para criar o esboço do seu Listview
    'Cria as colunas, define o nome delase e comprimento de cada uma
    ListView1.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "Revisão", ListView1.Width / 11
    ListView1.ColumnHeaders.Add , , "Data", ListView1.Width / 8
    ListView1.ColumnHeaders.Add , , "Detalhes", ListView1.Width / 1.5
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview

    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "ID", ListView2.Width / 12
    ListView2.ColumnHeaders.Add , , "Setores", ListView2.Width / 2
    
    ListView3.ColumnHeaders.Clear
    ListView3.ColumnHeaders.Add , , "ID", ListView3.Width / 12
    ListView3.ColumnHeaders.Add , , "Setores", ListView3.Width / 2
    
    ListView4.ColumnHeaders.Clear
    ListView4.ColumnHeaders.Add , , "Nível", ListView4.Width / 12
    ListView4.ColumnHeaders.Add , , "Descrição", ListView4.Width / 2
    
    ListView5.ColumnHeaders.Clear
    ListView5.ColumnHeaders.Add , , "Código", ListView5.Width / 6
    ListView5.ColumnHeaders.Add , , "Treinamento", ListView5.Width / 2
    ListView5.ColumnHeaders.Add , , "Fase", ListView5.Width / 8
    
    ListView1.View = lvwReport 'Modo de Exibição do seu Listview
    ListView2.View = lvwReport 'Modo de Exibição do seu Listview
    ListView3.View = lvwReport 'Modo de Exibição do seu Listview
    ListView4.View = lvwReport 'Modo de Exibição do seu Listview
    ListView5.View = lvwReport 'Modo de Exibição do seu Listview

End Sub

Private Sub GravarDados()
'On Error GoTo TrataErro
    'If ValidaCampo = False Then Exit Sub
    Dim rsSalvarTreinamento As New ADODB.Recordset
    Dim SqlSalvarTreinamento As String
    Dim rsDeletar As New ADODB.Recordset
    Dim sqlDeletar As String
    Dim rsSalvar As New ADODB.Recordset
    Dim SqlSalvar As String
    
    Dim Y As Integer
    cnBanco.BeginTrans
   
    SqlSalvarTreinamento = "select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' and codTreinamento = '" & txtCadTreinamento(0) & "'"
    rsSalvarTreinamento.Open SqlSalvarTreinamento, cnBanco, adOpenKeyset, adLockOptimistic
    
    If rsSalvarTreinamento.EOF Then rsSalvarTreinamento.AddNew
    rsSalvarTreinamento.Fields(0) = Val(txtCadTreinamento(0)) 'codtreinamento
    rsSalvarTreinamento.Fields(1) = txtCadTreinamento(1) 'nometreinamento
    rsSalvarTreinamento.Fields(2) = cboCadTreinamento(0) 'tipo
    rsSalvarTreinamento.Fields(3) = cboCadTreinamento(1) 'origem
    rsSalvarTreinamento.Fields(4) = txtCadTreinamento(2) 'conteudo
    rsSalvarTreinamento.Fields(5) = txtCadTreinamento(3) 'objetivo
    If Check1.Value = 1 Then
        rsSalvarTreinamento.Fields(6) = "S"
    Else
        rsSalvarTreinamento.Fields(6) = "N" 'introdutorio
        ListView2.ListItems.Clear
    End If
    If Check3.Value = 1 Then
        rsSalvarTreinamento.Fields(7) = "S"  'aplicavel
        rsSalvarTreinamento.Fields(8) = cboCadTreinamento(2) 'tempoaplic
        rsSalvarTreinamento.Fields(9) = cboCadTreinamento(3) 'mesanoaplic
    Else
        rsSalvarTreinamento.Fields(7) = "N" 'aplicavel
        rsSalvarTreinamento.Fields(8) = "" 'tempoaplic
        rsSalvarTreinamento.Fields(9) = "" 'mesanoaplic
    End If
    rsSalvarTreinamento.Fields(10) = txtCadTreinamento(8) 'observacao
    rsSalvarTreinamento.Fields(11) = MaskEdBox1.ClipText 'cargahora
    If Check4.Value = 1 Then rsSalvarTreinamento.Fields(12) = "S" Else rsSalvarTreinamento.Fields(12) = "N" 'ativo
    
    If Check2.Value = 1 Then
        rsSalvarTreinamento.Fields(13) = "S"
    Else
        rsSalvarTreinamento.Fields(13) = "N" 'obrigatorio
        ListView3.ListItems.Clear
    End If
    
    If Check5.Value = 1 Then
        rsSalvarTreinamento.Fields(14) = "S"
    Else
        rsSalvarTreinamento.Fields(14) = "N" 'nível
        ListView4.ListItems.Clear
    End If
    If txtCadTreinamento(9) <> "" Then rsSalvarTreinamento.Fields(15) = txtCadTreinamento(9) 'Valor do treinamento
    rsSalvarTreinamento.Fields(16) = vCodcoligada 'Codigo da coligada
    rsSalvarTreinamento.Update
    
    '>>>> GRAVA TREINAMENTO PARA ADM NA TABELA TBUSUMULTIPLIC
    SqlSalvar = "Select * from tbUsuMultiplic where codusuario = 1 and codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & txtCadTreinamento(0) & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    If rsSalvar.EOF Then rsSalvar.AddNew
    rsSalvar.Fields(0) = 1 'Codigo do usuário
    rsSalvar.Fields(1) = Val(txtCadTreinamento(0)) 'Código do treinamento
    rsSalvar.Fields(2) = vCodcoligada 'Código da coligada
    If Not rsSalvar.EOF Then rsSalvar.Update
    Set rsSalvar = Nothing
    
    '>>>> GRAVA SETORES INTRODUTORIOS
    sqlDeletar = "Delete from tbTreinamentosInt where tbTreinamentosInt.codcoligada = '" & vCodcoligada & "' and tbTreinamentosInt.codtreinamento = '" & Val(txtCadTreinamento(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbTreinamentosInt where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView2.ListItems.Count
        ListView2.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadTreinamento(0).Text)
        rsSalvar.Fields(1) = ListView2.ListItems.Item(X)
        rsSalvar.Fields(2) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>> GRAVA SETORES OBRIGATORIOS
    sqlDeletar = "Delete from tbTreinamentosObr where tbTreinamentosObr.codcoligada = '" & vCodcoligada & "' and tbTreinamentosObr.codtreinamento = '" & Val(txtCadTreinamento(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbTreinamentosObr where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView3.ListItems.Count
        ListView3.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadTreinamento(0).Text)
        rsSalvar.Fields(1) = ListView3.ListItems.Item(X)
        rsSalvar.Fields(2) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>> GRAVA REVISAO DE TREINAMENTO
    sqlDeletar = "Delete from tbTreinamentosRev where tbTreinamentosRev.codcoligada = '" & vCodcoligada & "' and tbTreinamentosRev.codtreinamento = '" & Val(txtCadTreinamento(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbTreinamentosRev where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView1.ListItems.Count
        ListView1.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadTreinamento(0).Text)
        rsSalvar.Fields(1) = ListView1.ListItems.Item(X)
        rsSalvar.Fields(2) = ListView1.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(3) = ListView1.SelectedItem.ListSubItems.Item(2)
        rsSalvar.Fields(4) = vCodcoligada 'Codigo da Coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>> GRAVA NIVEL DO CURSO/TREINAMENTO SE HOUVER
    sqlDeletar = "Delete from tbTreinamentosNiv where tbTreinamentosNiv.codcoligada = '" & vCodcoligada & "' and tbTreinamentosNiv.codtreinamento = '" & Val(txtCadTreinamento(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbTreinamentosNiv where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView4.ListItems.Count
        ListView4.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadTreinamento(0).Text)
        rsSalvar.Fields(1) = ListView4.ListItems.Item(X)
        rsSalvar.Fields(2) = ListView4.SelectedItem.ListSubItems.Item(1)
        rsSalvar.Fields(3) = vCodcoligada 'Codigo da coligada
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    Set rsSalvar = Nothing
    
    '>>>> GRAVA AGRUPAMENTOS E FASES DE TREINAMENTOS
    sqlDeletar = "Delete from tbTreinamentosAgr where tbTreinamentosAgr.codcoligada = '" & vCodcoligada & "' and tbTreinamentosAgr.codigoTrei = '" & Val(txtCadTreinamento(0).Text) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
    
    SqlSalvar = "Select * from tbTreinamentosAgr where codcoligada = '" & vCodcoligada & "'"
    rsSalvar.Open SqlSalvar, cnBanco, adOpenKeyset, adLockOptimistic
    
    For X = 1 To ListView5.ListItems.Count
        ListView5.ListItems.Item(X).Selected = True
        rsSalvar.AddNew
        rsSalvar.Fields(0) = Val(txtCadTreinamento(0).Text)
        rsSalvar.Fields(1) = ListView5.ListItems.Item(X)
        rsSalvar.Fields(2) = vCodcoligada 'Codigo da coligada
        If ListView5.SelectedItem.ListSubItems.Item(2) <> "-" Then
            rsSalvar.Fields(3) = ListView5.SelectedItem.ListSubItems.Item(2) 'fase
        Else
            rsSalvar.Fields(3) = 0 'fase
        End If
    Next
    If Not rsSalvar.EOF Then rsSalvar.Update
    rsSalvar.Close
    
    If Check6.Value = 1 Then
        For X = 1 To ListView5.ListItems.Count
            ListView5.ListItems.Item(X).Selected = True
            SqlSalvar = "Update tbtreinamentos set idGrFase = '" & Format(txtCadTreinamento(0).Text, "0000") & Format(ListView5.SelectedItem.ListSubItems.Item(2), "00") & "' Where codtreinamento = '" & Val(ListView5.ListItems.Item(X)) & "'"
            rsSalvar.Open SqlSalvar, cnBanco
        Next
    End If
    '*********************************************
    'ROTINA DE REVISÃO
    'GRAVA PROGRAMAÇÃO DE TREINAMENTOS INTRODUTÓRIOS/OBRIGATÓRIOS PARA COLABORADORES
    'DOS SETORES INFORMADOS NO TREINAMENTO
    'PESQUISA O TREINAMENTO NAS MATRIZES E SE ENCONTRAR GRAVA PROGRAMAÇÃO DE TREINAMENTO
    'PARA OS COLABORADORES ADMITIDOS NAS MATRIZES ENCONTRADAS
    If lblStatusRev = "REVISADO" And Check4.Value = 1 Then
        mobjMsg.Abrir "Foram realizadas novas revisões. Deseja realizar novas programações?", YesNo, pergunta, "SGC"
        If Tp = 1 Then
            '>>>>>>>>>>> ROTINA EM TESTE <<<<<<<<<<<<<<<<<<<<<<<
            '>>>>>>>>>>> PROGRAMA TREINAMENTOS REVISADOS NÃO INTRODUTORIOS/OBRIGATORIOS
            If Check1.Value = 0 And Check2.Value = 0 Then
                'QUERY DE TREINAMENTOS NÃO INTRODUTORIOS/OBRIGATORIOS
                SqlListaAgrup = "select a.codigoTrei,a.codigoTreiGrup,b.codsetor,c.codmatriz,d.cpf,e.nomecolaborador,e.codcolaborador,e.id from tbTreinamentosAgr as a inner join tbTreinamentosInt as b on a.codcoligada = '" & vCodcoligada & "' and " & _
                                "a.codigoTrei = b.codtreinamento inner join tbmatriz as c on b.codsetor = c.codsetor inner join tbcolaboradoreshist as d on c.codmatriz = d.codmatriz and d.ativo = 'S' and d.tipo = 'colaborador'" & _
                                "inner join tbcolaboradores as e on d.cpf = e.cpf where a.codigotreigrup = '" & Val(txtCadTreinamento(0)) & "' Order by a.codigotrei,b.codsetor,c.codmatriz,e.nomecolaborador"
                rsListaAgrup.Open SqlListaAgrup, cnBanco, adOpenKeyset, adLockReadOnly

                While Not rsListaAgrup.EOF
                    excluiProgramacao rsListaAgrup.Fields(4), rsListaAgrup.Fields(3)
                    GravaTreiIntObr rsListaAgrup(4), rsListaAgrup(3), rsListaAgrup(1)
                    rsListaAgrup.MoveNext
                Wend
                rsListaAgrup.Close
                Set rsListaAgrup = Nothing
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            Else
                'QUERY DE TREINAMENTOS INTRODUTORIOS/OBRIGATORIOS
                SqlListaIntObr = "Select MAX(a.codtreinamento) as codtreinamento,MAX(a.nometreinamento) as nometreinamento,MAX(a.introdutorio) as introdutorio,MAX(b.codsetor) as codsetorInt,MAX(d.codmatriz) as matrizInt,f.cpf as cpfInt,g.nomecolaborador as nomecolaboradorInt,MAX(a.obrigatorio) as obrigatorio,MAX(c.codsetor) as codsetorInt, " & _
                "MAX(e.codmatriz) as matrizObr,h.cpf as cpfObr,i.nomecolaborador as nomecolaboradorObr from tbtreinamentos as a left join tbtreinamentosInt as b on a.codtreinamento = b.codtreinamento left join tbtreinamentosObr as c on a.codtreinamento = c.codtreinamento left join tbmatriz as d on b.codsetor = d.codsetor " & _
                "left join tbmatriz as e on c.codsetor = e.codsetor inner join tbcolaboradoreshist as f on d.codmatriz = f.codmatriz and f.ativo = 'S' inner join tbcolaboradores as g on g.ativo = 'S' and f.cpf = g.cpf left join tbcolaboradoreshist as h on e.codmatriz = h.codmatriz and h.ativo = 'S' left join tbcolaboradores as i " & _
                "on i.ativo = 'S' and h.cpf = i.cpf where a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = '" & Val(txtCadTreinamento(0)) & "' group by f.cpf,g.nomecolaborador,h.cpf,i.nomecolaborador"
                rsListaIntObr.Open SqlListaIntObr, cnBanco, adOpenKeyset, adLockReadOnly
                While Not rsListaIntObr.EOF
                    If Not IsNull(rsListaIntObr.Fields(5)) Then
                        If GeraIntr = "S" Then excluiProgramacao rsListaIntObr.Fields(5), rsListaIntObr.Fields(4)
                    End If
                    If Not IsNull(rsListaIntObr.Fields(10)) Then
                        If GeraObri = "S" Then excluiProgramacao rsListaIntObr.Fields(10), rsListaIntObr.Fields(9)
                    End If
                    If GeraIntr = "S" And Not IsNull(rsListaIntObr.Fields(5)) Then GravaTreiIntObr rsListaIntObr(5), rsListaIntObr(4), rsListaIntObr(0)
                    If GeraObri = "S" And Not IsNull(rsListaIntObr.Fields(10)) Then GravaTreiIntObr rsListaIntObr(10), rsListaIntObr(9), rsListaIntObr(0)
                    rsListaIntObr.MoveNext
                Wend
                rsListaIntObr.Close
                Set rsListaIntObr = Nothing
    
                'QUERY DE TREINAMENTOS NA MATRIZ DE CAPACITAÇÃO
                SqlListaIntObr = "select a.codtreinamento,a.nometreinamento,b.codmatriz,c.codsetor,c.codcargo,d.cpf,e.nomecolaborador from tbtreinamentos as a inner join tbmatrizcur as b on a.codcoligada = '" & vCodcoligada & "' and a.codtreinamento = b.codtreinamento " & _
                "inner join tbmatriz as c on b.codmatriz = c.codmatriz inner join tbcolaboradoreshist as d on c.codmatriz = d.codmatriz and d.ativo = 'S' inner join tbcolaboradores as e on d.cpf = e.cpf and e.ativo = 'S' " & _
                "where a.codtreinamento = '" & txtCadTreinamento(0) & "'"
                rsListaIntObr.Open SqlListaIntObr, cnBanco, adOpenKeyset, adLockReadOnly
                While Not rsListaIntObr.EOF
                    excluiProgramacao rsListaIntObr.Fields(5), rsListaIntObr.Fields(2)
                    GravaTreiIntObr rsListaIntObr(5), rsListaIntObr(2), rsListaIntObr(0)
                    rsListaIntObr.MoveNext
                Wend
                rsListaIntObr.Close
                Set rsListaIntObr = Nothing
                mobjMsg.Abrir "A revisão deste curso/treinamento gerou novas programações a serem agendadas", Ok, informacao, "SGC"
            End If
        End If
    End If
    '*********************************************
    cnBanco.CommitTrans
    rsSalvarTreinamento.Close
    Set rsSalvarTreinamento = Nothing
    Set rsSalvar = Nothing
    
    mobjMsg.Abrir "Os dados do Treinamento foram salvos com sucesso", Ok, informacao, "SGC"
    
    AtualizaListview
    Exit Sub
TrataErro:
    mobjMsg.Abrir "Ocorreu um erro, as alterções nos registros serão desfeitas!", Ok, critico, "Atenção"
    cnBanco.RollbackTrans
    Exit Sub
End Sub

Private Sub LimpaControles()
    On Error Resume Next
    Dim X As Integer
    DTPicker1 = Date
    DTPicker2 = Date
    'DesbloqueiaControles
    For X = 0 To txtCadTreinamento.Count - 1
        txtCadTreinamento(X) = ""
    Next
    Check1.Value = 0
    Check3.Value = 0
    Check4.Value = 1
    ListView1.ListItems.Clear
    txtCadTreinamento(0) = Format(GeraCodigo, "000000")
    CompoeCombo5
End Sub

Private Sub CompoeCombo5()
    cboCadTreinamento(5).Clear
    cboCadTreinamento(5).Text = "Todos"
    cboCadTreinamento(5).AddItem "Todos"
    CompoeCombo cboCadTreinamento(5), "tbsetores", "codsetor", "nomesetor"

    cboCadTreinamento(4).Clear
    cboCadTreinamento(4).Text = "Todos"
    cboCadTreinamento(4).AddItem "Todos"
    CompoeCombo cboCadTreinamento(4), "tbsetores", "codsetor", "nomesetor"
End Sub
Private Sub CompoeControles()
    Dim X As Integer
    CompoeCombo5
    txtCadTreinamento(0).Text = Format(rsTreinamentos.Fields(0), "000000") 'codtreinamento
    txtCadTreinamento(1).Text = rsTreinamentos.Fields(1) 'nometreinamento
    txtCadTreinamento(2).Text = rsTreinamentos.Fields(4) 'conteudo
    txtCadTreinamento(3).Text = rsTreinamentos.Fields(5) 'objetivo
    txtCadTreinamento(8).Text = rsTreinamentos.Fields(10) 'observação
    cboCadTreinamento(0).Text = rsTreinamentos.Fields(2) 'tipo
    cboCadTreinamento(1).Text = rsTreinamentos.Fields(3) 'origem
    cboCadTreinamento(2).Text = rsTreinamentos.Fields(8) 'tempoaplic
    cboCadTreinamento(3).Text = rsTreinamentos.Fields(9) 'mesanoaplic
    If rsTreinamentos.Fields(6) = "S" Then
        Check1.Value = 1 'introdutorio
    Else
        Check1.Value = 0 'introdutorio
    End If
    If rsTreinamentos.Fields(7) = "S" Then Check3.Value = 1 Else Check3.Value = 0 'aplicavel
    If rsTreinamentos.Fields(12) = "S" Then Check4.Value = 1 Else Check4.Value = 0 'ativo
    If Check4.Value = 0 Then
        Frame6.Enabled = True
        Check4.Enabled = True
    End If
    
    If rsTreinamentos.Fields(13) = "S" Then
        Check2.Value = 1 'obrigatorio
    Else
        Check2.Value = 0 'obrigatorio
    End If
    If rsTreinamentos.Fields(14) = "S" Then
        Check5.Value = 1 'nível
    Else
        Check5.Value = 0 'nível
    End If
    If Not IsNull(rsTreinamentos.Fields(15)) Then txtCadTreinamento(9) = Format(rsTreinamentos.Fields(15), "#,##00.00;(#,##0.00)") 'Valor do treinamento
    MaskEdBox1.PromptInclude = False
    MaskEdBox1 = rsTreinamentos.Fields(11) 'cargahora
    MaskEdBox1.PromptInclude = True
End Sub

Private Sub Compoe_Listview()
    'PREENCHE O LISTVIEW DE REVISAO
    Dim ItemLst As ListItem
    Dim X As Integer
    X = 0
    While Not rsRevisao.EOF
        Set ItemLst = ListView1.ListItems.Add(, , rsRevisao.Fields(1))
        ItemLst.SubItems(1) = "" & rsRevisao.Fields(2)
        ItemLst.SubItems(2) = "" & rsRevisao.Fields(3)
        rsRevisao.MoveNext
        X = X + 1
    Wend
    'PREENCHE O LISTVIEW DE SETORES INTRODUTORIOS
    While Not rsSet.EOF
        Set ItemLst = ListView2.ListItems.Add(, , Format(rsSet.Fields(1), "000"))
        If rsSet.Fields(1) <> 0 Then
            ItemLst.SubItems(1) = "" & rsSet.Fields(3)
        Else
            ItemLst.SubItems(1) = "Todos"
        End If
        rsSet.MoveNext
        X = X + 1
    Wend
    'PREENCHE O LISTVIEW DE SETORES OBRIGATORIOS
    While Not rsSetObr.EOF
        Set ItemLst = ListView3.ListItems.Add(, , Format(rsSetObr.Fields(1), "000"))
        If rsSetObr.Fields(1) <> 0 Then
            ItemLst.SubItems(1) = "" & rsSetObr.Fields(3)
        Else
            ItemLst.SubItems(1) = "Todos"
        End If
        rsSetObr.MoveNext
        X = X + 1
    Wend
    'PREENCHE O LISTVIEW DE NÍVEL
    While Not rsNivel.EOF
        Set ItemLst = ListView4.ListItems.Add(, , rsNivel.Fields(1))
        ItemLst.SubItems(1) = "" & rsNivel.Fields(2)
        rsNivel.MoveNext
        X = X + 1
    Wend
    
    'PREENCHE O LISTVIEW DE AGRUPAMENTO
    While Not rsAgrup.EOF
        Set ItemLst = ListView5.ListItems.Add(, , Format(rsAgrup.Fields(0), "000000"))
        ItemLst.SubItems(1) = "" & rsAgrup.Fields(1)
        If IsNull(rsAgrup.Fields(2)) Or rsAgrup.Fields(2) = 0 Then
            ItemLst.SubItems(2) = "-"
            Check6.Value = 0
        Else
            ItemLst.SubItems(2) = rsAgrup.Fields(2)
            Check6.Value = 1
        End If
        rsAgrup.MoveNext
        X = X + 1
    Wend
    
    Me.ListView1.Sorted = True
    Me.ListView1.SortKey = 0
    Me.ListView1.SortOrder = lvwDescending

    Me.ListView2.Sorted = True
    Me.ListView2.SortKey = 0
    Me.ListView2.SortOrder = lvwAscending

    Me.ListView3.Sorted = True
    Me.ListView3.SortKey = 0
    Me.ListView3.SortOrder = lvwAscending

    Me.ListView4.Sorted = True
    Me.ListView4.SortKey = 0
    Me.ListView4.SortOrder = lvwAscending
    
    Me.ListView5.Sorted = True
    Me.ListView5.SortKey = 0
    Me.ListView5.SortOrder = lvwAscending
    
End Sub

Private Function ValidaCampo()
    ValidaCampo = False
    If txtCadTreinamento(0).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadTreinamento(0).Tag, Ok, critico, "Atenção"
        Me.txtCadTreinamento(0).SetFocus
        Exit Function
    End If
    If txtCadTreinamento(1).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadTreinamento(1).Tag, Ok, critico, "Atenção"
        Me.txtCadTreinamento(1).SetFocus
        Exit Function
    End If
    If txtCadTreinamento(2).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadTreinamento(2).Tag, Ok, critico, "Atenção"
        Me.txtCadTreinamento(2).SetFocus
        Exit Function
    End If
    If txtCadTreinamento(4).Text = "" Then
        mobjMsg.Abrir "Favor informar o campo " & Me.txtCadTreinamento(4).Tag, Ok, critico, "Atenção"
        Me.txtCadTreinamento(4).SetFocus
        Exit Function
    End If
    ValidaCampo = True
End Function

Private Sub BloqueiaControles()
    For X = 1 To txtCadTreinamento.Count - 1
        txtCadTreinamento(X).Enabled = False
    Next
End Sub

Private Sub DesbloqueiaControles()
    For X = 1 To txtCadTreinamento.Count - 1
        txtCadTreinamento(X).Enabled = True
    Next
End Sub

Private Function GeraCodigo()
    Dim rsGeraCodigo As New ADODB.Recordset
    Dim SqlGera
    AbrirTreinamento
    SqlGera = "Select top 1 * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' order by codTreinamento Desc"
    rsGeraCodigo.Open SqlGera, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTreinamentos.RecordCount > 0 Then
        GeraCodigo = rsGeraCodigo.Fields(0) + 1
    Else
        GeraCodigo = 1
    End If
    txtCadTreinamento(0) = GeraCodigo
    rsGeraCodigo.Close
    Set rsGeraCodigo = Nothing
    FecharTreinamentos
End Function

Private Sub AbrirTreinamento()
    SqlTreinamentos = "Select * from tbTreinamentos where codcoligada = '" & vCodcoligada & "' Order by codTreinamento"
    rsTreinamentos.Open SqlTreinamentos, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub AbrirRevisao()
    SqlRevisao = "Select * from tbTreinamentosRev where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(txtCadTreinamento(0)) & "'Order by codTreinamento,revisao"
    rsRevisao.Open SqlRevisao, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub AbrirNivel()
    SqlNivel = "Select * from tbTreinamentosNiv where codcoligada = '" & vCodcoligada & "' and codtreinamento = '" & Val(txtCadTreinamento(0)) & "'Order by codTreinamento,codnivel"
    rsNivel.Open SqlNivel, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub AbrirSetor()
    SqlSet = "Select tbTreinamentosInt.*,tbsetores.nomesetor from tbTreinamentosInt left join tbsetores on tbTreinamentosInt.codsetor = tbSetores.codsetor where tbTreinamentosInt.codcoligada = '" & vCodcoligada & "' and tbTreinamentosInt.codtreinamento = '" & Val(txtCadTreinamento(0)) & "' Order by tbTreinamentosInt.codTreinamento,tbTreinamentosInt.codSetor"
    rsSet.Open SqlSet, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub AbrirSetorObr()
    SqlSetObr = "Select tbTreinamentosObr.*,tbsetores.nomesetor from tbTreinamentosObr left join tbsetores on tbTreinamentosObr.codsetor = tbSetores.codsetor where tbTreinamentosObr.codcoligada = '" & vCodcoligada & "' and tbTreinamentosObr.codtreinamento = '" & Val(txtCadTreinamento(0)) & "' Order by tbTreinamentosObr.codTreinamento,tbTreinamentosObr.codSetor"
    rsSetObr.Open SqlSetObr, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub AbrirAgrup()
    SqlAgrup = "select a.codigoTreiGrup,b.nometreinamento,a.fase from tbTreinamentosAgr as a inner join tbTreinamentos as b on a.codcoligada = '" & vCodcoligada & "' and a.codigoTreiGrup = b.codtreinamento where a.codigotrei = '" & Val(txtCadTreinamento(0)) & "'"
    rsAgrup.Open SqlAgrup, cnBanco, adOpenKeyset, adLockOptimistic
End Sub

Private Sub FecharTreinamentos()
    rsTreinamentos.Close
    Set rsTreinamentos = Nothing
End Sub

Private Sub FecharRevisao()
    rsRevisao.Close
    Set rsRevisao = Nothing
End Sub

Private Sub FecharNivel()
    rsNivel.Close
    Set rsNivel = Nothing
End Sub

Private Sub FecharSetor()
    rsSet.Close
    Set rsSet = Nothing
End Sub

Private Sub FecharSetorObr()
    rsSetObr.Close
    Set rsSetObr = Nothing
End Sub

Private Sub FecharAgrup()
    rsAgrup.Close
    Set rsAgrup = Nothing
End Sub

Private Sub ResultPesq()
    SqlTreinamentos = "Select * from tbTreinamentos Where tbTreinamentos.codcoligada = '" & vCodcoligada & "' and tbTreinamentos.codTreinamento= '" & Val(varGlobal) & "' order by tbTreinamentos.codTreinamento"
    rsTreinamentos.Open SqlTreinamentos, cnBanco, adOpenKeyset, adLockReadOnly
    If rsTreinamentos.RecordCount > 0 Then
        CompoeControles
        AbrirRevisao
        AbrirSetor
        AbrirSetorObr
        AbrirNivel
        AbrirAgrup
        Compoe_Listview
        FecharRevisao
        FecharSetor
        FecharSetorObr
        FecharNivel
        FecharAgrup
    Else
        mobjMsg.Abrir "Treinamento não encontrado", Ok, critico, "Atenção"
    End If
    rsTreinamentos.Close
    Set rsTreinamentos = Nothing
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
    If Status = "novo" Then
        Set ItemLst = MeuLV.ListView1.ListItems.Add(, , Format(txtCadTreinamento(0), "000000"))
        ItemLst.SubItems(1) = txtCadTreinamento(1).Text
        ItemLst.SubItems(2) = cboCadTreinamento(1).Text
        If Check1.Value = 1 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        Else
            ItemLst.SubItems(3) = "" 'Introdutorio
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        End If
        
        If Check1.Value = 1 Then
            ItemLst.SubItems(3) = ""
            ItemLst.ListSubItems.Item(3).ReportIcon = "OK"
        Else
            ItemLst.SubItems(3) = "" 'Introdutorio
            ItemLst.ListSubItems.Item(3).ReportIcon = "EXC"
        End If
        
        ItemLst.SubItems(5) = cboCadTreinamento(0).Text
        'ItemLst.ListSubItems(3).Bold = True
        'ItemLst.ListSubItems(4).Bold = True
    Else
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(1) = txtCadTreinamento(1).Text
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(2) = cboCadTreinamento(1).Text
        If Check1.Value = 1 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "OK"
            'MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = cboCadTreinamento(5).Text
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).ReportIcon = "EXC"
            'MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
        End If
       
        If Check2.Value = 1 Then
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "OK"
        Else
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4) = ""
            MeuLV.ListView1.SelectedItem.ListSubItems.Item(4).ReportIcon = "EXC"
        End If
       
       
        MeuLV.ListView1.SelectedItem.ListSubItems.Item(5) = cboCadTreinamento(0).Text
        'MeuLV.ListView1.SelectedItem.ListSubItems.Item(3).Bold = True
    End If

    Exit Sub
Err:
    mobjMsg.Abrir "Não foi possível realizar as alterações", Ok, critico, "Atenção"
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    If vEdi <> "N" Then
        AlteraRevisao
    End If
End Sub

Private Sub ListView4_DblClick()
    If vEdi <> "N" Then
        AlteraNivel
    End If
End Sub

Private Sub txtCadTreinamento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo Error
    Select Case Index
    End Select
Error:
    Exit Sub
End Sub

Private Sub configControles()
    If vInc = "N" Then
        cmdCadastro(11).UseGreyscale = True
        cmdCadastro(11).DragMode = 1
        cmdCadastro(11).SpecialEffect = cbEngraved
        
        cmdCadastro(7).UseGreyscale = True
        cmdCadastro(7).DragMode = 1
        cmdCadastro(7).SpecialEffect = cbEngraved
        
        cmdCadastro(4).UseGreyscale = True
        cmdCadastro(4).DragMode = 1
        cmdCadastro(4).SpecialEffect = cbEngraved
    
        cmdCadastro(0).UseGreyscale = True
        cmdCadastro(0).DragMode = 1
        cmdCadastro(0).SpecialEffect = cbEngraved
    
        cmdCadastro(1).UseGreyscale = True
        cmdCadastro(1).DragMode = 1
        cmdCadastro(1).SpecialEffect = cbEngraved
    End If
    If vEdi = "N" Then
        cmdCadastro(2).UseGreyscale = True
        cmdCadastro(2).DragMode = 1
        cmdCadastro(2).SpecialEffect = cbEngraved
        
        cmdCadastro(9).UseGreyscale = True
        cmdCadastro(9).DragMode = 1
        cmdCadastro(9).SpecialEffect = cbEngraved
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
        
        cmdCadastro(8).UseGreyscale = True
        cmdCadastro(8).DragMode = 1
        cmdCadastro(8).SpecialEffect = cbEngraved
    End If
End Sub

Private Sub txtCadTreinamento_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 9
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    Case 5
        'aceitar somente números e "Back Space", "Enter", "virgula"
        If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
            KeyAscii = 0
        End If
    End Select
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
    sqlDeletar = "Delete from tbPendentesCur where tbPendentesCur.codcoligada = '" & vCodcoligada & "' and tbPendentesCur.cpf = '" & vCPF & "' and status = 'Pendente' and codmatriz = '" & vMatriz & "' and codtreinamento = '" & Val(txtCadTreinamento(0)) & "' or " & _
                                                  "tbPendentesCur.codcoligada = '" & vCodcoligada & "' and tbPendentesCur.cpf = '" & vCPF & "' and status = 'Agendado' and codmatriz = '" & vMatriz & "' and codtreinamento = '" & Val(txtCadTreinamento(0)) & "'"
    rsDeletar.Open sqlDeletar, cnBanco
End Sub

'Grava treinamentos introdutorios e obrigatorios
Private Sub GravaTreiIntObr(vCPF As String, vMatriz As Integer, vCodTreinamento As Integer)
    'On Error Resume Next
    Dim rsGravaTreiInt As New ADODB.Recordset
    Dim SqlGravaTreiInt As String
    Dim contaID As Integer
    
    SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,fase from tbPendentesCur"
    rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
    If Not rsGravaTreiInt.EOF Then
        rsGravaTreiInt.MoveLast
        contaID = rsGravaTreiInt.Fields(5) + 1
    Else
        contaID = 1
    End If
    rsGravaTreiInt.Close
    Set rsGravaTreiInt = Nothing
    
    SqlGravaTreiInt = "Select cpf,codmatriz,codtreinamento,codprogramacao,ativo,id,status,tipoprogramacao,codnivel,codcoligada,fase from tbPendentesCur where codcoligada = '" & vCodcoligada & "'"
    rsGravaTreiInt.Open SqlGravaTreiInt, cnBanco, adOpenKeyset, adLockOptimistic
            
    rsGravaTreiInt.AddNew
    rsGravaTreiInt.Fields(0) = vCPF
    rsGravaTreiInt.Fields(1) = vMatriz
    rsGravaTreiInt.Fields(2) = vCodTreinamento
    rsGravaTreiInt.Fields(4) = "S"
    rsGravaTreiInt.Fields(5) = contaID
    rsGravaTreiInt.Fields(6) = "Pendente"
    rsGravaTreiInt.Fields(7) = 0
    rsGravaTreiInt.Fields(8) = 0
    rsGravaTreiInt.Fields(9) = vCodcoligada 'Codigo da coligada
    contaID = contaID + 1
        
    rsGravaTreiInt.Update
    rsGravaTreiInt.Close
    Set rsGravaTreiInt = Nothing
End Sub
